"""Reusable workflows shared by the CLI and web app."""

from __future__ import annotations

import re
from dataclasses import dataclass, replace
from email.utils import parseaddr
from pathlib import Path
from typing import Callable, Iterable

from googleapiclient.errors import HttpError

from .analyzer import EmailAnalyzer
from .config import Settings
from .database import Database, RulesDatabase
from .database import FILTER_ACTION_ALWAYS_PROCESS, FILTER_ACTION_PRECLASSIFY
from .digest import DIGEST_PRIORITIES, EmailDigestGenerator
from .gmail_client import GMAIL_MODIFY_SCOPE, GMAIL_READONLY_SCOPE, GmailClient
from .llm_client import LLMClient
from .models import EmailMessage
from .report import ReportGenerator


ProgressCallback = Callable[[int, str, str], None]


@dataclass(frozen=True)
class LoadMailResult:
    fetched_count: int
    processable_count: int
    analyzed_count: int
    report_path: Path | None


def create_mail_database(settings: Settings) -> Database:
    database = Database(
        settings.db_path,
        max_size_bytes=int(settings.db_max_size_mb * 1024 * 1024),
        backup_dir=settings.db_backup_dir,
        store_email_body=settings.store_email_body,
    )
    database.create_all()
    return database


def create_rules_database(settings: Settings) -> RulesDatabase:
    rules_database = RulesDatabase(settings.rules_db_path)
    rules_database.create_all()
    return rules_database


def fetch_processable_messages(
    settings: Settings,
    progress: ProgressCallback | None = None,
) -> tuple[list[EmailMessage], int]:
    _notify(progress, 10, "fetch", "Gmail に接続しています")
    database = create_mail_database(settings)
    rules_database = create_rules_database(settings)
    client = GmailClient.from_oauth(
        settings,
        scopes=[GMAIL_READONLY_SCOPE, GMAIL_MODIFY_SCOPE],
    )
    account_email = client.get_profile_email()
    database.save_account_email(account_email)
    try:
        messages = client.fetch_recent_messages(
            lookback_days=settings.lookback_days,
            max_results=settings.max_emails_per_run,
            stop_when_message_id_seen=database.has_email,
        )
        sent_messages = client.fetch_recent_sent_messages(
            lookback_days=settings.lookback_days,
            max_results=settings.max_emails_per_run,
            stop_when_message_id_seen=database.has_email,
        )
    except HttpError:
        raise

    all_messages = [*messages, *sent_messages]
    _notify(progress, 30, "fetch", f"{len(all_messages)} 件のメールを保存しています")
    database.save_emails(all_messages)
    starred_message_ids = _apply_starred_priority(database, messages)
    processable_messages: list[EmailMessage] = []
    for message in database.filter_processable(messages):
        if message.gmail_message_id in starred_message_ids:
            continue
        decision = rules_database.get_filter_decision(message)
        instruction_priority = rules_database.get_highest_llm_instruction_priority(message)
        if decision is None:
            processable_messages.append(message)
            continue
        if (
            instruction_priority is not None
            and instruction_priority < decision["priority"]
        ):
            processable_messages.append(message)
            continue

        database.save_filter_decision(
            message.gmail_message_id,
            action=decision["action"],
            matched_rule_id=decision["matched_rule_id"],
            matched_rule_name=decision["matched_rule_name"],
            reason=decision["reason"],
            rule_snapshot=decision["rule_snapshot"],
        )
        if decision["action"] == FILTER_ACTION_PRECLASSIFY:
            database.apply_preclassified_analysis(
                message.gmail_message_id,
                decision["preset_analysis"],
            )
        elif decision["action"] == FILTER_ACTION_ALWAYS_PROCESS:
            processable_messages.append(message)

    return processable_messages, len(all_messages)


def analyze_saved_messages(
    settings: Settings,
    limit: int | None = None,
    progress: ProgressCallback | None = None,
) -> int:
    database = create_mail_database(settings)
    rules_database = create_rules_database(settings)
    limit = limit or settings.max_emails_per_run
    messages = database.list_unanalyzed_emails(limit=limit)
    messages = rules_database.filter_processable(messages)
    if not messages:
        _notify(progress, 70, "analyze", "解析が必要なメールはありません")
        return 0

    analyzer = EmailAnalyzer(llm_client=LLMClient.from_settings(settings))
    groups = group_duplicate_messages_for_analysis(messages)
    total = len(groups)
    analyzed_count = 0
    for index, group in enumerate(groups, start=1):
        _notify(progress, 45 + int(40 * (index - 1) / total), "analyze", f"{index}/{total} 件目を解析しています")
        merged_message = merge_message_group_for_analysis(group)
        instructions = rules_database.get_llm_instructions_for_email(merged_message)
        analysis = analyzer.analyze(
            merged_message,
            additional_instructions=instructions,
        )
        for message in group:
            database.save_analysis(
                message.gmail_message_id,
                analysis,
                settings.llm_model,
            )
        analyzed_count += len(group)
    _notify(progress, 85, "analyze", f"{analyzed_count} 件の解析が完了しました")
    return analyzed_count


def _apply_starred_priority(
    database: Database,
    messages: list[EmailMessage],
) -> set[str]:
    starred_message_ids: set[str] = set()
    for message in messages:
        if "STARRED" not in set(message.label_ids):
            continue
        database.apply_preclassified_analysis(
            message.gmail_message_id,
            {
                "summary_ja": "Gmailでスター付きだったためHighにしました。",
                "category": "gmail_starred",
                "priority": "high",
                "requires_reply": False,
                "suggested_action_ja": "Gmailのスターを確認してください。",
            },
            llm_model="gmail:starred",
        )
        starred_message_ids.add(message.gmail_message_id)
    return starred_message_ids


def star_high_priority_messages(
    settings: Settings,
    progress: ProgressCallback | None = None,
) -> int:
    database = create_mail_database(settings)
    messages = database.list_high_priority_emails_without_star()
    if not messages:
        return 0
    client = GmailClient.from_oauth(
        settings,
        scopes=[GMAIL_READONLY_SCOPE, GMAIL_MODIFY_SCOPE],
    )
    starred_count = 0
    for message in messages:
        try:
            client.star_message(message.gmail_message_id)
            starred_count += 1
        except HttpError:
            continue
    if starred_count:
        _notify(progress, 84, "star", f"{starred_count} 件のHighメールにスターを付けました")
    return starred_count


def generate_saved_digests(
    settings: Settings,
    limit: int | None = None,
    progress: ProgressCallback | None = None,
) -> int:
    database = create_mail_database(settings)
    limit = limit or settings.max_emails_per_run
    messages = database.list_emails_needing_digest(
        priorities=DIGEST_PRIORITIES,
        limit=limit,
    )
    messages.extend(database.list_sent_emails_needing_digest(limit=limit))
    if not messages:
        _notify(progress, 86, "digest", "翻訳と要約が必要なメールはありません")
        return 0

    generator = EmailDigestGenerator(llm_client=LLMClient.from_settings(settings))
    total = len(messages)
    completed = 0
    for index, message in enumerate(messages, start=1):
        _notify(progress, 86 + int(3 * (index - 1) / total), "digest", f"{index}/{total} 件目の翻訳と要約を生成しています")
        try:
            generator.generate_and_save(
                message,
                mail_database=database,
                llm_model=settings.llm_model,
            )
            completed += 1
        except ValueError:
            continue
    _notify(progress, 89, "digest", f"{completed} 件の翻訳と要約が完了しました")
    return completed


def generate_report(
    settings: Settings,
    open_browser: bool = False,
    progress: ProgressCallback | None = None,
) -> Path | None:
    _notify(progress, 90, "report", "HTML レポートを生成しています")
    database = create_mail_database(settings)
    return ReportGenerator(
        database=database,
        report_dir=settings.report_dir,
        timezone=settings.timezone,
        excluded_sender_addresses=settings.account_emails,
    ).generate_all(open_browser=open_browser)


def load_today_mail(
    settings: Settings,
    progress: ProgressCallback | None = None,
) -> LoadMailResult:
    processable_messages, fetched_count = fetch_processable_messages(settings, progress)
    analyzed_count = analyze_saved_messages(
        settings,
        limit=max(len(processable_messages), 1),
        progress=progress,
    )
    star_high_priority_messages(settings, progress=progress)
    generate_saved_digests(
        settings,
        limit=settings.max_emails_per_run,
        progress=progress,
    )
    report_path = generate_report(settings, open_browser=False, progress=progress)
    _notify(progress, 100, "done", "完了しました")
    return LoadMailResult(
        fetched_count=fetched_count,
        processable_count=len(processable_messages),
        analyzed_count=analyzed_count,
        report_path=report_path,
    )


def _notify(
    progress: ProgressCallback | None,
    percent: int,
    stage: str,
    message: str,
) -> None:
    if progress is not None:
        progress(percent, stage, message)


def group_duplicate_messages_for_analysis(
    messages: list[EmailMessage],
) -> list[tuple[EmailMessage, ...]]:
    groups: dict[tuple[str, str, str], list[EmailMessage]] = {}
    order: list[tuple[str, str, str]] = []
    for message in messages:
        key = _duplicate_message_key(message)
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(message)
    return [tuple(groups[key]) for key in order]


def merge_message_group_for_analysis(
    messages: tuple[EmailMessage, ...],
) -> EmailMessage:
    if not messages:
        raise ValueError("messages must not be empty")
    primary = messages[0]
    if len(messages) == 1:
        return primary
    return replace(
        primary,
        recipients=_unique_values(
            recipient
            for message in messages
            for recipient in message.recipients
        ),
        cc=_unique_values(
            recipient
            for message in messages
            for recipient in message.cc
        ),
    )


def _duplicate_message_key(message: EmailMessage) -> tuple[str, str, str]:
    body_key = message.body or message.body_html or message.snippet
    return (
        _normalize_address(message.sender),
        _normalize_text(message.subject),
        _normalize_text(body_key),
    )


def _normalize_address(value: str) -> str:
    return (parseaddr(value)[1] or value).strip().lower()


def _normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip().lower()


def _unique_values(values: Iterable[str]) -> tuple[str, ...]:
    unique: list[str] = []
    seen: set[str] = set()
    for value in values:
        key = value.strip().lower()
        if not key or key in seen:
            continue
        seen.add(key)
        unique.append(value)
    return tuple(unique)
