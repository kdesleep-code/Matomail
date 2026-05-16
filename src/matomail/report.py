"""HTML report generation."""

from __future__ import annotations

import calendar
import json
import re
import webbrowser
from dataclasses import dataclass, replace
from datetime import UTC, date, datetime, timedelta, timezone as fixed_timezone
from email.utils import parseaddr
from pathlib import Path
from typing import Iterable
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from jinja2 import Environment, FileSystemLoader, select_autoescape
from sqlalchemy import select

from .models import AttachmentInfo
from .database import Database, EmailAnalysisRecord, EmailDigestRecord, EmailRecord, FilterDecisionRecord
from .database import FILTER_ACTION_IGNORE, FILTER_ACTION_SKIP_ANALYSIS


PRIORITY_RANK = {"top": 4, "high": 3, "medium": 2, "low": 1}


@dataclass(frozen=True)
class ReportAnalysis:
    priority: str
    summary_ja: str
    requires_reply: bool
    category: str
    meeting_candidates: list


@dataclass(frozen=True)
class ReportDigest:
    summary_ja: str
    translation_ja: str


@dataclass(frozen=True)
class ReportEmail:
    gmail_message_id: str
    gmail_thread_id: str
    sender: str
    recipients: tuple[str, ...]
    cc: tuple[str, ...]
    subject: str
    loaded_at: datetime
    loaded_at_display: str
    received_at: datetime | None
    received_at_display: str
    report_date: str
    snippet: str
    body: str
    body_html: str
    has_attachments: bool
    attachments: tuple[AttachmentInfo, ...]
    is_sent: bool
    sender_color: str
    message_href: str
    analysis: ReportAnalysis
    digest: ReportDigest | None


@dataclass(frozen=True)
class ReportThread:
    key: str
    message_href: str
    gmail_href: str
    subject: str
    sender: str
    recipients: tuple[str, ...]
    cc: tuple[str, ...]
    loaded_at: datetime
    loaded_at_display: str
    received_at: datetime | None
    received_at_display: str
    report_date: str
    snippet: str
    search_text: str
    analysis: ReportAnalysis
    messages: tuple[ReportEmail, ...]
    current_message_count: int

    @property
    def message_count(self) -> int:
        return len(self.messages)

    @property
    def primary_message(self) -> ReportEmail:
        for message in self.messages:
            if not message.is_sent and message.report_date == self.report_date:
                return message
        return self.messages[0]

    @property
    def attachment_summary(self) -> str:
        return _attachment_summary(
            attachment
            for message in self.messages
            for attachment in message.attachments
        )


@dataclass(frozen=True)
class CalendarDay:
    day: int
    href: str
    has_emails: bool
    is_current: bool


@dataclass(frozen=True)
class CalendarWeek:
    days: tuple[CalendarDay | None, ...]


class ReportGenerator:
    """Generates local HTML reports from saved emails."""

    def __init__(
        self,
        database: Database,
        report_dir: Path | str,
        timezone: str = "Asia/Tokyo",
        excluded_sender_addresses: tuple[str, ...] = (),
    ) -> None:
        self.database = database
        self.report_dir = Path(report_dir)
        self.timezone = _load_timezone(timezone)
        self.excluded_sender_addresses = {
            _normalize_email_address(address) for address in excluded_sender_addresses
        }
        template_dir = Path(__file__).parent / "templates"
        self.environment = Environment(
            loader=FileSystemLoader(template_dir),
            autoescape=select_autoescape(("html", "xml", "j2")),
        )

    def generate_all(self, open_browser: bool = False) -> Path | None:
        emails = self._list_report_emails()
        list_emails = [email for email in emails if not email.is_sent]
        grouped = _group_by_date(list_emails)
        if not grouped:
            return None

        dates = sorted(grouped)
        for index, date_key in enumerate(dates):
            self._write_daily_report(
                date_key=date_key,
                emails=grouped[date_key],
                all_emails=emails,
                available_dates=dates,
                previous_date=dates[index - 1] if index > 0 else None,
                next_date=dates[index + 1] if index < len(dates) - 1 else None,
            )

        latest_path = self.report_dir / dates[-1] / "index.html"
        self._write_root_index(dates[-1])
        if open_browser:
            webbrowser.open(latest_path.resolve().as_uri())
        return latest_path

    def _write_daily_report(
        self,
        *,
        date_key: str,
        emails: list[ReportEmail],
        all_emails: list[ReportEmail],
        available_dates: list[str],
        previous_date: str | None,
        next_date: str | None,
    ) -> Path:
        daily_dir = self.report_dir / date_key
        messages_dir = daily_dir / "messages"
        messages_dir.mkdir(parents=True, exist_ok=True)
        threads = _build_report_threads(emails, all_emails, date_key)

        template = self.environment.get_template("report.html.j2")
        report_day = date.fromisoformat(date_key)
        context = {
            "generated_at": datetime.now(self.timezone).strftime("%Y-%m-%d %H:%M"),
            "report_date": date_key,
            "emails": threads,
            "email_count": sum(thread.current_message_count for thread in threads),
            "thread_count": len(threads),
            "previous_href": f"../{previous_date}/index.html" if previous_date else "",
            "next_href": f"../{next_date}/index.html" if next_date else "",
            "calendar_weeks": _build_calendar(report_day, available_dates),
            "calendar_month": report_day.strftime("%Y-%m"),
            "priority_rank": PRIORITY_RANK,
            "page": "list",
        }
        index_path = daily_dir / "index.html"
        index_path.write_text(template.render(context), encoding="utf-8")

        for index, thread in enumerate(threads):
            message_path = messages_dir / Path(thread.message_href).name
            previous_thread = threads[index - 1] if index > 0 else None
            next_thread = threads[index + 1] if index < len(threads) - 1 else None
            message_context = {
                **context,
                "page": "message",
                "thread": thread,
                "email": thread.primary_message,
                "back_href": "../index.html",
                "previous_href": (
                    Path(previous_thread.message_href).name
                    if previous_thread is not None
                    else ""
                ),
                "next_href": (
                    Path(next_thread.message_href).name
                    if next_thread is not None
                    else ""
                ),
            }
            message_path.write_text(template.render(message_context), encoding="utf-8")

        return index_path

    def _write_root_index(self, latest_date: str) -> None:
        self.report_dir.mkdir(parents=True, exist_ok=True)
        content = (
            '<!doctype html><html lang="ja"><head>'
            '<meta charset="utf-8">'
            f'<meta http-equiv="refresh" content="0; url={latest_date}/index.html">'
            "<title>Matomail Report</title></head><body>"
            f'<a href="{latest_date}/index.html">最新のメール一覧を開く</a>'
            "</body></html>"
        )
        (self.report_dir / "index.html").write_text(content, encoding="utf-8")

    def _list_report_emails(self) -> list[ReportEmail]:
        with self.database.session_factory() as session:
            records = session.scalars(
                select(EmailRecord).order_by(
                    EmailRecord.created_at.desc(),
                    EmailRecord.received_at.desc(),
                    EmailRecord.id.desc(),
                )
            ).all()
            emails = []
            excluded_sender_addresses = (
                self.excluded_sender_addresses | set(self.database.list_account_emails())
            )
            for record in records:
                is_sent = _is_sent_record(record, excluded_sender_addresses)
                decision = session.scalar(
                    select(FilterDecisionRecord)
                    .where(FilterDecisionRecord.email_id == record.id)
                    .order_by(FilterDecisionRecord.decided_at.desc(), FilterDecisionRecord.id.desc())
                    .limit(1)
                )
                if decision and decision.action in {
                    FILTER_ACTION_IGNORE,
                    FILTER_ACTION_SKIP_ANALYSIS,
                }:
                    continue
                analysis = session.scalar(
                    select(EmailAnalysisRecord)
                    .where(EmailAnalysisRecord.email_id == record.id)
                    .order_by(
                        EmailAnalysisRecord.created_at.desc(),
                        EmailAnalysisRecord.id.desc(),
                    )
                    .limit(1)
                )
                digest = session.scalar(
                    select(EmailDigestRecord)
                    .where(EmailDigestRecord.email_id == record.id)
                    .order_by(
                        EmailDigestRecord.created_at.desc(),
                        EmailDigestRecord.id.desc(),
                    )
                    .limit(1)
                )
                emails.append(self._record_to_report_email(record, analysis, digest, is_sent))
            return _merge_duplicate_report_emails(emails)

    def _record_to_report_email(
        self,
        record: EmailRecord,
        analysis: EmailAnalysisRecord | None,
        digest: EmailDigestRecord | None = None,
        is_sent: bool = False,
    ) -> ReportEmail:
        loaded_at = _to_timezone(record.created_at, self.timezone)
        received_at = _to_timezone(record.received_at, self.timezone)
        report_date = loaded_at.date().isoformat()
        recipients = tuple(json.loads(record.recipients or "[]"))
        cc = tuple(json.loads(record.cc or "[]"))
        priority = analysis.priority if analysis else "medium"
        return ReportEmail(
            gmail_message_id=record.gmail_message_id,
            gmail_thread_id=record.gmail_thread_id,
            sender=record.sender,
            recipients=recipients,
            cc=cc,
            subject=record.subject,
            loaded_at=loaded_at,
            loaded_at_display=loaded_at.strftime("%Y-%m-%d %H:%M"),
            received_at=received_at,
            received_at_display=(
                received_at.strftime("%Y-%m-%d %H:%M") if received_at else "日時不明"
            ),
            report_date=report_date,
            snippet=record.snippet,
            body=record.body,
            body_html=record.body_html,
            has_attachments=record.has_attachments,
            attachments=tuple(
                AttachmentInfo(
                    filename=item.get("filename", ""),
                    mime_type=item.get("mime_type", ""),
                    size=item.get("size"),
                    attachment_id=item.get("attachment_id"),
                )
                for item in json.loads(record.attachment_metadata or "[]")
            ),
            is_sent=is_sent,
            sender_color="#ffffff" if is_sent else _sender_background_color(record.sender),
            message_href=f"messages/{_safe_filename(record.gmail_message_id)}.html",
            analysis=ReportAnalysis(
                priority=priority,
                summary_ja=analysis.summary_ja if analysis else "",
                requires_reply=bool(analysis.requires_reply) if analysis else False,
                category=analysis.category if analysis else "",
                meeting_candidates=(
                    json.loads(analysis.meeting_candidates_json or "[]")
                    if analysis
                    else []
                ),
            ),
            digest=(
                ReportDigest(
                    summary_ja=digest.summary_ja,
                    translation_ja=(
                        digest.translation_ja
                        if _looks_like_japanese_translation(digest.translation_ja)
                        else ""
                    ),
                )
                if digest
                else None
            ),
        )


def _load_timezone(timezone: str) -> ZoneInfo | fixed_timezone:
    try:
        return ZoneInfo(timezone)
    except ZoneInfoNotFoundError:
        if timezone == "Asia/Tokyo":
            return fixed_timezone(timedelta(hours=9), name="Asia/Tokyo")
        return UTC


def _to_timezone(
    value: datetime | None,
    timezone: ZoneInfo | fixed_timezone,
) -> datetime | None:
    if value is None:
        return None
    if value.tzinfo is None:
        value = value.replace(tzinfo=UTC)
    return value.astimezone(timezone)


def _group_by_date(emails: list[ReportEmail]) -> dict[str, list[ReportEmail]]:
    grouped: dict[str, list[ReportEmail]] = {}
    for email in emails:
        grouped.setdefault(email.report_date, []).append(email)
    return grouped


def _safe_filename(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_.-]+", "_", value).strip("._")
    return cleaned or "message"


def _gmail_thread_href(thread_id: str) -> str:
    return f"https://mail.google.com/mail/u/0/#all/{thread_id}"


def _build_report_threads(
    emails: list[ReportEmail],
    all_emails: list[ReportEmail],
    report_date: str,
) -> list[ReportThread]:
    all_groups: dict[str, list[ReportEmail]] = {}
    for email in all_emails:
        key = email.gmail_thread_id or email.gmail_message_id
        all_groups.setdefault(key, []).append(email)

    current_groups: dict[str, list[ReportEmail]] = {}
    order: list[str] = []
    for email in emails:
        key = email.gmail_thread_id or email.gmail_message_id
        if key not in current_groups:
            current_groups[key] = []
            order.append(key)
        current_groups[key].append(email)

    threads = [
        _thread_from_emails(
            key=key,
            detail_emails=all_groups.get(key, current_groups[key]),
            current_emails=current_groups[key],
            report_date=report_date,
        )
        for key in order
    ]
    threads.sort(
        key=lambda thread: (
            PRIORITY_RANK.get(thread.analysis.priority, 2),
            thread.received_at or datetime.min.replace(tzinfo=UTC),
            thread.loaded_at,
        ),
        reverse=True,
    )
    return threads


def _thread_from_emails(
    *,
    key: str,
    detail_emails: list[ReportEmail],
    current_emails: list[ReportEmail],
    report_date: str,
) -> ReportThread:
    current_messages = sorted(
        [email for email in detail_emails if email.report_date == report_date],
        key=lambda email: (
            email.received_at or datetime.min.replace(tzinfo=UTC),
            email.loaded_at,
        ),
        reverse=True,
    )
    previous_messages = sorted(
        [email for email in detail_emails if email.report_date != report_date],
        key=lambda email: (
            email.received_at or datetime.min.replace(tzinfo=UTC),
            email.loaded_at,
        ),
        reverse=True,
    )
    messages = sorted(
        detail_emails,
        key=lambda email: (
            email.received_at or datetime.min.replace(tzinfo=UTC),
            email.loaded_at,
        ),
        reverse=True,
    )
    primary = current_messages[0]
    highest_analysis = max(
        (message.analysis for message in current_messages),
        key=lambda analysis: PRIORITY_RANK.get(analysis.priority, 2),
    )
    return ReportThread(
        key=key,
        message_href=(
            f"messages/{_safe_filename(primary.gmail_message_id if len(messages) == 1 else key)}.html"
        ),
        gmail_href=_gmail_thread_href(key),
        subject=primary.subject,
        sender=primary.sender,
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
        loaded_at=primary.loaded_at,
        loaded_at_display=primary.loaded_at_display,
        received_at=primary.received_at,
        received_at_display=primary.received_at_display,
        report_date=report_date,
        snippet=primary.snippet,
        search_text=" ".join(
            " ".join(
                [
                    message.sender,
                    " ".join(message.recipients),
                    message.subject,
                    message.snippet,
                    message.body[:300],
                ]
            )
            for message in messages
        ),
        analysis=highest_analysis,
        messages=tuple(messages),
        current_message_count=len(current_messages),
    )


def _merge_duplicate_report_emails(emails: list[ReportEmail]) -> list[ReportEmail]:
    groups: dict[tuple[str, str, str], list[ReportEmail]] = {}
    order: list[tuple[str, str, str]] = []
    for email in emails:
        key = _duplicate_report_key(email)
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(email)

    merged: list[ReportEmail] = []
    for key in order:
        group = groups[key]
        primary = group[0]
        if len(group) == 1:
            merged.append(primary)
            continue
        merged.append(
            replace(
                primary,
                recipients=_unique_values(
                    recipient
                    for email in group
                    for recipient in email.recipients
                ),
                cc=_unique_values(
                    recipient
                    for email in group
                    for recipient in email.cc
                ),
            )
        )
    return merged


def _duplicate_report_key(email: ReportEmail) -> tuple[str, str, str]:
    body_key = email.body or email.body_html or email.snippet
    return (
        _normalize_email_address(email.sender),
        _normalize_text(email.subject),
        _normalize_text(body_key),
    )


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


def _attachment_summary(attachments: Iterable[AttachmentInfo]) -> str:
    counts: dict[str, int] = {}
    for attachment in attachments:
        extension = Path(attachment.filename or "").suffix.lower().lstrip(".")
        if not extension:
            extension = "file"
        counts[extension] = counts.get(extension, 0) + 1
    return ", ".join(f"{extension} x {counts[extension]}" for extension in sorted(counts))


def _sender_background_color(sender: str) -> str:
    address = _normalize_email_address(sender) or sender.strip().lower()
    seed = sum((index + 1) * ord(character) for index, character in enumerate(address))
    hue = seed % 360
    return f"hsl({hue} 42% 96%)"


def _looks_like_japanese_translation(text: str) -> bool:
    if not text:
        return False
    sampled = text[:4000]
    japanese_count = sum(
        1
        for character in sampled
        if "\u3040" <= character <= "\u30ff" or "\u4e00" <= character <= "\u9fff"
    )
    ascii_letter_count = sum(1 for character in sampled if character.isascii() and character.isalpha())
    return japanese_count >= 6 and japanese_count >= max(3, ascii_letter_count // 8)


def _normalize_email_address(value: str) -> str:
    return parseaddr(value)[1].strip().lower()


def _is_sent_record(record: EmailRecord, excluded_sender_addresses: set[str]) -> bool:
    if "SENT" in set(json.loads(record.label_ids or "[]")):
        return True
    return _normalize_email_address(record.sender) in excluded_sender_addresses


def _build_calendar(current_date: date, available_dates: list[str]) -> list[CalendarWeek]:
    available = set(available_dates)
    weeks = []
    for week in calendar.Calendar(firstweekday=0).monthdayscalendar(
        current_date.year,
        current_date.month,
    ):
        days = []
        for day in week:
            if day == 0:
                days.append(None)
                continue
            day_date = date(current_date.year, current_date.month, day)
            day_key = day_date.isoformat()
            days.append(
                CalendarDay(
                    day=day,
                    href=(
                        "index.html"
                        if day_key == current_date.isoformat()
                        else f"../{day_key}/index.html"
                    ),
                    has_emails=day_key in available,
                    is_current=day_key == current_date.isoformat(),
                )
            )
        weeks.append(CalendarWeek(tuple(days)))
    return weeks
