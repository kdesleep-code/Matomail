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
from .database import CalendarEventRecord, Database, EmailAnalysisRecord, EmailDigestRecord, EmailRecord, FilterDecisionRecord, ProcessingStateRecord
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
class ReportCalendarEvent:
    id: int
    title: str
    start_time: datetime | None
    end_time: datetime | None
    start_display: str
    end_display: str
    timezone: str
    location: str
    calendar_event_id: str
    status: str


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
    is_opened: bool
    is_resolved: bool
    completed_at: datetime | None
    completed_date: str
    sender_color: str
    message_href: str
    calendar_candidates_hidden: bool
    filter_action: str
    analysis: ReportAnalysis
    digest: ReportDigest | None
    calendar_events: tuple[ReportCalendarEvent, ...]


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
    def latest_message(self) -> ReportEmail:
        return self.messages[0]

    @property
    def is_read(self) -> bool:
        return self.latest_message.is_opened

    @property
    def status_class(self) -> str:
        return "is-read" if self.is_read else "is-unread"

    @property
    def filter_action(self) -> str:
        return self.primary_message.filter_action

    @property
    def filter_label(self) -> str:
        if self.filter_action == FILTER_ACTION_IGNORE:
            return "ignore"
        if self.filter_action == FILTER_ACTION_SKIP_ANALYSIS:
            return "skip"
        return ""

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
        skipped_emails = _skipped_list_emails(emails)
        non_skipped_emails = [email for email in emails if not _is_skipped_email(email)]
        list_emails = _active_list_emails(non_skipped_emails)
        completed_emails = _completed_list_emails(non_skipped_emails)
        grouped = _group_by_date(list_emails)
        completed_grouped = _group_by_completed_date(completed_emails)
        skipped_grouped = _group_by_date(skipped_emails)
        for date_key in completed_grouped:
            grouped.setdefault(date_key, [])
        for date_key in skipped_grouped:
            grouped.setdefault(date_key, [])
        if not grouped:
            return None

        dates = sorted(grouped)
        for index, date_key in enumerate(dates):
            self._write_daily_report(
                date_key=date_key,
                emails=grouped[date_key],
                completed_emails=completed_grouped.get(date_key, []),
                skipped_emails=skipped_grouped.get(date_key, []),
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
        completed_emails: list[ReportEmail],
        skipped_emails: list[ReportEmail],
        all_emails: list[ReportEmail],
        available_dates: list[str],
        previous_date: str | None,
        next_date: str | None,
    ) -> Path:
        daily_dir = self.report_dir / date_key
        messages_dir = daily_dir / "messages"
        messages_dir.mkdir(parents=True, exist_ok=True)
        for stale_message_path in messages_dir.glob("*.html"):
            stale_message_path.unlink()
        threads = _build_report_threads(emails, all_emails, date_key)
        completed_threads = _build_report_threads(completed_emails, all_emails, date_key)
        skipped_threads = _build_report_threads(skipped_emails, all_emails, date_key)

        template = self.environment.get_template("report.html.j2")
        report_day = date.fromisoformat(date_key)
        context = {
            "generated_at": datetime.now(self.timezone).strftime("%Y-%m-%d %H:%M"),
            "report_date": date_key,
            "emails": threads,
            "active_emails": threads,
            "completed_emails": completed_threads,
            "skipped_emails": skipped_threads,
            "email_count": sum(thread.current_message_count for thread in threads),
            "thread_count": len(threads),
            "active_email_count": sum(thread.current_message_count for thread in threads),
            "active_thread_count": len(threads),
            "completed_email_count": sum(
                thread.current_message_count for thread in completed_threads
            ),
            "completed_thread_count": len(completed_threads),
            "skipped_email_count": sum(
                thread.current_message_count for thread in skipped_threads
            ),
            "skipped_thread_count": len(skipped_threads),
            "previous_href": f"../{previous_date}/index.html" if previous_date else "",
            "next_href": f"../{next_date}/index.html" if next_date else "",
            "calendar_weeks": _build_calendar(report_day, available_dates),
            "calendar_month": report_day.strftime("%Y-%m"),
            "priority_rank": PRIORITY_RANK,
            "page": "list",
        }
        index_path = daily_dir / "index.html"
        index_path.write_text(template.render(context), encoding="utf-8")

        detail_threads = _unique_threads([*threads, *completed_threads, *skipped_threads])
        for index, thread in enumerate(detail_threads):
            message_path = messages_dir / Path(thread.message_href).name
            previous_thread = detail_threads[index - 1] if index > 0 else None
            next_thread = detail_threads[index + 1] if index < len(detail_threads) - 1 else None
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
                state = session.scalar(
                    select(ProcessingStateRecord).where(
                        ProcessingStateRecord.gmail_message_id == record.gmail_message_id
                    )
                )
                calendar_events = session.scalars(
                    select(CalendarEventRecord)
                    .where(CalendarEventRecord.email_id == record.id)
                    .where(CalendarEventRecord.status == "registered")
                    .order_by(CalendarEventRecord.created_at.desc(), CalendarEventRecord.id.desc())
                ).all()
                emails.append(
                    self._record_to_report_email(
                        record,
                        analysis,
                        digest,
                        calendar_events,
                        is_sent,
                        bool(state.web_opened) if state else False,
                        bool(state.resolved) if state else False,
                        state.resolved_at if state else None,
                        bool(state.calendar_candidates_hidden) if state else False,
                        decision.action if decision else "",
                    )
                )
            return _merge_duplicate_report_emails(emails)

    def _record_to_report_email(
        self,
        record: EmailRecord,
        analysis: EmailAnalysisRecord | None,
        digest: EmailDigestRecord | None = None,
        calendar_events: list[CalendarEventRecord] | None = None,
        is_sent: bool = False,
        is_opened: bool = False,
        is_resolved: bool = False,
        resolved_at: datetime | None = None,
        calendar_candidates_hidden: bool = False,
        filter_action: str = "",
    ) -> ReportEmail:
        loaded_at = _to_timezone(record.created_at, self.timezone)
        received_at = _to_timezone(record.received_at, self.timezone)
        resolved_at_local = _to_timezone(resolved_at, self.timezone)
        completed_at = (received_at or loaded_at) if is_sent else resolved_at_local
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
            is_opened=is_opened,
            is_resolved=is_resolved,
            completed_at=completed_at,
            completed_date=completed_at.date().isoformat() if completed_at else "",
            sender_color="#ffffff" if is_sent else _sender_background_color(record.sender),
            message_href=f"messages/{_safe_filename(record.gmail_message_id)}.html",
            calendar_candidates_hidden=calendar_candidates_hidden,
            filter_action=filter_action,
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
            calendar_events=tuple(
                _calendar_event_to_report_event(event, self.timezone)
                for event in (calendar_events or [])
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


def _calendar_event_to_report_event(
    event: CalendarEventRecord,
    timezone: ZoneInfo | fixed_timezone,
) -> ReportCalendarEvent:
    start_time = _to_timezone(event.start_time, timezone)
    end_time = _to_timezone(event.end_time, timezone)
    return ReportCalendarEvent(
        id=event.id,
        title=event.title,
        start_time=start_time,
        end_time=end_time,
        start_display=start_time.strftime("%Y-%m-%d %H:%M") if start_time else "",
        end_display=end_time.strftime("%Y-%m-%d %H:%M") if end_time else "",
        timezone=event.timezone,
        location=event.location,
        calendar_event_id=event.calendar_event_id,
        status=event.status,
    )


def _group_by_date(emails: list[ReportEmail]) -> dict[str, list[ReportEmail]]:
    grouped: dict[str, list[ReportEmail]] = {}
    for email in emails:
        grouped.setdefault(email.report_date, []).append(email)
    return grouped


def _group_by_completed_date(emails: list[ReportEmail]) -> dict[str, list[ReportEmail]]:
    grouped: dict[str, list[ReportEmail]] = {}
    for email in emails:
        if email.completed_date:
            grouped.setdefault(email.completed_date, []).append(email)
    return grouped


def _active_list_emails(emails: list[ReportEmail]) -> list[ReportEmail]:
    groups: dict[str, list[ReportEmail]] = {}
    for email in emails:
        key = email.gmail_thread_id or email.gmail_message_id
        groups.setdefault(key, []).append(email)

    completed_keys = set()
    for key, group in groups.items():
        latest = _latest_email(group)
        if latest.is_sent or latest.is_resolved:
            completed_keys.add(key)

    return [
        email
        for email in emails
        if not email.is_sent
        and (email.gmail_thread_id or email.gmail_message_id) not in completed_keys
    ]


def _completed_list_emails(emails: list[ReportEmail]) -> list[ReportEmail]:
    groups: dict[str, list[ReportEmail]] = {}
    for email in emails:
        key = email.gmail_thread_id or email.gmail_message_id
        groups.setdefault(key, []).append(email)

    completed: list[ReportEmail] = []
    for group in groups.values():
        latest = _latest_email(group)
        if latest.is_sent or latest.is_resolved:
            completed.append(latest)
    return completed


def _skipped_list_emails(emails: list[ReportEmail]) -> list[ReportEmail]:
    return [
        email
        for email in emails
        if _is_skipped_email(email) and not email.is_sent
    ]


def _is_skipped_email(email: ReportEmail) -> bool:
    return email.filter_action in {
        FILTER_ACTION_IGNORE,
        FILTER_ACTION_SKIP_ANALYSIS,
    }


def _unique_threads(threads: list[ReportThread]) -> list[ReportThread]:
    unique: list[ReportThread] = []
    seen: set[str] = set()
    for thread in threads:
        if thread.key in seen:
            continue
        seen.add(thread.key)
        unique.append(thread)
    return unique


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
        current_emails,
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


def _latest_email(emails: list[ReportEmail]) -> ReportEmail:
    return max(
        emails,
        key=lambda email: (
            email.received_at or datetime.min.replace(tzinfo=UTC),
            email.loaded_at,
        ),
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
