"""SQLite persistence for Matomail."""

from __future__ import annotations

import json
import shutil
from datetime import UTC, datetime
from email.utils import parseaddr
from pathlib import Path
from typing import Iterable

from sqlalchemy import (
    Boolean,
    DateTime,
    Float,
    ForeignKey,
    Integer,
    String,
    Text,
    create_engine,
    inspect,
    select,
    text,
)
from sqlalchemy.orm import (
    DeclarativeBase,
    Mapped,
    Session,
    mapped_column,
    relationship,
    sessionmaker,
)

from .models import AttachmentInfo, EmailMessage


FINAL_STATUSES = {"processed", "skipped"}
PENDING_STATUS = "pending"
FILTER_ACTION_ALWAYS_PROCESS = "always_process"
FILTER_ACTION_IGNORE = "ignore"
FILTER_ACTION_PRECLASSIFY = "preclassify"
FILTER_ACTION_SKIP_ANALYSIS = "skip_analysis"
FILTER_ACTIONS = {
    FILTER_ACTION_IGNORE,
    FILTER_ACTION_ALWAYS_PROCESS,
    FILTER_ACTION_PRECLASSIFY,
    FILTER_ACTION_SKIP_ANALYSIS,
}
FILTER_ACTION_PRECEDENCE = {
    FILTER_ACTION_IGNORE: 30,
    FILTER_ACTION_PRECLASSIFY: 25,
    FILTER_ACTION_ALWAYS_PROCESS: 20,
    FILTER_ACTION_SKIP_ANALYSIS: 10,
}


class MailBase(DeclarativeBase):
    pass


class RulesBase(DeclarativeBase):
    pass


class EmailRecord(MailBase):
    __tablename__ = "emails"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    gmail_message_id: Mapped[str] = mapped_column(String, unique=True, index=True)
    gmail_thread_id: Mapped[str] = mapped_column(String, index=True)
    sender: Mapped[str] = mapped_column(Text, default="")
    recipients: Mapped[str] = mapped_column(Text, default="[]")
    cc: Mapped[str] = mapped_column(Text, default="[]")
    subject: Mapped[str] = mapped_column(Text, default="")
    received_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True))
    snippet: Mapped[str] = mapped_column(Text, default="")
    body: Mapped[str] = mapped_column(Text, default="")
    body_html: Mapped[str] = mapped_column(Text, default="")
    label_ids: Mapped[str] = mapped_column(Text, default="[]")
    sender_candidates: Mapped[str] = mapped_column(Text, default="[]")
    size_estimate: Mapped[int | None] = mapped_column(Integer)
    has_attachments: Mapped[bool] = mapped_column(Boolean, default=False)
    attachment_metadata: Mapped[str] = mapped_column(Text, default="[]")
    created_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        default=lambda: datetime.now(UTC),
    )

    analysis: Mapped[list["EmailAnalysisRecord"]] = relationship(back_populates="email")


class EmailAnalysisRecord(MailBase):
    __tablename__ = "email_analysis"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    email_id: Mapped[int] = mapped_column(ForeignKey("emails.id"), index=True)
    summary_ja: Mapped[str] = mapped_column(Text, default="")
    category: Mapped[str] = mapped_column(String, default="")
    priority: Mapped[str] = mapped_column(String, default="medium")
    requires_reply: Mapped[bool] = mapped_column(Boolean, default=False)
    suggested_action_ja: Mapped[str] = mapped_column(Text, default="")
    deadline_candidates_json: Mapped[str] = mapped_column(Text, default="[]")
    meeting_candidates_json: Mapped[str] = mapped_column(Text, default="[]")
    reply_draft: Mapped[str] = mapped_column(Text, default="")
    confidence: Mapped[float] = mapped_column(Float, default=0.0)
    llm_model: Mapped[str] = mapped_column(String, default="")
    created_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        default=lambda: datetime.now(UTC),
    )

    email: Mapped[EmailRecord] = relationship(back_populates="analysis")


class EmailDigestRecord(MailBase):
    __tablename__ = "email_digests"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    email_id: Mapped[int] = mapped_column(ForeignKey("emails.id"), unique=True, index=True)
    summary_ja: Mapped[str] = mapped_column(Text, default="")
    translation_ja: Mapped[str] = mapped_column(Text, default="")
    llm_model: Mapped[str] = mapped_column(String, default="")
    created_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        default=lambda: datetime.now(UTC),
    )


class ProcessingStateRecord(MailBase):
    __tablename__ = "processing_state"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    gmail_message_id: Mapped[str] = mapped_column(String, unique=True, index=True)
    status: Mapped[str] = mapped_column(String, default=PENDING_STATUS, index=True)
    action_taken: Mapped[str] = mapped_column(Text, default="")
    reply_sent: Mapped[bool] = mapped_column(Boolean, default=False)
    calendar_registered: Mapped[bool] = mapped_column(Boolean, default=False)
    attachment_opened: Mapped[bool] = mapped_column(Boolean, default=False)
    report_path: Mapped[str] = mapped_column(Text, default="")
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        default=lambda: datetime.now(UTC),
        onupdate=lambda: datetime.now(UTC),
    )


class AppSettingRecord(MailBase):
    __tablename__ = "app_settings"

    key: Mapped[str] = mapped_column(String, primary_key=True)
    value: Mapped[str] = mapped_column(Text, default="")
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        default=lambda: datetime.now(UTC),
        onupdate=lambda: datetime.now(UTC),
    )


class FilterDecisionRecord(MailBase):
    __tablename__ = "filter_decisions"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    email_id: Mapped[int] = mapped_column(ForeignKey("emails.id"), index=True)
    gmail_message_id: Mapped[str] = mapped_column(String, index=True)
    action: Mapped[str] = mapped_column(String, index=True)
    matched_rule_id: Mapped[int | None] = mapped_column(Integer)
    matched_rule_name: Mapped[str] = mapped_column(Text, default="")
    reason: Mapped[str] = mapped_column(Text, default="")
    rule_snapshot_json: Mapped[str] = mapped_column(Text, default="{}")
    decided_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        default=lambda: datetime.now(UTC),
    )


class FilterRuleRecord(RulesBase):
    __tablename__ = "filter_rules"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    name: Mapped[str] = mapped_column(String, default="")
    action: Mapped[str] = mapped_column(String, default=FILTER_ACTION_SKIP_ANALYSIS)
    priority: Mapped[int] = mapped_column(Integer, default=0, index=True)
    preset_priority: Mapped[str] = mapped_column(String, default="")
    preset_category: Mapped[str] = mapped_column(String, default="")
    preset_summary_ja: Mapped[str] = mapped_column(Text, default="")
    preset_suggested_action_ja: Mapped[str] = mapped_column(Text, default="")
    preset_requires_reply: Mapped[bool | None] = mapped_column(Boolean)
    preset_reply_recommended: Mapped[bool | None] = mapped_column(Boolean)
    from_query: Mapped[str] = mapped_column(Text, default="")
    to_query: Mapped[str] = mapped_column(Text, default="")
    delivered_to_query: Mapped[str] = mapped_column(Text, default="")
    cc_query: Mapped[str] = mapped_column(Text, default="")
    bcc_query: Mapped[str] = mapped_column(Text, default="")
    subject_query: Mapped[str] = mapped_column(Text, default="")
    has_words: Mapped[str] = mapped_column(Text, default="")
    doesnt_have: Mapped[str] = mapped_column(Text, default="")
    gmail_query: Mapped[str] = mapped_column(Text, default="")
    negated_gmail_query: Mapped[str] = mapped_column(Text, default="")
    has_attachment: Mapped[bool | None] = mapped_column(Boolean)
    filename_query: Mapped[str] = mapped_column(Text, default="")
    size_comparison: Mapped[str] = mapped_column(String, default="")
    size_bytes: Mapped[int | None] = mapped_column(Integer)
    after: Mapped[datetime | None] = mapped_column(DateTime(timezone=True))
    before: Mapped[datetime | None] = mapped_column(DateTime(timezone=True))
    older_than: Mapped[str] = mapped_column(String, default="")
    newer_than: Mapped[str] = mapped_column(String, default="")
    category: Mapped[str] = mapped_column(String, default="")
    label: Mapped[str] = mapped_column(String, default="")
    include_chats: Mapped[bool] = mapped_column(Boolean, default=False)
    note: Mapped[str] = mapped_column(Text, default="")
    enabled: Mapped[bool] = mapped_column(Boolean, default=True, index=True)
    created_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        default=lambda: datetime.now(UTC),
    )


class LLMInstructionRuleRecord(RulesBase):
    __tablename__ = "llm_instruction_rules"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    name: Mapped[str] = mapped_column(String, default="")
    instruction: Mapped[str] = mapped_column(Text)
    priority: Mapped[int] = mapped_column(Integer, default=0, index=True)
    from_query: Mapped[str] = mapped_column(Text, default="")
    to_query: Mapped[str] = mapped_column(Text, default="")
    delivered_to_query: Mapped[str] = mapped_column(Text, default="")
    cc_query: Mapped[str] = mapped_column(Text, default="")
    bcc_query: Mapped[str] = mapped_column(Text, default="")
    subject_query: Mapped[str] = mapped_column(Text, default="")
    has_words: Mapped[str] = mapped_column(Text, default="")
    doesnt_have: Mapped[str] = mapped_column(Text, default="")
    gmail_query: Mapped[str] = mapped_column(Text, default="")
    negated_gmail_query: Mapped[str] = mapped_column(Text, default="")
    has_attachment: Mapped[bool | None] = mapped_column(Boolean)
    filename_query: Mapped[str] = mapped_column(Text, default="")
    size_comparison: Mapped[str] = mapped_column(String, default="")
    size_bytes: Mapped[int | None] = mapped_column(Integer)
    after: Mapped[datetime | None] = mapped_column(DateTime(timezone=True))
    before: Mapped[datetime | None] = mapped_column(DateTime(timezone=True))
    older_than: Mapped[str] = mapped_column(String, default="")
    newer_than: Mapped[str] = mapped_column(String, default="")
    category: Mapped[str] = mapped_column(String, default="")
    label: Mapped[str] = mapped_column(String, default="")
    include_chats: Mapped[bool] = mapped_column(Boolean, default=False)
    note: Mapped[str] = mapped_column(Text, default="")
    enabled: Mapped[bool] = mapped_column(Boolean, default=True, index=True)
    created_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        default=lambda: datetime.now(UTC),
    )


class Database:
    """Owns local persistence for Matomail."""

    def __init__(
        self,
        db_path: Path | str,
        max_size_bytes: int | None = None,
        backup_dir: Path | str | None = None,
        store_email_body: bool = True,
    ) -> None:
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self.backup_dir = Path(backup_dir) if backup_dir is not None else self.db_path.parent / "backups"
        self.store_email_body = store_email_body
        self.rotate_if_needed(max_size_bytes)
        self.engine = create_engine(f"sqlite:///{self.db_path}", future=True)
        self.session_factory = sessionmaker(self.engine, expire_on_commit=False)

    def create_all(self) -> None:
        MailBase.metadata.create_all(self.engine)
        self._ensure_schema()

    def _ensure_schema(self) -> None:
        columns = {column["name"] for column in inspect(self.engine).get_columns("emails")}
        if "body_html" not in columns:
            with self.engine.begin() as connection:
                connection.execute(
                    text("ALTER TABLE emails ADD COLUMN body_html TEXT DEFAULT ''")
                )
        if "label_ids" not in columns:
            with self.engine.begin() as connection:
                connection.execute(
                    text("ALTER TABLE emails ADD COLUMN label_ids TEXT DEFAULT '[]'")
                )
        if "sender_candidates" not in columns:
            with self.engine.begin() as connection:
                connection.execute(
                    text("ALTER TABLE emails ADD COLUMN sender_candidates TEXT DEFAULT '[]'")
                )

    def rotate_if_needed(self, max_size_bytes: int | None) -> Path | None:
        if max_size_bytes is None or not self.db_path.exists():
            return None
        if self.db_path.stat().st_size <= max_size_bytes:
            return None

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        destination_dir = self.backup_dir / timestamp
        suffix = 1
        while destination_dir.exists():
            destination_dir = self.backup_dir / f"{timestamp}_{suffix}"
            suffix += 1

        destination_dir.mkdir(parents=True, exist_ok=False)
        destination = destination_dir / self.db_path.name
        shutil.move(str(self.db_path), destination)
        return destination

    def save_email(self, message: EmailMessage) -> EmailRecord:
        with self.session_factory() as session:
            record = self._get_email_record(session, message.gmail_message_id)
            if record is None:
                record = EmailRecord(gmail_message_id=message.gmail_message_id)
                session.add(record)

            record.gmail_thread_id = message.gmail_thread_id
            record.sender = message.sender
            record.recipients = json.dumps(list(message.recipients), ensure_ascii=False)
            record.cc = json.dumps(list(message.cc), ensure_ascii=False)
            record.subject = message.subject
            record.received_at = message.received_at
            record.snippet = message.snippet
            record.body = message.body if self.store_email_body else ""
            record.body_html = message.body_html if self.store_email_body else ""
            record.label_ids = json.dumps(list(message.label_ids), ensure_ascii=False)
            record.sender_candidates = json.dumps(
                list(message.sender_candidates),
                ensure_ascii=False,
            )
            record.size_estimate = message.size_estimate
            record.has_attachments = message.has_attachments
            record.attachment_metadata = json.dumps(
                [
                    {
                        "filename": attachment.filename,
                        "mime_type": attachment.mime_type,
                        "size": attachment.size,
                        "attachment_id": attachment.attachment_id,
                    }
                    for attachment in message.attachments
                ],
                ensure_ascii=False,
            )

            self._ensure_processing_state(session, message.gmail_message_id)
            session.commit()
            return record

    def save_analysis(
        self,
        gmail_message_id: str,
        analysis: dict,
        llm_model: str,
    ) -> EmailAnalysisRecord:
        with self.session_factory() as session:
            email_record = self._get_email_record(session, gmail_message_id)
            if email_record is None:
                raise ValueError(f"email is not saved: {gmail_message_id}")

            record = EmailAnalysisRecord(
                email_id=email_record.id,
                summary_ja=str(analysis["summary_ja"]),
                category=str(analysis["category"]),
                priority=str(analysis["priority"]),
                requires_reply=bool(analysis["requires_reply"]),
                suggested_action_ja=str(analysis["suggested_action_ja"]),
                deadline_candidates_json=json.dumps(
                    analysis["deadline_candidates"],
                    ensure_ascii=False,
                ),
                meeting_candidates_json=json.dumps(
                    analysis["meeting_candidates"],
                    ensure_ascii=False,
                ),
                reply_draft=str(analysis["reply_draft_ja"]),
                confidence=float(analysis["confidence"]),
                llm_model=llm_model,
            )
            session.add(record)
            session.commit()
            return record

    def save_digest(
        self,
        gmail_message_id: str,
        *,
        summary_ja: str,
        translation_ja: str,
        llm_model: str,
    ) -> EmailDigestRecord:
        with self.session_factory() as session:
            email_record = self._get_email_record(session, gmail_message_id)
            if email_record is None:
                raise ValueError(f"email is not saved: {gmail_message_id}")

            record = session.scalar(
                select(EmailDigestRecord).where(
                    EmailDigestRecord.email_id == email_record.id
                )
            )
            if record is None:
                record = EmailDigestRecord(email_id=email_record.id)
                session.add(record)
            record.summary_ja = summary_ja
            record.translation_ja = translation_ja
            record.llm_model = llm_model
            record.created_at = datetime.now(UTC)
            session.commit()
            return record

    def get_digest(self, gmail_message_id: str) -> EmailDigestRecord | None:
        with self.session_factory() as session:
            email_record = self._get_email_record(session, gmail_message_id)
            if email_record is None:
                return None
            return session.scalar(
                select(EmailDigestRecord).where(
                    EmailDigestRecord.email_id == email_record.id
                )
            )

    def save_filter_decision(
        self,
        gmail_message_id: str,
        action: str,
        matched_rule_id: int | None = None,
        matched_rule_name: str = "",
        reason: str = "",
        rule_snapshot: dict | None = None,
    ) -> FilterDecisionRecord:
        with self.session_factory() as session:
            email_record = self._get_email_record(session, gmail_message_id)
            if email_record is None:
                raise ValueError(f"email is not saved: {gmail_message_id}")

            existing = session.scalar(
                select(FilterDecisionRecord).where(
                    FilterDecisionRecord.email_id == email_record.id
                )
            )
            record = existing or FilterDecisionRecord(
                email_id=email_record.id,
                gmail_message_id=gmail_message_id,
            )
            record.action = action
            record.matched_rule_id = matched_rule_id
            record.matched_rule_name = matched_rule_name
            record.reason = reason
            record.rule_snapshot_json = json.dumps(
                rule_snapshot or {},
                ensure_ascii=False,
            )
            record.decided_at = datetime.now(UTC)
            session.add(record)
            session.commit()
            return record

    def apply_preclassified_analysis(
        self,
        gmail_message_id: str,
        preset_analysis: dict,
        llm_model: str = "rule:preclassify",
    ) -> EmailAnalysisRecord:
        analysis = {
            "summary_ja": preset_analysis.get("summary_ja") or "ルールにより自動分類されました。",
            "category": preset_analysis.get("category") or "preclassified",
            "priority": preset_analysis.get("priority") or "medium",
            "requires_reply": bool(preset_analysis.get("requires_reply", False)),
            "suggested_action_ja": preset_analysis.get("suggested_action_ja") or "",
            "deadline_candidates": [],
            "meeting_candidates": [],
            "reply_draft_ja": "",
            "confidence": 1.0,
        }
        return self.save_analysis(gmail_message_id, analysis, llm_model=llm_model)

    def save_emails(self, messages: Iterable[EmailMessage]) -> list[EmailRecord]:
        return [self.save_email(message) for message in messages]

    def list_unanalyzed_emails(self, limit: int = 1) -> list[EmailMessage]:
        with self.session_factory() as session:
            records = session.scalars(
                select(EmailRecord)
                .outerjoin(EmailAnalysisRecord)
                .outerjoin(FilterDecisionRecord)
                .where(EmailAnalysisRecord.id.is_(None))
                .where(
                    (FilterDecisionRecord.id.is_(None))
                    | (
                        FilterDecisionRecord.action
                        == FILTER_ACTION_ALWAYS_PROCESS
                    )
                )
                .order_by(EmailRecord.received_at.desc(), EmailRecord.id.desc())
            ).all()
            messages = [
                _email_record_to_message(record)
                for record in records
                if not _email_record_is_sent(record)
            ]
            return messages[:limit]

    def list_emails_needing_digest(
        self,
        priorities: set[str],
        limit: int,
    ) -> list[EmailMessage]:
        with self.session_factory() as session:
            records = session.scalars(select(EmailRecord)).all()
            selected: list[tuple[datetime | None, EmailRecord]] = []
            for record in records:
                digest = session.scalar(
                    select(EmailDigestRecord).where(
                        EmailDigestRecord.email_id == record.id
                    )
                )
                if digest is not None:
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
                if analysis is None or analysis.priority not in priorities:
                    continue
                selected.append((record.received_at, record))
            selected.sort(key=lambda item: item[0] or datetime.min.replace(tzinfo=UTC), reverse=True)
            return [_email_record_to_message(record) for _, record in selected[:limit]]

    def list_sent_emails_needing_digest(self, limit: int) -> list[EmailMessage]:
        with self.session_factory() as session:
            records = session.scalars(select(EmailRecord)).all()
            selected: list[tuple[datetime | None, EmailRecord]] = []
            for record in records:
                if not _email_record_is_sent(record):
                    continue
                digest = session.scalar(
                    select(EmailDigestRecord).where(
                        EmailDigestRecord.email_id == record.id
                    )
                )
                if digest is not None:
                    continue
                selected.append((record.received_at, record))
            selected.sort(key=lambda item: item[0] or datetime.min.replace(tzinfo=UTC), reverse=True)
            return [_email_record_to_message(record) for _, record in selected[:limit]]

    def list_high_priority_emails_without_star(self) -> list[EmailMessage]:
        with self.session_factory() as session:
            records = session.scalars(select(EmailRecord)).all()
            selected: list[tuple[datetime | None, EmailRecord]] = []
            for record in records:
                if _email_record_is_sent(record):
                    continue
                if "STARRED" in set(json.loads(record.label_ids or "[]")):
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
                if analysis is None or analysis.priority != "high":
                    continue
                selected.append((record.received_at, record))
            selected.sort(key=lambda item: item[0] or datetime.min.replace(tzinfo=UTC), reverse=True)
            return [_email_record_to_message(record) for _, record in selected]

    def get_email(self, gmail_message_id: str) -> EmailMessage | None:
        with self.session_factory() as session:
            record = self._get_email_record(session, gmail_message_id)
            if record is None:
                return None
            return _email_record_to_message(record)

    def has_email(self, gmail_message_id: str) -> bool:
        with self.session_factory() as session:
            return self._get_email_record(session, gmail_message_id) is not None

    def list_saved_emails(self) -> list[EmailMessage]:
        with self.session_factory() as session:
            records = session.scalars(
                select(EmailRecord).order_by(
                    EmailRecord.created_at.desc(),
                    EmailRecord.received_at.desc(),
                    EmailRecord.id.desc(),
                )
            ).all()
            return [_email_record_to_message(record) for record in records]

    def clear_filter_decisions(self) -> None:
        with self.session_factory() as session:
            for record in session.scalars(select(FilterDecisionRecord)).all():
                session.delete(record)
            session.commit()

    def list_emails_loaded_on_same_day(self, gmail_message_id: str) -> list[EmailMessage]:
        with self.session_factory() as session:
            source = self._get_email_record(session, gmail_message_id)
            if source is None:
                return []
            loaded_date = source.created_at.date()
            records = session.scalars(select(EmailRecord)).all()
            return [
                _email_record_to_message(record)
                for record in records
                if record.created_at.date() == loaded_date
            ]

    def should_process(self, gmail_message_id: str) -> bool:
        with self.session_factory() as session:
            state = self._get_processing_state(session, gmail_message_id)
            return state is None or state.status not in FINAL_STATUSES

    def filter_processable(self, messages: Iterable[EmailMessage]) -> list[EmailMessage]:
        return [
            message
            for message in messages
            if self.should_process(message.gmail_message_id)
        ]

    def mark_status(
        self,
        gmail_message_id: str,
        status: str,
        action_taken: str = "",
        report_path: str = "",
    ) -> None:
        if status not in {*FINAL_STATUSES, PENDING_STATUS}:
            raise ValueError("status must be one of processed, pending, or skipped")

        with self.session_factory() as session:
            state = self._ensure_processing_state(session, gmail_message_id)
            state.status = status
            state.action_taken = action_taken
            state.report_path = report_path
            state.updated_at = datetime.now(UTC)
            session.commit()

    def save_account_email(self, email_address: str) -> None:
        normalized = _normalize_email_address(email_address)
        if not normalized:
            return
        with self.session_factory() as session:
            record = session.get(AppSettingRecord, "gmail_account_emails")
            existing: list[str] = []
            if record is not None and record.value:
                existing = json.loads(record.value)
            values = sorted({*existing, normalized})
            if record is None:
                record = AppSettingRecord(key="gmail_account_emails")
                session.add(record)
            record.value = json.dumps(values, ensure_ascii=False)
            record.updated_at = datetime.now(UTC)
            session.commit()

    def list_account_emails(self) -> tuple[str, ...]:
        with self.session_factory() as session:
            record = session.get(AppSettingRecord, "gmail_account_emails")
            if record is None or not record.value:
                return ()
            return tuple(json.loads(record.value))

    def count_emails(self) -> int:
        with self.session_factory() as session:
            return len(session.scalars(select(EmailRecord.id)).all())

    def get_processing_status(self, gmail_message_id: str) -> str | None:
        with self.session_factory() as session:
            state = self._get_processing_state(session, gmail_message_id)
            return state.status if state else None

    @staticmethod
    def _get_email_record(session: Session, gmail_message_id: str) -> EmailRecord | None:
        return session.scalar(
            select(EmailRecord).where(
                EmailRecord.gmail_message_id == gmail_message_id
            )
        )

    @staticmethod
    def _get_processing_state(
        session: Session,
        gmail_message_id: str,
    ) -> ProcessingStateRecord | None:
        return session.scalar(
            select(ProcessingStateRecord).where(
                ProcessingStateRecord.gmail_message_id == gmail_message_id
            )
        )

    def _ensure_processing_state(
        self,
        session: Session,
        gmail_message_id: str,
    ) -> ProcessingStateRecord:
        state = self._get_processing_state(session, gmail_message_id)
        if state is None:
            state = ProcessingStateRecord(
                gmail_message_id=gmail_message_id,
                status=PENDING_STATUS,
            )
            session.add(state)
        return state

    def close(self) -> None:
        self.engine.dispose()


class RulesDatabase:
    """Owns user-maintained filtering and LLM instruction rules."""

    def __init__(self, db_path: Path | str) -> None:
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self.engine = create_engine(f"sqlite:///{self.db_path}", future=True)
        self.session_factory = sessionmaker(self.engine, expire_on_commit=False)

    def create_all(self) -> None:
        RulesBase.metadata.create_all(self.engine)

    def filter_processable(self, messages: Iterable[EmailMessage]) -> list[EmailMessage]:
        return [
            message
            for message in messages
            if self.get_filter_action(message)
            not in {FILTER_ACTION_IGNORE, FILTER_ACTION_SKIP_ANALYSIS}
        ]

    def add_filter_rule(
        self,
        *,
        action: str,
        name: str = "",
        priority: int = 0,
        preset_priority: str = "",
        preset_category: str = "",
        preset_summary_ja: str = "",
        preset_suggested_action_ja: str = "",
        preset_requires_reply: bool | None = None,
        preset_reply_recommended: bool | None = None,
        from_query: str = "",
        to_query: str = "",
        delivered_to_query: str = "",
        cc_query: str = "",
        bcc_query: str = "",
        subject_query: str = "",
        has_words: str = "",
        doesnt_have: str = "",
        gmail_query: str = "",
        negated_gmail_query: str = "",
        has_attachment: bool | None = None,
        filename_query: str = "",
        size_comparison: str = "",
        size_bytes: int | None = None,
        after: datetime | None = None,
        before: datetime | None = None,
        older_than: str = "",
        newer_than: str = "",
        category: str = "",
        label: str = "",
        include_chats: bool = False,
        note: str = "",
        enabled: bool = True,
    ) -> FilterRuleRecord:
        if action not in FILTER_ACTIONS:
            raise ValueError(
                "action must be one of ignore, always_process, preclassify, or skip_analysis"
            )
        if preset_priority and preset_priority not in {"top", "high", "medium", "low"}:
            raise ValueError("preset_priority must be top, high, medium, or low")
        if size_comparison and size_comparison not in {"larger", "smaller"}:
            raise ValueError("size_comparison must be larger or smaller")

        if priority == 0:
            priority = self._next_filter_rule_priority()

        with self.session_factory() as session:
            rule = FilterRuleRecord(
                name=name,
                action=action,
                priority=priority,
                preset_priority=preset_priority,
                preset_category=preset_category,
                preset_summary_ja=preset_summary_ja,
                preset_suggested_action_ja=preset_suggested_action_ja,
                preset_requires_reply=preset_requires_reply,
                preset_reply_recommended=preset_reply_recommended,
                from_query=_normalize_email_address(from_query) or from_query.strip(),
                to_query=_normalize_email_address(to_query) or to_query.strip(),
                delivered_to_query=_normalize_email_address(delivered_to_query)
                or delivered_to_query.strip(),
                cc_query=_normalize_email_address(cc_query) or cc_query.strip(),
                bcc_query=_normalize_email_address(bcc_query) or bcc_query.strip(),
                subject_query=subject_query.strip(),
                has_words=has_words.strip(),
                doesnt_have=doesnt_have.strip(),
                gmail_query=gmail_query.strip(),
                negated_gmail_query=negated_gmail_query.strip(),
                has_attachment=has_attachment,
                filename_query=filename_query.strip(),
                size_comparison=size_comparison,
                size_bytes=size_bytes,
                after=after,
                before=before,
                older_than=older_than.strip(),
                newer_than=newer_than.strip(),
                category=category.strip(),
                label=label.strip(),
                include_chats=include_chats,
                note=note,
                enabled=enabled,
            )
            session.add(rule)
            session.commit()
            return rule

    def _next_filter_rule_priority(self) -> int:
        with self.session_factory() as session:
            priorities = session.scalars(select(FilterRuleRecord.priority)).all()
            return (max(priorities) if priorities else 0) + 1000

    def add_sender_filter(
        self,
        email_address: str,
        name: str = "",
        note: str = "",
        action: str = FILTER_ACTION_SKIP_ANALYSIS,
    ) -> FilterRuleRecord:
        normalized = _normalize_email_address(email_address)
        if not normalized:
            raise ValueError("email_address must contain a valid email address")
        return self.add_filter_rule(
            action=action,
            name=name,
            from_query=normalized,
            note=note,
        )

    def list_filter_rules(self) -> list[FilterRuleRecord]:
        with self.session_factory() as session:
            return session.scalars(
                select(FilterRuleRecord).order_by(
                    FilterRuleRecord.priority.asc(),
                    FilterRuleRecord.id.asc(),
                )
            ).all()

    def get_filter_rule(self, rule_id: int) -> FilterRuleRecord | None:
        with self.session_factory() as session:
            return session.get(FilterRuleRecord, rule_id)

    def update_filter_rule(
        self,
        rule_id: int,
        *,
        action: str,
        name: str = "",
        priority: int = 0,
        preset_priority: str = "",
        preset_category: str = "",
        preset_summary_ja: str = "",
        preset_suggested_action_ja: str = "",
        from_query: str = "",
        subject_query: str = "",
        has_words: str = "",
        note: str = "",
        enabled: bool = True,
    ) -> bool:
        if action not in FILTER_ACTIONS:
            raise ValueError(
                "action must be one of ignore, always_process, preclassify, or skip_analysis"
            )
        if preset_priority and preset_priority not in {"top", "high", "medium", "low"}:
            raise ValueError("preset_priority must be top, high, medium, or low")

        with self.session_factory() as session:
            rule = session.get(FilterRuleRecord, rule_id)
            if rule is None:
                return False
            rule.action = action
            rule.name = name
            rule.priority = priority
            rule.preset_priority = preset_priority
            rule.preset_category = preset_category
            rule.preset_summary_ja = preset_summary_ja
            rule.preset_suggested_action_ja = preset_suggested_action_ja
            rule.from_query = _normalize_email_address(from_query) or from_query.strip()
            rule.subject_query = subject_query.strip()
            rule.has_words = has_words.strip()
            rule.note = note
            rule.enabled = enabled
            session.commit()
            return True

    def move_filter_rule(self, rule_id: int, direction: str) -> bool:
        rules = self.list_filter_rules()
        index = next((i for i, rule in enumerate(rules) if rule.id == rule_id), None)
        if index is None:
            return False
        if direction == "up" and index > 0:
            rules[index - 1], rules[index] = rules[index], rules[index - 1]
        elif direction == "down" and index < len(rules) - 1:
            rules[index + 1], rules[index] = rules[index], rules[index + 1]
        else:
            return False

        return self.reorder_filter_rules([rule.id for rule in rules])

    def reorder_filter_rules(self, ordered_rule_ids: list[int]) -> bool:
        rules = self.list_filter_rules()
        existing_ids = [rule.id for rule in rules]
        if sorted(ordered_rule_ids) != sorted(existing_ids):
            return False

        with self.session_factory() as session:
            for offset, rule_id in enumerate(ordered_rule_ids):
                stored = session.get(FilterRuleRecord, rule_id)
                if stored is not None:
                    stored.priority = (offset + 1) * 1000
            session.commit()
        return True

    def set_filter_rule_enabled(self, rule_id: int, enabled: bool) -> bool:
        with self.session_factory() as session:
            rule = session.get(FilterRuleRecord, rule_id)
            if rule is None:
                return False
            rule.enabled = enabled
            session.commit()
            return True

    def delete_filter_rule(self, rule_id: int) -> bool:
        with self.session_factory() as session:
            rule = session.get(FilterRuleRecord, rule_id)
            if rule is None:
                return False
            session.delete(rule)
            session.commit()
            return True

    def add_llm_instruction_rule(
        self,
        *,
        instruction: str,
        name: str = "",
        priority: int = 0,
        from_query: str = "",
        to_query: str = "",
        delivered_to_query: str = "",
        cc_query: str = "",
        bcc_query: str = "",
        subject_query: str = "",
        has_words: str = "",
        doesnt_have: str = "",
        gmail_query: str = "",
        negated_gmail_query: str = "",
        has_attachment: bool | None = None,
        filename_query: str = "",
        size_comparison: str = "",
        size_bytes: int | None = None,
        after: datetime | None = None,
        before: datetime | None = None,
        older_than: str = "",
        newer_than: str = "",
        category: str = "",
        label: str = "",
        include_chats: bool = False,
        note: str = "",
        enabled: bool = True,
    ) -> LLMInstructionRuleRecord:
        cleaned_instruction = instruction.strip()
        if not cleaned_instruction:
            raise ValueError("instruction must not be empty")
        if size_comparison and size_comparison not in {"larger", "smaller"}:
            raise ValueError("size_comparison must be larger or smaller")
        if priority == 0:
            priority = self._next_llm_instruction_rule_priority()

        with self.session_factory() as session:
            rule = LLMInstructionRuleRecord(
                name=name,
                instruction=cleaned_instruction,
                priority=priority,
                from_query=_normalize_email_address(from_query) or from_query.strip(),
                to_query=_normalize_email_address(to_query) or to_query.strip(),
                delivered_to_query=_normalize_email_address(delivered_to_query)
                or delivered_to_query.strip(),
                cc_query=_normalize_email_address(cc_query) or cc_query.strip(),
                bcc_query=_normalize_email_address(bcc_query) or bcc_query.strip(),
                subject_query=subject_query.strip(),
                has_words=has_words.strip(),
                doesnt_have=doesnt_have.strip(),
                gmail_query=gmail_query.strip(),
                negated_gmail_query=negated_gmail_query.strip(),
                has_attachment=has_attachment,
                filename_query=filename_query.strip(),
                size_comparison=size_comparison,
                size_bytes=size_bytes,
                after=after,
                before=before,
                older_than=older_than.strip(),
                newer_than=newer_than.strip(),
                category=category.strip(),
                label=label.strip(),
                include_chats=include_chats,
                note=note,
                enabled=enabled,
            )
            session.add(rule)
            session.commit()
            return rule

    def _next_llm_instruction_rule_priority(self) -> int:
        with self.session_factory() as session:
            priorities = session.scalars(select(LLMInstructionRuleRecord.priority)).all()
            return (max(priorities) if priorities else 0) + 1000

    def get_llm_instructions_for_email(self, message: EmailMessage) -> list[str]:
        with self.session_factory() as session:
            rules = session.scalars(
                select(LLMInstructionRuleRecord).where(
                    LLMInstructionRuleRecord.enabled.is_(True)
                )
            ).all()

        matching_rules = [
            rule for rule in rules if _conditional_rule_matches(rule, message)
        ]
        matching_rules.sort(key=lambda rule: (rule.priority, rule.id))
        return [rule.instruction for rule in matching_rules]

    def list_llm_instruction_rules(self) -> list[LLMInstructionRuleRecord]:
        with self.session_factory() as session:
            return session.scalars(
                select(LLMInstructionRuleRecord).order_by(
                    LLMInstructionRuleRecord.priority.asc(),
                    LLMInstructionRuleRecord.id.asc(),
                )
            ).all()

    def get_llm_instruction_rule(self, rule_id: int) -> LLMInstructionRuleRecord | None:
        with self.session_factory() as session:
            return session.get(LLMInstructionRuleRecord, rule_id)

    def update_llm_instruction_rule(
        self,
        rule_id: int,
        *,
        instruction: str,
        name: str = "",
        priority: int = 0,
        from_query: str = "",
        to_query: str = "",
        subject_query: str = "",
        has_words: str = "",
        doesnt_have: str = "",
        note: str = "",
        enabled: bool = True,
    ) -> bool:
        cleaned_instruction = instruction.strip()
        if not cleaned_instruction:
            raise ValueError("instruction must not be empty")

        with self.session_factory() as session:
            rule = session.get(LLMInstructionRuleRecord, rule_id)
            if rule is None:
                return False
            rule.name = name
            rule.instruction = cleaned_instruction
            rule.priority = priority
            rule.from_query = _normalize_email_address(from_query) or from_query.strip()
            rule.to_query = _normalize_email_address(to_query) or to_query.strip()
            rule.subject_query = subject_query.strip()
            rule.has_words = has_words.strip()
            rule.doesnt_have = doesnt_have.strip()
            rule.note = note
            rule.enabled = enabled
            session.commit()
            return True

    def reorder_llm_instruction_rules(self, ordered_rule_ids: list[int]) -> bool:
        rules = self.list_llm_instruction_rules()
        existing_ids = [rule.id for rule in rules]
        if sorted(ordered_rule_ids) != sorted(existing_ids):
            return False

        with self.session_factory() as session:
            for offset, rule_id in enumerate(ordered_rule_ids):
                stored = session.get(LLMInstructionRuleRecord, rule_id)
                if stored is not None:
                    stored.priority = (offset + 1) * 1000
            session.commit()
        return True

    def reorder_priority_rules(self, ordered_rule_keys: list[str]) -> bool:
        filter_ids = [rule.id for rule in self.list_filter_rules()]
        instruction_ids = [rule.id for rule in self.list_llm_instruction_rules()]
        existing_keys = {
            *(f"filter:{rule_id}" for rule_id in filter_ids),
            *(f"instruction:{rule_id}" for rule_id in instruction_ids),
        }
        if set(ordered_rule_keys) != existing_keys or len(ordered_rule_keys) != len(existing_keys):
            return False

        with self.session_factory() as session:
            for offset, key in enumerate(ordered_rule_keys):
                kind, raw_id = key.split(":", 1)
                rule_id = int(raw_id)
                model = FilterRuleRecord if kind == "filter" else LLMInstructionRuleRecord
                stored = session.get(model, rule_id)
                if stored is not None:
                    stored.priority = (offset + 1) * 1000
            session.commit()
        return True

    def set_llm_instruction_rule_enabled(self, rule_id: int, enabled: bool) -> bool:
        with self.session_factory() as session:
            rule = session.get(LLMInstructionRuleRecord, rule_id)
            if rule is None:
                return False
            rule.enabled = enabled
            session.commit()
            return True

    def delete_llm_instruction_rule(self, rule_id: int) -> bool:
        with self.session_factory() as session:
            rule = session.get(LLMInstructionRuleRecord, rule_id)
            if rule is None:
                return False
            session.delete(rule)
            session.commit()
            return True

    def should_skip_analysis(self, message: EmailMessage) -> bool:
        return self.get_filter_action(message) == FILTER_ACTION_SKIP_ANALYSIS

    def get_filter_action(self, message: EmailMessage) -> str | None:
        decision = self.get_filter_decision(message)
        return decision["action"] if decision else None

    def get_filter_decision(self, message: EmailMessage) -> dict | None:
        with self.session_factory() as session:
            rules = session.scalars(
                select(FilterRuleRecord).where(
                    FilterRuleRecord.enabled.is_(True)
                )
            ).all()

        matching_rules = [rule for rule in rules if _conditional_rule_matches(rule, message)]
        if not matching_rules:
            return None

        matching_rules.sort(
            key=lambda rule: (
                rule.priority,
                -FILTER_ACTION_PRECEDENCE.get(rule.action, 0),
                rule.id,
            )
        )
        rule = matching_rules[0]
        return {
            "action": rule.action,
            "priority": rule.priority,
            "matched_rule_id": rule.id,
            "matched_rule_name": rule.name,
            "reason": _describe_rule(rule),
            "rule_snapshot": _rule_snapshot(rule),
            "preset_analysis": _preset_analysis(rule),
        }

    def get_highest_llm_instruction_priority(self, message: EmailMessage) -> int | None:
        with self.session_factory() as session:
            rules = session.scalars(
                select(LLMInstructionRuleRecord).where(
                    LLMInstructionRuleRecord.enabled.is_(True)
                )
            ).all()

        matching_priorities = [
            rule.priority for rule in rules if _conditional_rule_matches(rule, message)
        ]
        if not matching_priorities:
            return None
        return min(matching_priorities)

    def close(self) -> None:
        self.engine.dispose()


def _normalize_email_address(value: str) -> str:
    return parseaddr(value)[1].strip().lower()


def _describe_rule(rule: FilterRuleRecord) -> str:
    if rule.name:
        return f"Matched filter rule: {rule.name}"
    return f"Matched filter rule #{rule.id}"


def _rule_snapshot(rule: FilterRuleRecord) -> dict:
    return {
        "id": rule.id,
        "name": rule.name,
        "action": rule.action,
        "priority": rule.priority,
        "preset_priority": rule.preset_priority,
        "preset_category": rule.preset_category,
        "preset_summary_ja": rule.preset_summary_ja,
        "preset_suggested_action_ja": rule.preset_suggested_action_ja,
        "preset_requires_reply": rule.preset_requires_reply,
        "preset_reply_recommended": rule.preset_reply_recommended,
        "from_query": rule.from_query,
        "to_query": rule.to_query,
        "delivered_to_query": rule.delivered_to_query,
        "cc_query": rule.cc_query,
        "subject_query": rule.subject_query,
        "has_words": rule.has_words,
        "doesnt_have": rule.doesnt_have,
        "has_attachment": rule.has_attachment,
        "filename_query": rule.filename_query,
        "size_comparison": rule.size_comparison,
        "size_bytes": rule.size_bytes,
    }


def _preset_analysis(rule: FilterRuleRecord) -> dict:
    return {
        "priority": rule.preset_priority or "medium",
        "category": rule.preset_category or "preclassified",
        "summary_ja": rule.preset_summary_ja or "ルールにより自動分類されました。",
        "suggested_action_ja": rule.preset_suggested_action_ja or "",
        "requires_reply": bool(rule.preset_requires_reply),
        "reply_recommended": bool(rule.preset_reply_recommended),
    }


def _email_record_to_message(record: EmailRecord) -> EmailMessage:
    attachments = []
    for item in json.loads(record.attachment_metadata or "[]"):
        attachments.append(
            AttachmentInfo(
                filename=item.get("filename", ""),
                mime_type=item.get("mime_type", ""),
                size=item.get("size"),
                attachment_id=item.get("attachment_id"),
            )
        )

    return EmailMessage(
        gmail_message_id=record.gmail_message_id,
        gmail_thread_id=record.gmail_thread_id,
        sender=record.sender,
        recipients=tuple(json.loads(record.recipients or "[]")),
        cc=tuple(json.loads(record.cc or "[]")),
        subject=record.subject,
        received_at=record.received_at,
        snippet=record.snippet,
        body=record.body,
        attachments=tuple(attachments),
        body_html=record.body_html,
        size_estimate=record.size_estimate,
        label_ids=tuple(json.loads(record.label_ids or "[]")),
        sender_candidates=tuple(json.loads(record.sender_candidates or "[]")),
    )


def _email_record_is_sent(record: EmailRecord) -> bool:
    return "SENT" in set(json.loads(record.label_ids or "[]"))


def _conditional_rule_matches(
    rule: FilterRuleRecord | LLMInstructionRuleRecord,
    message: EmailMessage,
) -> bool:
    message_text = " ".join(
        part for part in [message.subject, message.snippet, message.body] if part
    )
    checks = [
        _matches_address_query(rule.from_query, [message.sender]),
        _matches_address_query(rule.to_query, message.recipients),
        _matches_address_query(rule.delivered_to_query, message.recipients),
        _matches_address_query(rule.cc_query, message.cc),
        _matches_text_query(rule.subject_query, message.subject),
        _matches_text_query(rule.has_words, message_text),
        not rule.doesnt_have or not _matches_text_query(rule.doesnt_have, message_text),
        _matches_attachment(rule.has_attachment, message),
        _matches_filename(rule.filename_query, message),
        _matches_size(rule.size_comparison, rule.size_bytes, message.size_estimate),
        _matches_after(rule.after, message.received_at),
        _matches_before(rule.before, message.received_at),
    ]

    # These Gmail-compatible fields are stored now, but need Gmail metadata that
    # Matomail does not fetch yet before they can be evaluated locally.
    unsupported_active_fields = [
        rule.bcc_query,
        rule.gmail_query,
        rule.negated_gmail_query,
        rule.older_than,
        rule.newer_than,
        rule.category,
        rule.label,
    ]
    if any(field for field in unsupported_active_fields):
        return False

    return all(checks)


def _matches_address_query(query: str, addresses: Iterable[str]) -> bool:
    if not query:
        return True
    normalized_query = query.strip().lower()
    normalized_addresses = {
        _normalize_email_address(address) or address.strip().lower()
        for address in addresses
        if address
    }
    if normalized_query.startswith("*@"):
        domain = normalized_query[2:]
        return any(address.endswith(f"@{domain}") for address in normalized_addresses)
    if "@" in normalized_query:
        return normalized_query in normalized_addresses
    return any(normalized_query in address for address in normalized_addresses)


def _matches_text_query(query: str, value: str) -> bool:
    if not query:
        return True
    return query.lower() in value.lower()


def _matches_attachment(
    has_attachment: bool | None,
    message: EmailMessage,
) -> bool:
    if has_attachment is None:
        return True
    return message.has_attachments is has_attachment


def _matches_filename(query: str, message: EmailMessage) -> bool:
    if not query:
        return True
    return any(query.lower() in attachment.filename.lower() for attachment in message.attachments)


def _matches_size(
    comparison: str,
    size_bytes: int | None,
    message_size: int | None,
) -> bool:
    if not comparison or size_bytes is None:
        return True
    if message_size is None:
        return False
    if comparison == "larger":
        return message_size > size_bytes
    if comparison == "smaller":
        return message_size < size_bytes
    return False


def _matches_after(after: datetime | None, received_at: datetime | None) -> bool:
    if after is None:
        return True
    if received_at is None:
        return False
    return _comparable_datetime(received_at) > _comparable_datetime(after)


def _matches_before(before: datetime | None, received_at: datetime | None) -> bool:
    if before is None:
        return True
    if received_at is None:
        return False
    return _comparable_datetime(received_at) < _comparable_datetime(before)


def _comparable_datetime(value: datetime) -> datetime:
    if value.tzinfo is None:
        return value.replace(tzinfo=UTC)
    return value
