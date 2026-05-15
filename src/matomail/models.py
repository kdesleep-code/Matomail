"""Domain model placeholders."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime


@dataclass(frozen=True)
class AttachmentInfo:
    filename: str
    mime_type: str
    size: int | None = None
    attachment_id: str | None = None


@dataclass(frozen=True)
class EmailMessage:
    gmail_message_id: str
    gmail_thread_id: str
    sender: str
    recipients: tuple[str, ...]
    cc: tuple[str, ...]
    subject: str
    received_at: datetime | None
    snippet: str
    body: str
    attachments: tuple[AttachmentInfo, ...]
    size_estimate: int | None = None

    @property
    def has_attachments(self) -> bool:
        return bool(self.attachments)
