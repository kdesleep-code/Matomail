"""Gmail API integration."""

from __future__ import annotations

import base64
from datetime import UTC, datetime
from email.utils import getaddresses, parsedate_to_datetime
from pathlib import Path
from typing import Any

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

from .config import Settings
from .models import AttachmentInfo, EmailMessage


GMAIL_READONLY_SCOPE = "https://www.googleapis.com/auth/gmail.readonly"


class GmailClient:
    """Fetches and sends Gmail messages."""

    def __init__(self, service: Any, user_id: str = "me") -> None:
        self._service = service
        self._user_id = user_id

    @classmethod
    def from_oauth(
        cls,
        settings: Settings | None = None,
        scopes: list[str] | None = None,
    ) -> "GmailClient":
        settings = settings or Settings()
        scopes = scopes or [GMAIL_READONLY_SCOPE]
        credentials = load_or_create_credentials(
            token_file=settings.google_token_file,
            client_secrets_file=settings.google_client_secrets_file,
            scopes=scopes,
            port=settings.google_oauth_port,
        )
        service = build("gmail", "v1", credentials=credentials)
        return cls(service)

    def fetch_recent_messages(
        self,
        lookback_days: int = 7,
        max_results: int = 30,
    ) -> list[EmailMessage]:
        message_refs = self._list_message_refs(
            query=f"newer_than:{lookback_days}d",
            max_results=max_results,
        )
        return [self._get_message(message_ref["id"]) for message_ref in message_refs]

    def _list_message_refs(self, query: str, max_results: int) -> list[dict[str, str]]:
        refs: list[dict[str, str]] = []
        page_token: str | None = None

        while len(refs) < max_results:
            response = (
                self._service.users()
                .messages()
                .list(
                    userId=self._user_id,
                    q=query,
                    maxResults=min(100, max_results - len(refs)),
                    pageToken=page_token,
                )
                .execute()
            )
            refs.extend(response.get("messages", []))
            page_token = response.get("nextPageToken")
            if not page_token:
                break

        return refs[:max_results]

    def _get_message(self, message_id: str) -> EmailMessage:
        raw_message = (
            self._service.users()
            .messages()
            .get(userId=self._user_id, id=message_id, format="full")
            .execute()
        )
        payload = raw_message.get("payload", {})
        headers = _headers_by_name(payload.get("headers", []))
        attachments: list[AttachmentInfo] = []
        body_parts: list[str] = []

        _collect_payload_content(payload, body_parts, attachments)

        return EmailMessage(
            gmail_message_id=raw_message["id"],
            gmail_thread_id=raw_message.get("threadId", ""),
            sender=headers.get("from", ""),
            recipients=_parse_addresses(headers.get("to", "")),
            cc=_parse_addresses(headers.get("cc", "")),
            subject=headers.get("subject", ""),
            received_at=_parse_received_at(raw_message, headers),
            snippet=raw_message.get("snippet", ""),
            body="\n\n".join(part for part in body_parts if part).strip(),
            attachments=tuple(attachments),
            size_estimate=raw_message.get("sizeEstimate"),
        )


def load_or_create_credentials(
    token_file: Path,
    client_secrets_file: Path,
    scopes: list[str],
    port: int = 8080,
) -> Credentials:
    credentials: Credentials | None = None
    if token_file.exists():
        credentials = Credentials.from_authorized_user_file(str(token_file), scopes)

    if credentials and credentials.expired and credentials.refresh_token:
        credentials.refresh(Request())

    if not credentials or not credentials.valid:
        if not client_secrets_file.exists():
            raise FileNotFoundError(
                f"Google OAuth client secrets file was not found: {client_secrets_file}"
            )
        flow = InstalledAppFlow.from_client_secrets_file(str(client_secrets_file), scopes)
        credentials = flow.run_local_server(port=port)

    token_file.parent.mkdir(parents=True, exist_ok=True)
    token_file.write_text(credentials.to_json(), encoding="utf-8")
    return credentials


def _headers_by_name(headers: list[dict[str, str]]) -> dict[str, str]:
    return {
        header.get("name", "").lower(): header.get("value", "")
        for header in headers
        if header.get("name")
    }


def _parse_addresses(value: str) -> tuple[str, ...]:
    return tuple(address for _, address in getaddresses([value]) if address)


def _parse_received_at(
    raw_message: dict[str, Any],
    headers: dict[str, str],
) -> datetime | None:
    internal_date = raw_message.get("internalDate")
    if internal_date:
        return datetime.fromtimestamp(int(internal_date) / 1000, tz=UTC)

    date_header = headers.get("date")
    if not date_header:
        return None
    parsed = parsedate_to_datetime(date_header)
    if parsed.tzinfo is None:
        return parsed.replace(tzinfo=UTC)
    return parsed


def _collect_payload_content(
    payload: dict[str, Any],
    body_parts: list[str],
    attachments: list[AttachmentInfo],
) -> None:
    filename = payload.get("filename") or ""
    body = payload.get("body", {})
    mime_type = payload.get("mimeType", "")

    if filename or body.get("attachmentId"):
        attachments.append(
            AttachmentInfo(
                filename=filename,
                mime_type=mime_type,
                size=body.get("size"),
                attachment_id=body.get("attachmentId"),
            )
        )

    if mime_type == "text/plain" and body.get("data"):
        body_parts.append(_decode_base64url(body["data"]))

    for part in payload.get("parts", []):
        _collect_payload_content(part, body_parts, attachments)


def _decode_base64url(value: str) -> str:
    padding = "=" * (-len(value) % 4)
    return base64.urlsafe_b64decode(value + padding).decode("utf-8", errors="replace")
