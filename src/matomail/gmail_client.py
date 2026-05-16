"""Gmail API integration."""

from __future__ import annotations

import base64
import json
from datetime import UTC, datetime
from email.headerregistry import Address
from email.message import EmailMessage as MimeEmailMessage
from email.utils import getaddresses, parsedate_to_datetime
from pathlib import Path
from typing import Any, Callable

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

from .config import Settings
from .models import AttachmentInfo, EmailMessage


GMAIL_READONLY_SCOPE = "https://www.googleapis.com/auth/gmail.readonly"
GMAIL_MODIFY_SCOPE = "https://www.googleapis.com/auth/gmail.modify"
GMAIL_SEND_SCOPE = "https://www.googleapis.com/auth/gmail.send"


class GmailClient:
    """Fetches and sends Gmail messages."""

    def __init__(self, service: Any, user_id: str = "me", sender_name: str = "") -> None:
        self._service = service
        self._user_id = user_id
        self._sender_name = sender_name.strip()

    @classmethod
    def from_oauth(
        cls,
        settings: Settings | None = None,
        scopes: list[str] | None = None,
        force_consent: bool = False,
    ) -> "GmailClient":
        settings = settings or Settings()
        scopes = scopes or [GMAIL_READONLY_SCOPE]
        credentials = load_or_create_credentials(
            token_file=settings.google_token_file,
            client_secrets_file=settings.google_client_secrets_file,
            scopes=scopes,
            port=settings.google_oauth_port,
            force_consent=force_consent,
        )
        service = build("gmail", "v1", credentials=credentials)
        return cls(service, sender_name=settings.gmail_sender_name)

    def fetch_recent_messages(
        self,
        lookback_days: int = 7,
        max_results: int = 30,
        stop_when_message_id_seen: Callable[[str], bool] | None = None,
    ) -> list[EmailMessage]:
        message_refs = self._list_message_refs(
            query=f"newer_than:{lookback_days}d -in:sent",
            max_results=max_results,
            stop_when_message_id_seen=stop_when_message_id_seen,
        )
        return [self._get_message(message_ref["id"]) for message_ref in message_refs]

    def fetch_recent_sent_messages(
        self,
        lookback_days: int = 7,
        max_results: int = 30,
        stop_when_message_id_seen: Callable[[str], bool] | None = None,
    ) -> list[EmailMessage]:
        message_refs = self._list_message_refs(
            query=f"newer_than:{lookback_days}d in:sent",
            max_results=max_results,
            stop_when_message_id_seen=stop_when_message_id_seen,
        )
        return [self._get_message(message_ref["id"]) for message_ref in message_refs]

    def get_profile_email(self) -> str:
        profile = self._service.users().getProfile(userId=self._user_id).execute()
        return str(profile.get("emailAddress", "")).strip().lower()

    def star_message(self, message_id: str) -> dict[str, Any]:
        return (
            self._service.users()
            .messages()
            .modify(
                userId=self._user_id,
                id=message_id,
                body={"addLabelIds": ["STARRED"]},
            )
            .execute()
        )

    def download_attachment(self, message_id: str, attachment_id: str) -> bytes:
        response = (
            self._service.users()
            .messages()
            .attachments()
            .get(userId=self._user_id, messageId=message_id, id=attachment_id)
            .execute()
        )
        return _decode_base64url_bytes(str(response.get("data", "")))

    def get_message_headers(
        self,
        message_id: str,
        header_names: tuple[str, ...],
    ) -> dict[str, str]:
        response = (
            self._service.users()
            .messages()
            .get(
                userId=self._user_id,
                id=message_id,
                format="metadata",
                metadataHeaders=list(header_names),
            )
            .execute()
        )
        return _headers_by_name(response.get("payload", {}).get("headers", []))

    def send_reply(
        self,
        *,
        to: tuple[str, ...],
        cc: tuple[str, ...] = (),
        bcc: tuple[str, ...] = (),
        subject: str,
        body: str,
        thread_id: str = "",
        source_message_id: str = "",
    ) -> dict[str, Any]:
        message = MimeEmailMessage()
        from_address = self.get_profile_email()
        if from_address:
            message["From"] = _format_from_header(from_address, self._sender_name)
        message["To"] = ", ".join(to)
        if cc:
            message["Cc"] = ", ".join(cc)
        if bcc:
            message["Bcc"] = ", ".join(bcc)
        message["Subject"] = _reply_subject(subject)
        source_headers = (
            self.get_message_headers(
                source_message_id,
                ("Message-ID", "References"),
            )
            if source_message_id
            else {}
        )
        source_rfc_message_id = source_headers.get("message-id", "").strip()
        if source_rfc_message_id:
            references = source_headers.get("references", "").strip()
            message["In-Reply-To"] = source_rfc_message_id
            message["References"] = (
                f"{references} {source_rfc_message_id}".strip()
                if references
                else source_rfc_message_id
            )
        message.set_content(body)

        raw = base64.urlsafe_b64encode(message.as_bytes()).decode("ascii").rstrip("=")
        payload: dict[str, Any] = {"raw": raw}
        if thread_id and (not source_message_id or source_rfc_message_id):
            payload["threadId"] = thread_id
        return (
            self._service.users()
            .messages()
            .send(userId=self._user_id, body=payload)
            .execute()
        )

    def send_message(
        self,
        *,
        to: tuple[str, ...],
        cc: tuple[str, ...] = (),
        bcc: tuple[str, ...] = (),
        subject: str,
        body: str,
    ) -> dict[str, Any]:
        message = MimeEmailMessage()
        from_address = self.get_profile_email()
        if from_address:
            message["From"] = _format_from_header(from_address, self._sender_name)
        message["To"] = ", ".join(to)
        if cc:
            message["Cc"] = ", ".join(cc)
        if bcc:
            message["Bcc"] = ", ".join(bcc)
        message["Subject"] = subject
        message.set_content(body)
        raw = base64.urlsafe_b64encode(message.as_bytes()).decode("ascii").rstrip("=")
        return (
            self._service.users()
            .messages()
            .send(userId=self._user_id, body={"raw": raw})
            .execute()
        )

    def _list_message_refs(
        self,
        query: str,
        max_results: int,
        stop_when_message_id_seen: Callable[[str], bool] | None = None,
    ) -> list[dict[str, str]]:
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
            for message_ref in response.get("messages", []):
                message_id = message_ref.get("id", "")
                if stop_when_message_id_seen and stop_when_message_id_seen(message_id):
                    return refs
                refs.append(message_ref)
                if len(refs) >= max_results:
                    return refs[:max_results]
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
        html_body_parts: list[str] = []

        _collect_payload_content(payload, body_parts, html_body_parts, attachments)

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
            body_html="\n\n".join(part for part in html_body_parts if part).strip(),
            size_estimate=raw_message.get("sizeEstimate"),
            label_ids=tuple(raw_message.get("labelIds", [])),
            sender_candidates=_sender_candidates(headers),
        )


def load_or_create_credentials(
    token_file: Path,
    client_secrets_file: Path,
    scopes: list[str],
    port: int = 8080,
    force_consent: bool = False,
) -> Credentials:
    credentials: Credentials | None = None
    force_consent = force_consent or _should_force_consent(token_file, scopes)
    if token_file.exists() and not force_consent:
        try:
            credentials = Credentials.from_authorized_user_file(str(token_file))
        except ValueError:
            credentials = None
            force_consent = True

    if credentials and not _token_has_scopes(token_file, scopes):
        credentials = None
        force_consent = True

    if credentials and not credentials.has_scopes(scopes):
        credentials = None
        force_consent = True

    if credentials and not credentials.refresh_token:
        credentials = None
        force_consent = True

    if credentials and credentials.expired:
        if credentials.refresh_token:
            credentials.refresh(Request())
        else:
            credentials = None
            force_consent = True

    if not credentials or not credentials.valid:
        if not client_secrets_file.exists():
            raise FileNotFoundError(
                f"Google OAuth client secrets file was not found: {client_secrets_file}"
            )
        flow = InstalledAppFlow.from_client_secrets_file(str(client_secrets_file), scopes)
        credentials = flow.run_local_server(
            port=port,
            prompt="consent" if force_consent else None,
        )

    token_file.parent.mkdir(parents=True, exist_ok=True)
    token_file.write_text(credentials.to_json(), encoding="utf-8")
    return credentials


def _should_force_consent(token_file: Path, scopes: list[str]) -> bool:
    if GMAIL_SEND_SCOPE not in set(scopes):
        return False
    if not token_file.exists():
        return True
    return not _token_has_scopes(token_file, scopes)


def _token_has_scopes(token_file: Path, scopes: list[str]) -> bool:
    try:
        payload = json.loads(token_file.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return False
    raw_scopes = payload.get("scopes") or payload.get("scope") or []
    if isinstance(raw_scopes, str):
        granted_scopes = set(raw_scopes.split())
    elif isinstance(raw_scopes, list):
        granted_scopes = {str(scope) for scope in raw_scopes}
    else:
        granted_scopes = set()
    return set(scopes).issubset(granted_scopes)


def _reply_subject(subject: str) -> str:
    stripped = subject.strip()
    if stripped.lower().startswith("re:"):
        return stripped
    return f"Re: {stripped}" if stripped else "Re:"


def _format_from_header(address: str, sender_name: str) -> str | Address:
    if not sender_name:
        return address
    local_part, separator, domain = address.partition("@")
    if not separator:
        return address
    return Address(display_name=sender_name, username=local_part, domain=domain)


def _headers_by_name(headers: list[dict[str, str]]) -> dict[str, str]:
    return {
        header.get("name", "").lower(): header.get("value", "")
        for header in headers
        if header.get("name")
    }


def _parse_addresses(value: str) -> tuple[str, ...]:
    return tuple(address for _, address in getaddresses([value]) if address)


def _sender_candidates(headers: dict[str, str]) -> tuple[str, ...]:
    candidates: list[str] = []
    for header_name in ["from", "reply-to", "sender", "return-path"]:
        for address in _parse_addresses(headers.get(header_name, "")):
            if address and address.lower() not in {item.lower() for item in candidates}:
                candidates.append(address)
    return tuple(candidates)


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
    html_body_parts: list[str],
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
    elif mime_type == "text/html" and body.get("data"):
        html_body_parts.append(_decode_base64url(body["data"]))

    for part in payload.get("parts", []):
        _collect_payload_content(part, body_parts, html_body_parts, attachments)


def _decode_base64url(value: str) -> str:
    padding = "=" * (-len(value) % 4)
    return base64.urlsafe_b64decode(value + padding).decode("utf-8", errors="replace")


def _decode_base64url_bytes(value: str) -> bytes:
    padding = "=" * (-len(value) % 4)
    return base64.urlsafe_b64decode(value + padding)
