"""Google Calendar integration."""

from __future__ import annotations

from typing import Any

from googleapiclient.discovery import build

from .config import Settings
from .gmail_client import load_or_create_credentials


CALENDAR_EVENTS_SCOPE = "https://www.googleapis.com/auth/calendar.events"


class CalendarClient:
    """Creates calendar events after user confirmation."""

    def __init__(self, service: Any, calendar_id: str = "primary") -> None:
        self._service = service
        self._calendar_id = calendar_id or "primary"

    @classmethod
    def from_oauth(
        cls,
        settings: Settings | None = None,
        scopes: list[str] | None = None,
        force_consent: bool = False,
    ) -> "CalendarClient":
        settings = settings or Settings()
        scopes = scopes or [CALENDAR_EVENTS_SCOPE]
        credentials = load_or_create_credentials(
            token_file=settings.google_token_file,
            client_secrets_file=settings.google_client_secrets_file,
            scopes=scopes,
            port=settings.google_oauth_port,
            force_consent=force_consent,
        )
        service = build("calendar", "v3", credentials=credentials)
        return cls(service, calendar_id=getattr(settings, "calendar_id", "primary"))

    def create_event(
        self,
        *,
        title: str,
        start_time: str,
        end_time: str,
        timezone: str,
        location: str = "",
        attendees: tuple[str, ...] = (),
        description: str = "",
    ) -> dict[str, Any]:
        body: dict[str, Any] = {
            "summary": title,
            "start": {"dateTime": start_time, "timeZone": timezone},
            "end": {"dateTime": end_time, "timeZone": timezone},
        }
        if location:
            body["location"] = location
        if description:
            body["description"] = description
        if attendees:
            body["attendees"] = [{"email": attendee} for attendee in attendees if attendee]
        return (
            self._service.events()
            .insert(calendarId=self._calendar_id, body=body)
            .execute()
        )

    def delete_event(self, event_id: str) -> None:
        (
            self._service.events()
            .delete(calendarId=self._calendar_id, eventId=event_id)
            .execute()
        )
