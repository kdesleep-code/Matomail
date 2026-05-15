import base64

from matomail.gmail_client import GmailClient


def _b64(value: str) -> str:
    return base64.urlsafe_b64encode(value.encode("utf-8")).decode("ascii").rstrip("=")


class _Request:
    def __init__(self, response):
        self.response = response

    def execute(self):
        return self.response


class _MessagesResource:
    def __init__(self, messages):
        self.messages = messages
        self.list_calls = []
        self.get_calls = []

    def list(self, **kwargs):
        self.list_calls.append(kwargs)
        return _Request({"messages": [{"id": key} for key in self.messages]})

    def get(self, **kwargs):
        self.get_calls.append(kwargs)
        return _Request(self.messages[kwargs["id"]])


class _UsersResource:
    def __init__(self, messages_resource):
        self.messages_resource = messages_resource

    def messages(self):
        return self.messages_resource


class _GmailService:
    def __init__(self, messages):
        self.messages_resource = _MessagesResource(messages)

    def users(self):
        return _UsersResource(self.messages_resource)


def test_fetch_recent_messages_returns_structured_email_objects() -> None:
    service = _GmailService(
        {
            "msg-1": {
                "id": "msg-1",
                "threadId": "thread-1",
                "internalDate": "1778745600000",
                "snippet": "A short snippet",
                "payload": {
                    "headers": [
                        {"name": "From", "value": "Sender <sender@example.com>"},
                        {"name": "To", "value": "Reader <reader@example.com>"},
                        {"name": "Cc", "value": "cc@example.com"},
                        {"name": "Subject", "value": "Meeting notes"},
                    ],
                    "mimeType": "multipart/mixed",
                    "parts": [
                        {
                            "mimeType": "text/plain",
                            "body": {"data": _b64("Hello from Gmail.")},
                        },
                        {
                            "filename": "agenda.pdf",
                            "mimeType": "application/pdf",
                            "body": {"attachmentId": "att-1", "size": 1234},
                        },
                    ],
                },
            }
        }
    )

    messages = GmailClient(service).fetch_recent_messages(
        lookback_days=7,
        max_results=10,
    )

    assert len(messages) == 1
    message = messages[0]
    assert message.gmail_message_id == "msg-1"
    assert message.gmail_thread_id == "thread-1"
    assert message.sender == "Sender <sender@example.com>"
    assert message.recipients == ("reader@example.com",)
    assert message.cc == ("cc@example.com",)
    assert message.subject == "Meeting notes"
    assert message.snippet == "A short snippet"
    assert message.body == "Hello from Gmail."
    assert message.has_attachments is True
    assert message.attachments[0].filename == "agenda.pdf"
    assert message.attachments[0].mime_type == "application/pdf"
    assert message.attachments[0].attachment_id == "att-1"
    assert message.attachments[0].size == 1234


def test_fetch_recent_messages_uses_gmail_query_and_limit() -> None:
    service = _GmailService({})

    messages = GmailClient(service).fetch_recent_messages(
        lookback_days=3,
        max_results=5,
    )

    assert messages == []
    assert service.messages_resource.list_calls == [
        {
            "userId": "me",
            "q": "newer_than:3d",
            "maxResults": 5,
            "pageToken": None,
        }
    ]

