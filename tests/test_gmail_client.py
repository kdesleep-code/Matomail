import base64
from types import SimpleNamespace

from matomail import gmail_client
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
        self.attachments_resource = _AttachmentsResource()
        self.list_calls = []
        self.get_calls = []
        self.send_calls = []
        self.modify_calls = []

    def list(self, **kwargs):
        self.list_calls.append(kwargs)
        return _Request({"messages": [{"id": key} for key in self.messages]})

    def get(self, **kwargs):
        self.get_calls.append(kwargs)
        return _Request(self.messages[kwargs["id"]])

    def attachments(self):
        return self.attachments_resource

    def send(self, **kwargs):
        self.send_calls.append(kwargs)
        return _Request({"id": "sent-msg", "threadId": kwargs["body"].get("threadId", "")})

    def modify(self, **kwargs):
        self.modify_calls.append(kwargs)
        return _Request({"id": kwargs["id"], "labelIds": ["STARRED"]})


class _AttachmentsResource:
    def __init__(self):
        self.get_calls = []

    def get(self, **kwargs):
        self.get_calls.append(kwargs)
        return _Request({"data": _b64("attachment bytes")})


class _UsersResource:
    def __init__(self, messages_resource):
        self.messages_resource = messages_resource

    def messages(self):
        return self.messages_resource

    def getProfile(self, **kwargs):
        return _Request({"emailAddress": "reader@example.com"})


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
                "labelIds": ["INBOX", "UNREAD"],
                "internalDate": "1778745600000",
                "snippet": "A short snippet",
                "payload": {
                    "headers": [
                        {"name": "From", "value": "Sender <sender@example.com>"},
                        {"name": "To", "value": "Reader <reader@example.com>"},
                        {"name": "Cc", "value": "cc@example.com"},
                        {"name": "Subject", "value": "Meeting notes"},
                        {"name": "Reply-To", "value": "reply@example.com"},
                    ],
                    "mimeType": "multipart/mixed",
                    "parts": [
                        {
                            "mimeType": "text/plain",
                            "body": {"data": _b64("Hello from Gmail.")},
                        },
                        {
                            "mimeType": "text/html",
                            "body": {"data": _b64("<p>Hello from <b>Gmail</b>.</p>")},
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
    assert message.body_html == "<p>Hello from <b>Gmail</b>.</p>"
    assert message.label_ids == ("INBOX", "UNREAD")
    assert message.sender_candidates == ("sender@example.com", "reply@example.com")
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
            "q": "newer_than:3d -in:sent",
            "maxResults": 5,
            "pageToken": None,
        }
    ]


def test_fetch_recent_sent_messages_uses_sent_query() -> None:
    service = _GmailService({})

    messages = GmailClient(service).fetch_recent_sent_messages(
        lookback_days=3,
        max_results=5,
    )

    assert messages == []
    assert service.messages_resource.list_calls == [
        {
            "userId": "me",
            "q": "newer_than:3d in:sent",
            "maxResults": 5,
            "pageToken": None,
        }
    ]


def test_fetch_recent_messages_stops_before_already_seen_message() -> None:
    service = _GmailService(
        {
            "new": {
                "id": "new",
                "threadId": "thread-new",
                "payload": {"headers": [], "mimeType": "text/plain"},
            },
            "existing": {
                "id": "existing",
                "threadId": "thread-existing",
                "payload": {"headers": [], "mimeType": "text/plain"},
            },
            "old": {
                "id": "old",
                "threadId": "thread-old",
                "payload": {"headers": [], "mimeType": "text/plain"},
            },
        }
    )

    messages = GmailClient(service).fetch_recent_messages(
        max_results=10,
        stop_when_message_id_seen=lambda message_id: message_id == "existing",
    )

    assert [message.gmail_message_id for message in messages] == ["new"]
    assert [call["id"] for call in service.messages_resource.get_calls] == ["new"]


def test_get_profile_email_returns_authenticated_account() -> None:
    service = _GmailService({})

    assert GmailClient(service).get_profile_email() == "reader@example.com"


def test_download_attachment_returns_decoded_bytes() -> None:
    service = _GmailService({})

    data = GmailClient(service).download_attachment("msg-1", "att-1")

    assert data == b"attachment bytes"
    assert service.messages_resource.attachments_resource.get_calls == [
        {"userId": "me", "messageId": "msg-1", "id": "att-1"}
    ]


def test_star_message_adds_starred_label() -> None:
    service = _GmailService({})

    response = GmailClient(service).star_message("msg-1")

    assert response == {"id": "msg-1", "labelIds": ["STARRED"]}
    assert service.messages_resource.modify_calls == [
        {"userId": "me", "id": "msg-1", "body": {"addLabelIds": ["STARRED"]}}
    ]


def test_send_reply_sends_raw_message_in_thread() -> None:
    service = _GmailService(
        {
            "msg-1": {
                "id": "msg-1",
                "threadId": "thread-1",
                "payload": {
                    "headers": [
                        {"name": "Message-ID", "value": "<original@example.com>"},
                        {"name": "References", "value": "<older@example.com>"},
                    ]
                },
            }
        }
    )

    response = GmailClient(service).send_reply(
        to=("sender@example.com",),
        cc=("cc@example.com",),
        bcc=("bcc@example.com",),
        subject="Meeting notes",
        body="Thank you.",
        thread_id="thread-1",
        source_message_id="msg-1",
    )

    assert response == {"id": "sent-msg", "threadId": "thread-1"}
    call = service.messages_resource.send_calls[0]
    assert call["userId"] == "me"
    assert call["body"]["threadId"] == "thread-1"
    raw_value = call["body"]["raw"]
    raw = base64.urlsafe_b64decode(raw_value + "=" * (-len(raw_value) % 4)).decode(
        "utf-8"
    )
    assert "From: reader@example.com" in raw
    assert "To: sender@example.com" in raw
    assert "Cc: cc@example.com" in raw
    assert "Bcc: bcc@example.com" in raw
    assert "Subject: Re: Meeting notes" in raw
    assert "In-Reply-To: <original@example.com>" in raw
    assert "References: <older@example.com> <original@example.com>" in raw
    assert "Thank you." in raw
    assert service.messages_resource.get_calls[0] == {
        "userId": "me",
        "id": "msg-1",
        "format": "metadata",
        "metadataHeaders": ["Message-ID", "References"],
    }


def test_send_reply_can_encode_sender_display_name() -> None:
    service = _GmailService(
        {
            "msg-1": {
                "id": "msg-1",
                "threadId": "thread-1",
                "payload": {"headers": [{"name": "Message-ID", "value": "<original@example.com>"}]},
            }
        }
    )

    GmailClient(service, sender_name="堀江 一正").send_reply(
        to=("sender@example.com",),
        subject="Meeting notes",
        body="Thank you.",
        thread_id="thread-1",
        source_message_id="msg-1",
    )

    raw_value = service.messages_resource.send_calls[0]["body"]["raw"]
    raw = base64.urlsafe_b64decode(raw_value + "=" * (-len(raw_value) % 4)).decode(
        "utf-8"
    )
    assert "From: =?utf-8?" in raw
    assert "<reader@example.com>" in raw


def test_send_reply_omits_thread_id_when_source_message_id_is_missing() -> None:
    service = _GmailService(
        {
            "msg-1": {
                "id": "msg-1",
                "threadId": "thread-1",
                "payload": {"headers": []},
            }
        }
    )

    GmailClient(service).send_reply(
        to=("sender@example.com",),
        subject="Meeting notes",
        body="Thank you.",
        thread_id="thread-1",
        source_message_id="msg-1",
    )

    call = service.messages_resource.send_calls[0]
    assert "threadId" not in call["body"]


def test_load_credentials_reauths_when_existing_token_lacks_requested_scope(
    tmp_path, monkeypatch
) -> None:
    token_file = tmp_path / "token.json"
    secrets_file = tmp_path / "credentials.json"
    token_file.write_text(
        f'{{"scopes": ["{gmail_client.GMAIL_READONLY_SCOPE}"]}}',
        encoding="utf-8",
    )
    secrets_file.write_text("{}", encoding="utf-8")
    requested_scopes = [
        gmail_client.GMAIL_READONLY_SCOPE,
        gmail_client.GMAIL_SEND_SCOPE,
    ]
    calls = SimpleNamespace(read_scopes=None, flow_scopes=None, prompt=None)

    class _OldCredentials:
        expired = False
        refresh_token = "refresh"
        valid = True

        def has_scopes(self, scopes):
            calls.read_scopes = scopes
            return False

    class _NewCredentials:
        expired = False
        refresh_token = "refresh"
        valid = True

        def has_scopes(self, scopes):
            return True

        def to_json(self):
            return "{}"

    class _Flow:
        def run_local_server(self, port, prompt=None):
            assert port == 8080
            calls.prompt = prompt
            return _NewCredentials()

    monkeypatch.setattr(
        gmail_client.Credentials,
        "from_authorized_user_file",
        staticmethod(lambda filename: _OldCredentials()),
    )
    monkeypatch.setattr(
        gmail_client.InstalledAppFlow,
        "from_client_secrets_file",
        staticmethod(
            lambda filename, scopes: setattr(calls, "flow_scopes", scopes) or _Flow()
        ),
    )

    credentials = gmail_client.load_or_create_credentials(
        token_file=token_file,
        client_secrets_file=secrets_file,
        scopes=requested_scopes,
    )

    assert isinstance(credentials, _NewCredentials)
    assert calls.read_scopes is None
    assert set(requested_scopes).issubset(calls.flow_scopes)
    assert gmail_client.GMAIL_MODIFY_SCOPE in calls.flow_scopes
    assert gmail_client.GOOGLE_CALENDAR_EVENTS_SCOPE in calls.flow_scopes
    assert calls.prompt == "consent"


def test_load_credentials_reauths_when_token_format_is_invalid(
    tmp_path, monkeypatch
) -> None:
    token_file = tmp_path / "token.json"
    secrets_file = tmp_path / "credentials.json"
    token_file.write_text("{}", encoding="utf-8")
    secrets_file.write_text("{}", encoding="utf-8")
    requested_scopes = [
        gmail_client.GMAIL_READONLY_SCOPE,
        gmail_client.GMAIL_MODIFY_SCOPE,
    ]
    calls = SimpleNamespace(flow_scopes=None, prompt=None)

    class _NewCredentials:
        expired = False
        refresh_token = "refresh"
        valid = True

        def has_scopes(self, scopes):
            return True

        def to_json(self):
            return "{}"

    class _Flow:
        def run_local_server(self, port, prompt=None):
            assert port == 8080
            calls.prompt = prompt
            return _NewCredentials()

    def _raise_invalid_token(filename):
        raise ValueError(
            "Authorized user info was not in the expected format, missing fields refresh_token."
        )

    monkeypatch.setattr(
        gmail_client.Credentials,
        "from_authorized_user_file",
        staticmethod(_raise_invalid_token),
    )
    monkeypatch.setattr(
        gmail_client.InstalledAppFlow,
        "from_client_secrets_file",
        staticmethod(
            lambda filename, scopes: setattr(calls, "flow_scopes", scopes) or _Flow()
        ),
    )

    credentials = gmail_client.load_or_create_credentials(
        token_file=token_file,
        client_secrets_file=secrets_file,
        scopes=requested_scopes,
    )

    assert isinstance(credentials, _NewCredentials)
    assert set(requested_scopes).issubset(calls.flow_scopes)
    assert gmail_client.GMAIL_SEND_SCOPE in calls.flow_scopes
    assert gmail_client.GOOGLE_CALENDAR_EVENTS_SCOPE in calls.flow_scopes
    assert calls.prompt == "consent"


def test_load_credentials_reauths_when_loaded_token_lacks_refresh_token(
    tmp_path, monkeypatch
) -> None:
    token_file = tmp_path / "token.json"
    secrets_file = tmp_path / "credentials.json"
    token_file.write_text(
        f'{{"scopes": ["{gmail_client.GMAIL_READONLY_SCOPE}"]}}',
        encoding="utf-8",
    )
    secrets_file.write_text("{}", encoding="utf-8")
    requested_scopes = [gmail_client.GMAIL_READONLY_SCOPE]
    calls = SimpleNamespace(prompt=None)

    class _OldCredentials:
        expired = False
        refresh_token = None
        valid = True

        def has_scopes(self, scopes):
            return True

    class _NewCredentials:
        expired = False
        refresh_token = "refresh"
        valid = True

        def has_scopes(self, scopes):
            return True

        def to_json(self):
            return "{}"

    class _Flow:
        def run_local_server(self, port, prompt=None):
            calls.prompt = prompt
            return _NewCredentials()

    monkeypatch.setattr(
        gmail_client.Credentials,
        "from_authorized_user_file",
        staticmethod(lambda filename: _OldCredentials()),
    )
    monkeypatch.setattr(
        gmail_client.InstalledAppFlow,
        "from_client_secrets_file",
        staticmethod(lambda filename, scopes: _Flow()),
    )

    credentials = gmail_client.load_or_create_credentials(
        token_file=token_file,
        client_secrets_file=secrets_file,
        scopes=requested_scopes,
    )

    assert isinstance(credentials, _NewCredentials)
    assert calls.prompt == "consent"


def test_token_has_scopes_accepts_space_separated_scope(tmp_path) -> None:
    token_file = tmp_path / "token.json"
    token_file.write_text(
        '{"scope": "https://www.googleapis.com/auth/gmail.readonly https://www.googleapis.com/auth/gmail.send"}',
        encoding="utf-8",
    )

    assert gmail_client._token_has_scopes(
        token_file,
        [gmail_client.GMAIL_READONLY_SCOPE, gmail_client.GMAIL_SEND_SCOPE],
    )


def test_should_force_consent_for_send_scope_when_token_is_stale(tmp_path) -> None:
    token_file = tmp_path / "token.json"
    token_file.write_text(
        f'{{"scopes": ["{gmail_client.GMAIL_READONLY_SCOPE}"]}}',
        encoding="utf-8",
    )

    assert gmail_client._should_force_consent(
        token_file,
        [gmail_client.GMAIL_READONLY_SCOPE, gmail_client.GMAIL_SEND_SCOPE],
    )


def test_sender_candidates_use_reply_to_when_from_is_missing() -> None:
    service = _GmailService(
        {
            "msg-1": {
                "id": "msg-1",
                "threadId": "thread-1",
                "payload": {
                    "headers": [
                        {"name": "Reply-To", "value": "list@example.com"},
                        {"name": "Subject", "value": "List mail"},
                    ],
                    "mimeType": "text/plain",
                    "body": {"data": _b64("Hello from list.")},
                },
            }
        }
    )

    message = GmailClient(service).fetch_recent_messages()[0]

    assert message.sender == ""
    assert message.sender_candidates == ("list@example.com",)
