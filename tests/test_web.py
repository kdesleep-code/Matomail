from datetime import UTC, datetime
from types import SimpleNamespace

from fastapi.testclient import TestClient
from googleapiclient.errors import HttpError
import pytest

from matomail.database import Database, EmailAnalysisRecord, RulesDatabase
from matomail.models import AttachmentInfo, EmailMessage
from matomail import web
from matomail.web import app
from sqlalchemy import select
from httplib2 import Response


def test_home_page_has_load_button() -> None:
    client = TestClient(app)

    response = client.get("/")

    assert response.status_code == 200
    assert "未対応メール一覧" in response.text
    assert "本日のメールを読み込む" in response.text
    assert 'action="/runs"' in response.text


def test_missing_run_status_returns_done_error() -> None:
    client = TestClient(app)

    response = client.get("/runs/missing/status")

    assert response.status_code == 200
    payload = response.json()
    assert payload["done"] is True
    assert payload["error"] == "run not found"


def test_rule_form_can_create_skip_rule_from_message(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "matomail.sqlite3"
    rules_db_path = tmp_path / "matomail_rules.sqlite3"
    monkeypatch.setattr(
        web,
        "settings",
        SimpleNamespace(
            db_path=db_path,
            rules_db_path=rules_db_path,
            db_max_size_mb=512,
            db_backup_dir=tmp_path / "backups",
            store_email_body=True,
            report_dir=tmp_path / "reports",
            timezone="Asia/Tokyo",
            account_emails=(),
        ),
    )
    database = Database(db_path)
    database.create_all()
    database.save_email(_message())
    client = TestClient(app)

    response = client.post(
        "/rules/new",
        data={
            "message_id": "msg-1",
            "use_from": "on",
            "from_query": "Sender <sender@example.com>",
            "rule_mode": "skip",
        },
    )

    assert response.status_code == 200
    assert "ルールを追加しました" in response.text
    rules_database = RulesDatabase(rules_db_path)
    rules_database.create_all()
    rules = rules_database.list_filter_rules()
    assert len(rules) == 1
    assert rules[0].from_query == "sender@example.com"
    assert rules[0].action == "skip_analysis"


def test_rule_form_can_create_high_priority_subject_rule(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "matomail.sqlite3"
    rules_db_path = tmp_path / "matomail_rules.sqlite3"
    monkeypatch.setattr(
        web,
        "settings",
        SimpleNamespace(
            db_path=db_path,
            rules_db_path=rules_db_path,
            db_max_size_mb=512,
            db_backup_dir=tmp_path / "backups",
            store_email_body=True,
            report_dir=tmp_path / "reports",
            timezone="Asia/Tokyo",
            account_emails=(),
        ),
    )
    database = Database(db_path)
    database.create_all()
    database.save_email(_message())
    client = TestClient(app)

    response = client.post(
        "/rules/new",
        data={
            "message_id": "msg-1",
            "use_subject": "on",
            "subject_query": "Meeting notes",
            "rule_mode": "high",
        },
    )

    assert response.status_code == 200
    rules_database = RulesDatabase(rules_db_path)
    rules_database.create_all()
    rules = rules_database.list_filter_rules()
    assert len(rules) == 1
    assert rules[0].subject_query == "Meeting notes"
    assert rules[0].action == "preclassify"
    assert rules[0].preset_priority == "high"
    with database.session_factory() as session:
        analysis = session.scalar(select(EmailAnalysisRecord))
    assert analysis is not None
    assert analysis.priority == "high"


def test_rule_form_can_create_top_subject_rule(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "matomail.sqlite3"
    rules_db_path = tmp_path / "matomail_rules.sqlite3"
    monkeypatch.setattr(
        web,
        "settings",
        SimpleNamespace(
            db_path=db_path,
            rules_db_path=rules_db_path,
            db_max_size_mb=512,
            db_backup_dir=tmp_path / "backups",
            store_email_body=True,
            report_dir=tmp_path / "reports",
            timezone="Asia/Tokyo",
            account_emails=(),
        ),
    )
    database = Database(db_path)
    database.create_all()
    database.save_email(_message())
    client = TestClient(app)

    response = client.post(
        "/rules/new",
        data={
            "message_id": "msg-1",
            "use_subject": "on",
            "subject_query": "Meeting notes",
            "rule_mode": "top",
        },
    )

    assert response.status_code == 200
    rules_database = RulesDatabase(rules_db_path)
    rules_database.create_all()
    rules = rules_database.list_filter_rules()
    assert len(rules) == 1
    assert rules[0].subject_query == "Meeting notes"
    assert rules[0].action == "preclassify"
    assert rules[0].preset_priority == "top"
    with database.session_factory() as session:
        analysis = session.scalar(select(EmailAnalysisRecord))
    assert analysis is not None
    assert analysis.priority == "top"


@pytest.mark.parametrize(
    ("rule_mode", "expected_priority"),
    [("middle", "medium"), ("low", "low")],
)
def test_rule_form_can_create_middle_and_low_priority_rules(
    tmp_path, monkeypatch, rule_mode: str, expected_priority: str
) -> None:
    db_path = tmp_path / "matomail.sqlite3"
    rules_db_path = tmp_path / "matomail_rules.sqlite3"
    monkeypatch.setattr(
        web,
        "settings",
        SimpleNamespace(
            db_path=db_path,
            rules_db_path=rules_db_path,
            db_max_size_mb=512,
            db_backup_dir=tmp_path / "backups",
            store_email_body=True,
            report_dir=tmp_path / "reports",
            timezone="Asia/Tokyo",
            account_emails=(),
        ),
    )
    database = Database(db_path)
    database.create_all()
    database.save_email(_message())
    client = TestClient(app)

    response = client.post(
        "/rules/new",
        data={
            "message_id": "msg-1",
            "use_subject": "on",
            "subject_query": "Meeting notes",
            "rule_mode": rule_mode,
        },
    )

    assert response.status_code == 200
    rules_database = RulesDatabase(rules_db_path)
    rules_database.create_all()
    rules = rules_database.list_filter_rules()
    assert len(rules) == 1
    assert rules[0].action == "preclassify"
    assert rules[0].preset_priority == expected_priority
    with database.session_factory() as session:
        analysis = session.scalar(select(EmailAnalysisRecord))
    assert analysis is not None
    assert analysis.priority == expected_priority


def test_rule_form_can_create_llm_instruction_rule(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "matomail.sqlite3"
    rules_db_path = tmp_path / "matomail_rules.sqlite3"
    monkeypatch.setattr(
        web,
        "settings",
        SimpleNamespace(
            db_path=db_path,
            rules_db_path=rules_db_path,
            db_max_size_mb=512,
            db_backup_dir=tmp_path / "backups",
            store_email_body=True,
            report_dir=tmp_path / "reports",
            timezone="Asia/Tokyo",
            account_emails=(),
        ),
    )
    database = Database(db_path)
    database.create_all()
    database.save_email(_message())
    client = TestClient(app)

    response = client.post(
        "/rules/new",
        data={
            "message_id": "msg-1",
            "use_subject": "on",
            "subject_query": "Meeting notes",
            "rule_mode": "instruction",
            "instruction": "This sender's meeting notes should be low priority.",
        },
    )

    assert response.status_code == 200
    rules_database = RulesDatabase(rules_db_path)
    rules_database.create_all()
    instructions = rules_database.list_llm_instruction_rules()
    assert len(instructions) == 1
    assert instructions[0].subject_query == "Meeting notes"
    assert instructions[0].instruction == "This sender's meeting notes should be low priority."
    assert rules_database.list_filter_rules() == []


def test_rule_form_prefills_sender_candidate_when_sender_is_empty(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "matomail.sqlite3"
    rules_db_path = tmp_path / "matomail_rules.sqlite3"
    monkeypatch.setattr(
        web,
        "settings",
        SimpleNamespace(
            db_path=db_path,
            rules_db_path=rules_db_path,
            db_max_size_mb=512,
            db_backup_dir=tmp_path / "backups",
            store_email_body=True,
            report_dir=tmp_path / "reports",
            timezone="Asia/Tokyo",
            account_emails=(),
        ),
    )
    database = Database(db_path)
    database.create_all()
    database.save_email(_message(sender="", sender_candidates=("list@example.com",)))
    client = TestClient(app)

    response = client.get("/rules/new?message_id=msg-1")

    assert response.status_code == 200
    assert 'value="list@example.com"' in response.text


def test_rule_form_extracts_sender_address_from_quoted_empty_display_name(
    tmp_path, monkeypatch
) -> None:
    db_path = tmp_path / "matomail.sqlite3"
    rules_db_path = tmp_path / "matomail_rules.sqlite3"
    monkeypatch.setattr(
        web,
        "settings",
        SimpleNamespace(
            db_path=db_path,
            rules_db_path=rules_db_path,
            db_max_size_mb=512,
            db_backup_dir=tmp_path / "backups",
            store_email_body=True,
            report_dir=tmp_path / "reports",
            timezone="Asia/Tokyo",
            account_emails=(),
        ),
    )
    database = Database(db_path)
    database.create_all()
    database.save_email(_message(sender='"" <sender@example.com>'))
    client = TestClient(app)

    response = client.get("/rules/new?message_id=msg-1")

    assert response.status_code == 200
    assert 'value="sender@example.com"' in response.text
    assert 'value="&quot;&quot; &lt;sender@example.com&gt;"' not in response.text


def test_rules_page_lists_edits_deletes_and_reorders_rules(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "matomail.sqlite3"
    rules_db_path = tmp_path / "matomail_rules.sqlite3"
    monkeypatch.setattr(
        web,
        "settings",
        SimpleNamespace(
            db_path=db_path,
            rules_db_path=rules_db_path,
            db_max_size_mb=512,
            db_backup_dir=tmp_path / "backups",
            store_email_body=True,
            report_dir=tmp_path / "reports",
            timezone="Asia/Tokyo",
            account_emails=(),
        ),
    )
    database = Database(db_path)
    database.create_all()
    database.save_email(_message())
    rules_database = RulesDatabase(rules_db_path)
    rules_database.create_all()
    first = rules_database.add_filter_rule(
        action="skip_analysis",
        name="first",
        from_query="first@example.com",
        priority=10,
    )
    second = rules_database.add_filter_rule(
        action="preclassify",
        name="second",
        subject_query="Meeting",
        preset_priority="high",
        priority=20,
    )
    client = TestClient(app)

    response = client.get("/rules")
    assert response.status_code == 200
    assert "first" in response.text
    assert "second" in response.text

    response = client.post(
        f"/rules/{first.id}/edit",
        data={
            "name": "edited",
            "action": "skip_analysis",
            "priority": "30",
            "from_query": "edited@example.com",
            "subject_query": "",
            "has_words": "",
            "note": "memo",
            "enabled": "on",
        },
        follow_redirects=False,
    )
    assert response.status_code == 303
    edited = rules_database.get_filter_rule(first.id)
    assert edited is not None
    assert edited.name == "edited"
    assert edited.from_query == "edited@example.com"

    response = client.post(
        f"/rules/{first.id}/move",
        data={"direction": "down"},
        follow_redirects=False,
    )
    assert response.status_code == 303
    ordered_ids = [rule.id for rule in rules_database.list_filter_rules()]
    assert ordered_ids.index(first.id) > ordered_ids.index(second.id)

    response = client.post(f"/rules/{second.id}/delete", follow_redirects=False)
    assert response.status_code == 303
    assert rules_database.get_filter_rule(second.id) is None


def test_rules_page_reorders_rules_from_drag_payload(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "matomail.sqlite3"
    rules_db_path = tmp_path / "matomail_rules.sqlite3"
    monkeypatch.setattr(
        web,
        "settings",
        SimpleNamespace(
            db_path=db_path,
            rules_db_path=rules_db_path,
            db_max_size_mb=512,
            db_backup_dir=tmp_path / "backups",
            store_email_body=True,
            report_dir=tmp_path / "reports",
            timezone="Asia/Tokyo",
            account_emails=(),
        ),
    )
    database = Database(db_path)
    database.create_all()
    database.save_email(_message())
    rules_database = RulesDatabase(rules_db_path)
    rules_database.create_all()
    first = rules_database.add_filter_rule(
        action="skip_analysis",
        name="first",
        from_query="first@example.com",
    )
    second = rules_database.add_filter_rule(
        action="skip_analysis",
        name="second",
        from_query="second@example.com",
    )
    third = rules_database.add_filter_rule(
        action="skip_analysis",
        name="third",
        from_query="third@example.com",
    )
    client = TestClient(app)

    response = client.post(
        "/rules/reorder",
        json={"rule_ids": [first.id, third.id, second.id]},
    )

    assert response.status_code == 200
    assert response.json() == {"ok": True}
    assert [rule.id for rule in rules_database.list_filter_rules()] == [
        first.id,
        third.id,
        second.id,
    ]

    response = client.get("/rules")
    assert response.status_code == 200
    assert 'draggable="true"' in response.text
    assert "/rules/reorder" in response.text


def test_instruction_pages_create_edit_delete_and_reorder_rules(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "matomail.sqlite3"
    rules_db_path = tmp_path / "matomail_rules.sqlite3"
    monkeypatch.setattr(
        web,
        "settings",
        SimpleNamespace(
            db_path=db_path,
            rules_db_path=rules_db_path,
            db_max_size_mb=512,
            db_backup_dir=tmp_path / "backups",
            store_email_body=True,
            report_dir=tmp_path / "reports",
            timezone="Asia/Tokyo",
            account_emails=(),
        ),
    )
    database = Database(db_path)
    database.create_all()
    rules_database = RulesDatabase(rules_db_path)
    rules_database.create_all()
    first = rules_database.add_llm_instruction_rule(
        name="first",
        instruction="first instruction",
        subject_query="Meeting",
    )
    second = rules_database.add_llm_instruction_rule(
        name="second",
        instruction="second instruction",
        has_words="body",
    )
    client = TestClient(app)

    response = client.get("/rules")
    assert response.status_code == 200
    assert "first instruction" in response.text
    assert "LLM追加指示" in response.text
    assert 'draggable="true"' in response.text

    response = client.post(
        "/instructions/new",
        data={
            "name": "created",
            "instruction": "Prioritize invoices as high.",
            "subject_query": "Invoice",
            "enabled": "on",
        },
        follow_redirects=False,
    )
    assert response.status_code == 303
    created = [
        rule
        for rule in rules_database.list_llm_instruction_rules()
        if rule.name == "created"
    ][0]
    assert created.instruction == "Prioritize invoices as high."
    assert created.subject_query == "Invoice"

    response = client.post(
        f"/instructions/{first.id}/edit",
        data={
            "name": "edited",
            "instruction": "edited instruction",
            "from_query": "Sender <sender@example.com>",
            "to_query": "",
            "subject_query": "Meeting",
            "has_words": "",
            "doesnt_have": "",
            "note": "memo",
            "enabled": "on",
        },
        follow_redirects=False,
    )
    assert response.status_code == 303
    edited = rules_database.get_llm_instruction_rule(first.id)
    assert edited is not None
    assert edited.name == "edited"
    assert edited.instruction == "edited instruction"
    assert edited.from_query == "sender@example.com"

    response = client.post(
        "/rules/reorder",
        json={
            "rule_ids": [
                f"instruction:{first.id}",
                f"instruction:{created.id}",
                f"instruction:{second.id}",
            ]
        },
    )
    assert response.status_code == 200
    assert [rule.id for rule in rules_database.list_llm_instruction_rules()] == [
        first.id,
        created.id,
        second.id,
    ]

    response = client.post(
        f"/instructions/{second.id}/delete",
        follow_redirects=False,
    )
    assert response.status_code == 303
    assert rules_database.get_llm_instruction_rule(second.id) is None


def test_attachment_download_fetches_once_then_uses_cache(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "matomail.sqlite3"
    rules_db_path = tmp_path / "matomail_rules.sqlite3"
    attachment_cache_dir = tmp_path / "attachments"
    monkeypatch.setattr(
        web,
        "settings",
        SimpleNamespace(
            db_path=db_path,
            rules_db_path=rules_db_path,
            db_max_size_mb=512,
            db_backup_dir=tmp_path / "backups",
            attachment_cache_dir=attachment_cache_dir,
            store_email_body=True,
            report_dir=tmp_path / "reports",
            timezone="Asia/Tokyo",
            account_emails=(),
            google_token_file=tmp_path / "token.json",
            google_client_secrets_file=tmp_path / "credentials.json",
            google_oauth_port=8080,
        ),
    )
    database = Database(db_path)
    database.create_all()
    database.save_email(
        EmailMessage(
            gmail_message_id="msg-attachment",
            gmail_thread_id="thread-attachment",
            sender="sender@example.com",
            recipients=("reader@example.com",),
            cc=(),
            subject="Attachment",
            received_at=datetime(2026, 5, 15, 12, 0, tzinfo=UTC),
            snippet="snippet",
            body="body",
            attachments=(
                AttachmentInfo(
                    filename="agenda.pdf",
                    mime_type="application/pdf",
                    size=12,
                    attachment_id="att-1",
                ),
            ),
        )
    )

    class _FakeGmailClient:
        calls = 0

        def download_attachment(self, message_id: str, attachment_id: str) -> bytes:
            type(self).calls += 1
            assert message_id == "msg-attachment"
            assert attachment_id == "att-1"
            return b"cached attachment"

    monkeypatch.setattr(
        web.GmailClient,
        "from_oauth",
        lambda _settings: _FakeGmailClient(),
    )
    client = TestClient(app)

    first = client.get("/attachments/msg-attachment/0")
    second = client.get("/attachments/msg-attachment/0")

    assert first.status_code == 200
    assert first.content == b"cached attachment"
    assert second.status_code == 200
    assert second.content == b"cached attachment"
    assert _FakeGmailClient.calls == 1
    assert list(attachment_cache_dir.rglob("*agenda.pdf"))


def test_reply_draft_endpoint_generates_draft(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "matomail.sqlite3"
    rules_db_path = tmp_path / "matomail_rules.sqlite3"
    monkeypatch.setattr(
        web,
        "settings",
        SimpleNamespace(
            db_path=db_path,
            rules_db_path=rules_db_path,
            db_max_size_mb=512,
            db_backup_dir=tmp_path / "backups",
            store_email_body=True,
            report_dir=tmp_path / "reports",
            timezone="Asia/Tokyo",
            account_emails=(),
        ),
    )
    database = Database(db_path)
    database.create_all()
    database.save_email(_message())
    database.set_app_setting("llm_instruction.reply", "Always write a concise reply.")

    class _FakeLLMClient:
        prompts: list[str] = []

        def generate_text(self, prompt: str) -> str:
            self.prompts.append(prompt)
            return "承知しました。確認して返信いたします。"

    fake_client = _FakeLLMClient()
    monkeypatch.setattr(
        web.LLMClient,
        "from_settings",
        staticmethod(lambda _settings: fake_client),
    )
    client = TestClient(app)

    response = client.post(
        "/reply-drafts",
        json={"message_id": "msg-1", "policy": "来週確認すると伝える"},
    )

    assert response.status_code == 200
    assert response.json()["draft"] == "承知しました。確認して返信いたします。"
    assert "来週確認すると伝える" in fake_client.prompts[0]
    assert "Always write a concise reply." in fake_client.prompts[0]
    assert "Meeting notes" in fake_client.prompts[0]
    assert "body text for rule creation" in fake_client.prompts[0]


def test_reply_draft_endpoint_rejects_sent_mail(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "matomail.sqlite3"
    rules_db_path = tmp_path / "matomail_rules.sqlite3"
    monkeypatch.setattr(
        web,
        "settings",
        SimpleNamespace(
            db_path=db_path,
            rules_db_path=rules_db_path,
            db_max_size_mb=512,
            db_backup_dir=tmp_path / "backups",
            store_email_body=True,
            report_dir=tmp_path / "reports",
            timezone="Asia/Tokyo",
            account_emails=(),
        ),
    )
    database = Database(db_path)
    database.create_all()
    sent = EmailMessage(
        gmail_message_id="sent-1",
        gmail_thread_id="thread-1",
        sender="reader@example.com",
        recipients=("sender@example.com",),
        cc=(),
        subject="Sent reply",
        received_at=datetime(2026, 5, 15, 12, 5, tzinfo=UTC),
        snippet="sent snippet",
        body="sent body",
        attachments=(),
        label_ids=("SENT",),
    )
    database.save_email(sent)
    client = TestClient(app)

    response = client.post(
        "/reply-drafts",
        json={"message_id": "sent-1", "policy": "もう一度送る"},
    )

    assert response.status_code == 400
    assert response.json()["error"] == "reply drafts are only available for received mail"


def test_reply_form_prefills_reply_all_addresses(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "matomail.sqlite3"
    rules_db_path = tmp_path / "matomail_rules.sqlite3"
    monkeypatch.setattr(
        web,
        "settings",
        SimpleNamespace(
            db_path=db_path,
            rules_db_path=rules_db_path,
            db_max_size_mb=512,
            db_backup_dir=tmp_path / "backups",
            store_email_body=True,
            report_dir=tmp_path / "reports",
            timezone="Asia/Tokyo",
            account_emails=("reader@example.com",),
        ),
    )
    database = Database(db_path)
    database.create_all()
    database.save_email(
        EmailMessage(
            gmail_message_id="reply-source",
            gmail_thread_id="thread-1",
            sender="Sender <sender@example.com>",
            recipients=("reader@example.com", "other@example.com"),
            cc=("cc@example.com",),
            subject="Meeting notes",
            received_at=datetime(2026, 5, 15, 12, 0, tzinfo=UTC),
            snippet="snippet",
            body="original body",
            attachments=(),
            sender_candidates=("reply@example.com",),
        )
    )
    client = TestClient(app)

    response = client.get(
        "/replies/new?message_id=reply-source",
        headers={"referer": "http://127.0.0.1:8017/reports/2026-05-15/messages/source.html"},
    )

    assert response.status_code == 200
    assert 'name="to" value="reply@example.com, sender@example.com"' in response.text
    assert 'name="cc" value="other@example.com, cc@example.com"' in response.text
    assert (
        'name="return_url" value="/reports/2026-05-15/messages/source.html"'
        in response.text
    )
    assert "/replies/schedule" not in response.text
    assert "original body" in response.text
    assert "2026" in response.text
    assert "21:00 Sender <sender@example.com>:" in response.text


def test_reply_form_can_prefill_generated_draft_before_original_quote(
    tmp_path, monkeypatch
) -> None:
    db_path = tmp_path / "matomail.sqlite3"
    rules_db_path = tmp_path / "matomail_rules.sqlite3"
    monkeypatch.setattr(
        web,
        "settings",
        SimpleNamespace(
            db_path=db_path,
            rules_db_path=rules_db_path,
            db_max_size_mb=512,
            db_backup_dir=tmp_path / "backups",
            store_email_body=True,
            report_dir=tmp_path / "reports",
            timezone="Asia/Tokyo",
            account_emails=("reader@example.com",),
        ),
    )
    database = Database(db_path)
    database.create_all()
    database.save_email(_message())
    client = TestClient(app)

    response = client.post(
        "/replies/new",
        data={
            "message_id": "msg-1",
            "draft": "自動生成した返信案です。",
            "return_url": "/reports/2026-05-15/messages/msg-1.html",
        },
    )

    assert response.status_code == 200
    assert "自動生成した返信案です。" in response.text
    assert "21:00 Sender <sender@example.com>:" in response.text
    assert "自動生成した返信案です。\n\n\n2026" in response.text
    assert response.text.index("自動生成した返信案です。") < response.text.index("21:00 Sender")
    assert "body text for rule creation" in response.text


def test_send_reply_posts_to_gmail(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "matomail.sqlite3"
    rules_db_path = tmp_path / "matomail_rules.sqlite3"
    monkeypatch.setattr(
        web,
        "settings",
        SimpleNamespace(
            db_path=db_path,
            rules_db_path=rules_db_path,
            db_max_size_mb=512,
            db_backup_dir=tmp_path / "backups",
            store_email_body=True,
            report_dir=tmp_path / "reports",
            timezone="Asia/Tokyo",
            account_emails=(),
            google_token_file=tmp_path / "token.json",
            google_client_secrets_file=tmp_path / "credentials.json",
            google_oauth_port=8080,
        ),
    )
    database = Database(db_path)
    database.create_all()
    database.save_email(_message())

    class _FakeGmailClient:
        calls = []

        def send_reply(self, **kwargs):
            self.calls.append(kwargs)
            return {"id": "sent-1"}

    fake_client = _FakeGmailClient()
    monkeypatch.setattr(
        web.GmailClient,
        "from_oauth",
        lambda _settings, scopes=None, force_consent=False: fake_client,
    )
    client = TestClient(app)

    response = client.post(
        "/replies/send",
        data={
            "message_id": "msg-1",
            "to": "sender@example.com",
            "cc": "cc@example.com",
            "bcc": "bcc@example.com",
            "body": "返信本文です。",
            "return_url": "/reports/2026-05-15/messages/msg-1.html",
        },
        follow_redirects=False,
    )

    assert response.status_code == 303
    assert response.headers["location"] == "/reports/2026-05-15/messages/msg-1.html"
    assert fake_client.calls == [
        {
            "to": ("sender@example.com",),
            "cc": ("cc@example.com",),
            "bcc": ("bcc@example.com",),
            "subject": "Meeting notes",
            "body": "返信本文です。",
            "thread_id": "thread-1",
            "source_message_id": "msg-1",
        }
    ]


def test_send_reply_failure_renders_form_without_bad_gateway(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "matomail.sqlite3"
    rules_db_path = tmp_path / "matomail_rules.sqlite3"
    monkeypatch.setattr(
        web,
        "settings",
        SimpleNamespace(
            db_path=db_path,
            rules_db_path=rules_db_path,
            db_max_size_mb=512,
            db_backup_dir=tmp_path / "backups",
            store_email_body=True,
            report_dir=tmp_path / "reports",
            timezone="Asia/Tokyo",
            account_emails=(),
            google_token_file=tmp_path / "token.json",
            google_client_secrets_file=tmp_path / "credentials.json",
            google_oauth_port=8080,
        ),
    )
    database = Database(db_path)
    database.create_all()
    database.save_email(_message())

    class _FailingGmailClient:
        def send_reply(self, **kwargs):
            response = Response({"status": "400", "reason": "Bad Request"})
            raise HttpError(
                response,
                b'{"error":{"message":"Invalid thread reference"}}',
            )

    monkeypatch.setattr(
        web.GmailClient,
        "from_oauth",
        lambda _settings, scopes=None, force_consent=False: _FailingGmailClient(),
    )
    client = TestClient(app)

    response = client.post(
        "/replies/send",
        data={
            "message_id": "msg-1",
            "to": "sender@example.com",
            "body": "返信本文です。",
        },
    )

    assert response.status_code == 200
    assert "Invalid thread reference" in response.text


def test_send_reply_reauths_once_when_scope_is_insufficient(
    tmp_path, monkeypatch
) -> None:
    db_path = tmp_path / "matomail.sqlite3"
    rules_db_path = tmp_path / "matomail_rules.sqlite3"
    monkeypatch.setattr(
        web,
        "settings",
        SimpleNamespace(
            db_path=db_path,
            rules_db_path=rules_db_path,
            db_max_size_mb=512,
            db_backup_dir=tmp_path / "backups",
            store_email_body=True,
            report_dir=tmp_path / "reports",
            timezone="Asia/Tokyo",
            account_emails=(),
            google_token_file=tmp_path / "token.json",
            google_client_secrets_file=tmp_path / "credentials.json",
            google_oauth_port=8080,
        ),
    )
    database = Database(db_path)
    database.create_all()
    database.save_email(_message())
    force_consent_values: list[bool] = []

    class _ScopeRetryGmailClient:
        def __init__(self, force_consent: bool) -> None:
            self.force_consent = force_consent

        def send_reply(self, **kwargs):
            if not self.force_consent:
                response = Response({"status": "403", "reason": "Forbidden"})
                raise HttpError(
                    response,
                    b'{"error":{"message":"Request had insufficient authentication scopes."}}',
                )
            return {"id": "sent-1"}

    def _from_oauth(_settings, scopes=None, force_consent=False):
        force_consent_values.append(force_consent)
        return _ScopeRetryGmailClient(force_consent)

    monkeypatch.setattr(web.GmailClient, "from_oauth", _from_oauth)
    client = TestClient(app)

    response = client.post(
        "/replies/send",
        data={
            "message_id": "msg-1",
            "to": "sender@example.com",
            "body": "返信本文です。",
            "return_url": "/reports/2026-05-15/messages/msg-1.html",
        },
        follow_redirects=False,
    )

    assert response.status_code == 303
    assert response.headers["location"] == "/reports/2026-05-15/messages/msg-1.html"
    assert force_consent_values == [False, True]


def test_send_new_message_redirects_to_latest_report(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "matomail.sqlite3"
    rules_db_path = tmp_path / "matomail_rules.sqlite3"
    report_dir = tmp_path / "reports"
    report_dir.mkdir()
    (report_dir / "index.html").write_text("<html></html>", encoding="utf-8")
    monkeypatch.setattr(
        web,
        "settings",
        SimpleNamespace(
            db_path=db_path,
            rules_db_path=rules_db_path,
            db_max_size_mb=512,
            db_backup_dir=tmp_path / "backups",
            store_email_body=True,
            report_dir=report_dir,
            timezone="Asia/Tokyo",
            account_emails=(),
            google_token_file=tmp_path / "token.json",
            google_client_secrets_file=tmp_path / "credentials.json",
            google_oauth_port=8080,
        ),
    )
    sent_calls = []

    class _FakeGmailClient:
        def send_message(self, **kwargs):
            sent_calls.append(kwargs)
            return {"id": "new-sent"}

    monkeypatch.setattr(
        web.GmailClient,
        "from_oauth",
        lambda _settings, scopes=None: _FakeGmailClient(),
    )
    client = TestClient(app)

    response = client.post(
        "/compose",
        data={
            "to": "receiver@example.com",
            "cc": "",
            "bcc": "",
            "subject": "Hello",
            "body": "New message body.",
        },
        follow_redirects=False,
    )

    assert response.status_code == 303
    assert response.headers["location"] == "/reports/index.html"
    assert sent_calls == [
        {
            "to": ("receiver@example.com",),
            "cc": (),
            "bcc": (),
            "subject": "Hello",
            "body": "New message body.",
        }
    ]


def test_generation_instruction_pages_edit_compose_and_reply_settings(
    tmp_path, monkeypatch
) -> None:
    db_path = tmp_path / "matomail.sqlite3"
    rules_db_path = tmp_path / "matomail_rules.sqlite3"
    monkeypatch.setattr(
        web,
        "settings",
        SimpleNamespace(
            db_path=db_path,
            rules_db_path=rules_db_path,
            db_max_size_mb=512,
            db_backup_dir=tmp_path / "backups",
            store_email_body=True,
            report_dir=tmp_path / "reports",
            timezone="Asia/Tokyo",
            account_emails=(),
        ),
    )
    database = Database(db_path)
    database.create_all()
    database.set_app_setting("llm_instruction.compose", "Use short paragraphs.")
    client = TestClient(app)

    response = client.get("/generation-instructions")

    assert response.status_code == 200
    assert "新規メールの文面を生成するとき" in response.text
    assert "返信メールの文面を生成するとき" in response.text
    assert "Use short paragraphs." in response.text
    assert "/generation-instructions/reply/edit" in response.text

    response = client.post(
        "/generation-instructions/reply/edit",
        data={"instruction": "Use polite Japanese."},
        follow_redirects=False,
    )

    assert response.status_code == 303
    assert response.headers["location"] == "/generation-instructions"
    assert database.get_app_setting("llm_instruction.reply") == "Use polite Japanese."


def test_unhandled_page_lists_active_high_or_medium_threads(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "matomail.sqlite3"
    rules_db_path = tmp_path / "matomail_rules.sqlite3"
    monkeypatch.setattr(
        web,
        "settings",
        SimpleNamespace(
            db_path=db_path,
            rules_db_path=rules_db_path,
            db_max_size_mb=512,
            db_backup_dir=tmp_path / "backups",
            store_email_body=True,
            report_dir=tmp_path / "reports",
            timezone="Asia/Tokyo",
            account_emails=("me@example.com",),
        ),
    )
    database = Database(db_path)
    database.create_all()
    single_sent = EmailMessage(
        gmail_message_id="single-sent",
        gmail_thread_id="thread-single-sent",
        sender="Me <me@example.com>",
        recipients=("other@example.com",),
        cc=(),
        subject="Single sent",
        received_at=datetime(2026, 5, 15, 12, 0, tzinfo=UTC),
        snippet="single sent snippet",
        body="single sent body",
        attachments=(),
        label_ids=("SENT",),
    )
    paired_sent = EmailMessage(
        gmail_message_id="paired-sent",
        gmail_thread_id="thread-paired",
        sender="Me <me@example.com>",
        recipients=("other@example.com",),
        cc=(),
        subject="Paired sent",
        received_at=datetime(2026, 5, 15, 12, 5, tzinfo=UTC),
        snippet="paired sent snippet",
        body="paired sent body",
        attachments=(),
        label_ids=("SENT",),
    )
    paired_received = EmailMessage(
        gmail_message_id="paired-received",
        gmail_thread_id="thread-paired",
        sender="other@example.com",
        recipients=("me@example.com",),
        cc=(),
        subject="Re: Paired sent",
        received_at=datetime(2026, 5, 15, 12, 10, tzinfo=UTC),
        snippet="paired received snippet",
        body="paired received body",
        attachments=(),
    )
    single_received = _message()
    resolved_received = EmailMessage(
        gmail_message_id="resolved-received",
        gmail_thread_id="thread-resolved",
        sender="other@example.com",
        recipients=("me@example.com",),
        cc=(),
        subject="Resolved received",
        received_at=datetime(2026, 5, 15, 12, 20, tzinfo=UTC),
        snippet="resolved snippet",
        body="resolved body",
        attachments=(),
    )
    database.save_email(single_sent)
    database.save_email(paired_sent)
    database.save_email(paired_received)
    database.save_email(single_received)
    database.save_email(resolved_received)
    for message in [single_sent, paired_sent, paired_received, single_received, resolved_received]:
        database.save_analysis(
            message.gmail_message_id,
            {
                "summary_ja": "",
                "category": "test",
                "priority": "medium",
                "requires_reply": False,
                "suggested_action_ja": "",
                "deadline_candidates": [],
                "meeting_candidates": [],
                "reply_draft_ja": "",
                "confidence": 1.0,
            },
            llm_model="test",
        )
    database.set_resolved("resolved-received", True)
    client = TestClient(app)

    response = client.get("/unhandled")

    assert response.status_code == 200
    assert "未対応メール一覧" in response.text
    assert "Re: Paired sent" in response.text
    assert "Meeting notes" in response.text
    assert "Single sent" not in response.text
    assert "Resolved received" not in response.text

    redirect = client.get("/singletons", follow_redirects=False)
    assert redirect.status_code == 303
    assert redirect.headers["location"] == "/unhandled"

    completed_redirect = client.get("/completed", follow_redirects=False)
    assert completed_redirect.status_code == 303


def test_message_state_endpoints_mark_opened_and_resolved(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "matomail.sqlite3"
    rules_db_path = tmp_path / "matomail_rules.sqlite3"
    monkeypatch.setattr(
        web,
        "settings",
        SimpleNamespace(
            db_path=db_path,
            rules_db_path=rules_db_path,
            db_max_size_mb=512,
            db_backup_dir=tmp_path / "backups",
            store_email_body=True,
            report_dir=tmp_path / "reports",
            timezone="Asia/Tokyo",
            account_emails=(),
        ),
    )
    database = Database(db_path)
    database.create_all()
    database.save_email(_message())
    monkeypatch.setattr(web, "generate_report", lambda _settings, open_browser=False: None)
    client = TestClient(app)

    opened = client.post("/messages/msg-1/opened")
    resolved = client.post("/messages/msg-1/resolved", json={"resolved": True})

    assert opened.status_code == 200
    assert opened.json() == {"ok": True}
    assert resolved.status_code == 200
    assert resolved.json() == {"ok": True}
    with database.session_factory() as session:
        state = session.scalar(
            select(web.ProcessingStateRecord).where(
                web.ProcessingStateRecord.gmail_message_id == "msg-1"
            )
        )
    assert state is not None
    assert state.web_opened is True
    assert state.resolved is True


def test_calendar_candidate_cancel_preserves_cached_candidates(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "matomail.sqlite3"
    rules_db_path = tmp_path / "matomail_rules.sqlite3"
    monkeypatch.setattr(
        web,
        "settings",
        SimpleNamespace(
            db_path=db_path,
            rules_db_path=rules_db_path,
            db_max_size_mb=512,
            db_backup_dir=tmp_path / "backups",
            store_email_body=True,
            report_dir=tmp_path / "reports",
            timezone="Asia/Tokyo",
            account_emails=(),
        ),
    )
    database = Database(db_path)
    database.create_all()
    database.save_email(_message())
    candidates = [
        {
            "title": "Meeting",
            "start_time": "2026-05-20 12:00",
            "end_time": "2026-05-20 13:00",
            "timezone": "Asia/Tokyo",
        }
    ]
    database.set_meeting_candidates("msg-1", candidates)
    monkeypatch.setattr(web, "generate_report", lambda _settings, open_browser=False: None)

    client = TestClient(app)
    clear = client.post(
        "/calendar-candidates/clear",
        data={"message_id": "msg-1", "return_url": "/reports/2026-05-16/messages/msg-1.html"},
        follow_redirects=False,
    )
    assert clear.status_code == 303

    assert database.get_meeting_candidates("msg-1") == candidates
    with database.session_factory() as session:
        state = session.scalar(
            select(web.ProcessingStateRecord).where(
                web.ProcessingStateRecord.gmail_message_id == "msg-1"
            )
        )
    assert state is not None
    assert state.calendar_candidates_hidden is True


def test_calendar_candidate_extract_uses_cache_without_llm(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "matomail.sqlite3"
    rules_db_path = tmp_path / "matomail_rules.sqlite3"
    monkeypatch.setattr(
        web,
        "settings",
        SimpleNamespace(
            db_path=db_path,
            rules_db_path=rules_db_path,
            db_max_size_mb=512,
            db_backup_dir=tmp_path / "backups",
            store_email_body=True,
            report_dir=tmp_path / "reports",
            timezone="Asia/Tokyo",
            account_emails=(),
        ),
    )
    database = Database(db_path)
    database.create_all()
    database.save_email(_message())
    database.set_meeting_candidates(
        "msg-1",
        [{"title": "Cached meeting", "start_time": "2026-05-20 12:00"}],
        hidden=True,
    )
    monkeypatch.setattr(web, "generate_report", lambda _settings, open_browser=False: None)

    class FailingLLMClient:
        @classmethod
        def from_settings(cls, _settings):
            raise AssertionError("LLM should not be called for cached candidates")

    monkeypatch.setattr(web, "LLMClient", FailingLLMClient)
    client = TestClient(app)

    response = client.post(
        "/calendar-candidates/extract",
        data={"message_id": "msg-1", "return_url": "/reports/2026-05-16/messages/msg-1.html"},
        follow_redirects=False,
    )

    assert response.status_code == 303
    with database.session_factory() as session:
        state = session.scalar(
            select(web.ProcessingStateRecord).where(
                web.ProcessingStateRecord.gmail_message_id == "msg-1"
            )
        )
    assert state is not None
    assert state.calendar_candidates_hidden is False


def test_calendar_event_creation_overwrites_cached_candidate_with_submitted_values(
    tmp_path,
    monkeypatch,
) -> None:
    db_path = tmp_path / "matomail.sqlite3"
    rules_db_path = tmp_path / "matomail_rules.sqlite3"
    monkeypatch.setattr(
        web,
        "settings",
        SimpleNamespace(
            db_path=db_path,
            rules_db_path=rules_db_path,
            db_max_size_mb=512,
            db_backup_dir=tmp_path / "backups",
            store_email_body=True,
            report_dir=tmp_path / "reports",
            timezone="Asia/Tokyo",
            account_emails=(),
        ),
    )
    database = Database(db_path)
    database.create_all()
    database.save_email(_message())
    database.set_meeting_candidates(
        "msg-1",
        [{"title": "Old title", "start_time": "2026-05-20 12:00"}],
    )
    monkeypatch.setattr(web, "generate_report", lambda _settings, open_browser=False: None)
    monkeypatch.setattr(
        web,
        "_create_calendar_event_via_google",
        lambda **_kwargs: {"id": "calendar-event-1"},
    )
    client = TestClient(app)

    response = client.post(
        "/calendar-events",
        data={
            "message_id": "msg-1",
            "return_url": "/reports/2026-05-16/messages/msg-1.html",
            "title": "Edited meeting",
            "start_time": "2026-05-21 10:00",
            "end_time": "2026-05-21 11:00",
            "timezone": "Asia/Tokyo",
            "location": "Teams",
            "attendees": "reader@example.com, sender@example.com",
            "description": "Edited description",
        },
        follow_redirects=False,
    )

    assert response.status_code == 303
    cached = database.get_meeting_candidates("msg-1")
    assert cached[0]["title"] == "Edited meeting"
    assert cached[0]["start_time"] == "2026-05-21 10:00"
    assert cached[0]["end_time"] == "2026-05-21 11:00"
    assert cached[0]["location"] == "Teams"
    assert cached[0]["attendees"] == ["reader@example.com", "sender@example.com"]
    assert "Edited description" in cached[0]["description"]
    assert "https://mail.google.com/mail/u/0/#all/thread-1" in cached[0]["description"]


def test_schedule_reply_reports_gmail_api_limitation(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "matomail.sqlite3"
    rules_db_path = tmp_path / "matomail_rules.sqlite3"
    monkeypatch.setattr(
        web,
        "settings",
        SimpleNamespace(
            db_path=db_path,
            rules_db_path=rules_db_path,
            db_max_size_mb=512,
            db_backup_dir=tmp_path / "backups",
            store_email_body=True,
            report_dir=tmp_path / "reports",
            timezone="Asia/Tokyo",
            account_emails=(),
        ),
    )
    database = Database(db_path)
    database.create_all()
    database.save_email(_message())
    client = TestClient(app)

    response = client.post(
        "/replies/schedule/confirm",
        data={
            "message_id": "msg-1",
            "to": "sender@example.com",
            "body": "返信本文です。",
            "scheduled_at": "2026-05-16T10:00",
        },
    )

    assert response.status_code == 501
    assert "Gmail API" in response.text
    assert "送信予約" in response.text


def _message(
    sender: str = "Sender <sender@example.com>",
    sender_candidates: tuple[str, ...] = (),
) -> EmailMessage:
    return EmailMessage(
        gmail_message_id="msg-1",
        gmail_thread_id="thread-1",
        sender=sender,
        recipients=("reader@example.com",),
        cc=(),
        subject="Meeting notes",
        received_at=datetime(2026, 5, 15, 12, 0, tzinfo=UTC),
        snippet="snippet",
        body="body text for rule creation",
        attachments=(),
        sender_candidates=sender_candidates,
    )
