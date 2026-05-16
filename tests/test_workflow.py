from __future__ import annotations

import json
from datetime import UTC, datetime
from types import SimpleNamespace

from matomail.database import (
    FILTER_ACTION_PRECLASSIFY,
    Database,
    EmailAnalysisRecord,
    RulesDatabase,
)
from matomail.models import EmailMessage
from matomail import workflow


def _settings(tmp_path):
    return SimpleNamespace(
        db_path=tmp_path / "matomail.sqlite3",
        rules_db_path=tmp_path / "matomail_rules.sqlite3",
        db_backup_dir=tmp_path / "backups",
        db_max_size_mb=512.0,
        store_email_body=True,
        max_emails_per_run=30,
        lookback_days=7,
        llm_model="test-model",
    )


def _message(
    message_id: str = "msg-1",
    recipients: tuple[str, ...] = ("reader@example.com",),
) -> EmailMessage:
    return EmailMessage(
        gmail_message_id=message_id,
        gmail_thread_id=f"thread-{message_id}",
        sender="JCB Webmaster <mail@qa.jcb.co.jp>",
        recipients=recipients,
        cc=(),
        subject="JCBカード／ショッピングご利用のお知らせ",
        received_at=datetime(2026, 5, 15, 12, 0, tzinfo=UTC),
        snippet="ショッピング利用のお知らせです",
        body="堀江 様\n利用日時 2026-05-15 14:00\n利用金額 1,000円",
        attachments=(),
    )


def _analysis_json(priority: str = "low") -> str:
    return json.dumps(
        {
            "summary_ja": "カード利用通知です。",
            "category": "card",
            "priority": priority,
            "requires_reply": False,
            "suggested_action_ja": "確認のみ。",
            "deadline_candidates": [],
            "meeting_candidates": [],
            "attachment_action_required": False,
            "reply_recommended": False,
            "reply_draft_ja": "",
            "confidence": 0.9,
        },
        ensure_ascii=False,
    )


class _FakeLLMClient:
    def __init__(self) -> None:
        self.prompts: list[str] = []

    def generate_text(self, prompt: str) -> str:
        self.prompts.append(prompt)
        return _analysis_json()


class _FakeGmailClient:
    def __init__(
        self,
        messages: list[EmailMessage],
        sent_messages: list[EmailMessage] | None = None,
    ) -> None:
        self.messages = messages
        self.sent_messages = sent_messages or []
        self.starred_message_ids: list[str] = []

    def get_profile_email(self) -> str:
        return "me@example.com"

    def fetch_recent_messages(self, **_kwargs) -> list[EmailMessage]:
        return self.messages

    def fetch_recent_sent_messages(self, **_kwargs) -> list[EmailMessage]:
        return self.sent_messages

    def star_message(self, message_id: str):
        self.starred_message_ids.append(message_id)
        return {"id": message_id}


def test_duplicate_messages_are_merged_before_llm_input(tmp_path, monkeypatch) -> None:
    settings = _settings(tmp_path)
    mail_database = Database(settings.db_path)
    rules_database = RulesDatabase(settings.rules_db_path)
    mail_database.create_all()
    rules_database.create_all()
    first = _message("msg-1", ("first@example.com",))
    second = _message("msg-2", ("second@example.com",))
    mail_database.save_emails([first, second])
    rules_database.add_llm_instruction_rule(
        instruction="通常利用なら Low と判定してください。",
        from_query="mail@qa.jcb.co.jp",
        subject_query="JCBカード／ショッピングご利用のお知らせ",
    )
    fake_client = _FakeLLMClient()
    monkeypatch.setattr(workflow.LLMClient, "from_settings", lambda _settings: fake_client)

    analyzed_count = workflow.analyze_saved_messages(settings)

    assert analyzed_count == 2
    assert len(fake_client.prompts) == 1
    assert "first@example.com" in fake_client.prompts[0]
    assert "second@example.com" in fake_client.prompts[0]
    assert "通常利用なら Low と判定してください。" in fake_client.prompts[0]
    with mail_database.session_factory() as session:
        assert session.query(EmailAnalysisRecord).count() == 2


def test_lower_number_llm_instruction_rule_keeps_mail_processable(tmp_path, monkeypatch) -> None:
    settings = _settings(tmp_path)
    rules_database = RulesDatabase(settings.rules_db_path)
    rules_database.create_all()
    rules_database.add_filter_rule(
        name="broad high",
        action=FILTER_ACTION_PRECLASSIFY,
        preset_priority="high",
        has_words="堀江",
        priority=1000,
    )
    rules_database.add_llm_instruction_rule(
        instruction="通常利用なら Low と判定してください。",
        from_query="mail@qa.jcb.co.jp",
        subject_query="JCBカード／ショッピングご利用のお知らせ",
        priority=500,
    )
    message = _message()
    monkeypatch.setattr(
        workflow.GmailClient,
        "from_oauth",
        lambda _settings, scopes=None: _FakeGmailClient([message]),
    )

    processable_messages, fetched_count = workflow.fetch_processable_messages(settings)

    assert fetched_count == 1
    assert processable_messages == [message]


def test_fetch_saves_sent_messages_without_making_them_processable(tmp_path, monkeypatch) -> None:
    settings = _settings(tmp_path)
    received = _message("received")
    sent = EmailMessage(
        gmail_message_id="sent",
        gmail_thread_id="thread-received",
        sender="Me <me@example.com>",
        recipients=("reader@example.com",),
        cc=(),
        subject="Sent mail",
        received_at=datetime(2026, 5, 15, 13, 0, tzinfo=UTC),
        snippet="sent snippet",
        body="sent body",
        attachments=(),
        label_ids=("SENT",),
    )
    monkeypatch.setattr(
        workflow.GmailClient,
        "from_oauth",
        lambda _settings, scopes=None: _FakeGmailClient([received], [sent]),
    )

    processable_messages, fetched_count = workflow.fetch_processable_messages(settings)

    database = Database(settings.db_path)
    assert fetched_count == 2
    assert processable_messages == [received]
    saved_sent = database.get_email("sent")
    assert saved_sent is not None
    assert saved_sent.gmail_message_id == sent.gmail_message_id
    assert saved_sent.label_ids == ("SENT",)
    unanalyzed = database.list_unanalyzed_emails(limit=10)
    assert [message.gmail_message_id for message in unanalyzed] == ["received"]


def test_fetch_preclassifies_gmail_starred_mail_as_high(tmp_path, monkeypatch) -> None:
    settings = _settings(tmp_path)
    starred = EmailMessage(
        gmail_message_id="starred",
        gmail_thread_id="thread-starred",
        sender="sender@example.com",
        recipients=("reader@example.com",),
        cc=(),
        subject="Starred mail",
        received_at=datetime(2026, 5, 15, 12, 0, tzinfo=UTC),
        snippet="snippet",
        body="body",
        attachments=(),
        label_ids=("INBOX", "STARRED"),
    )
    monkeypatch.setattr(
        workflow.GmailClient,
        "from_oauth",
        lambda _settings, scopes=None: _FakeGmailClient([starred]),
    )

    processable_messages, fetched_count = workflow.fetch_processable_messages(settings)

    database = Database(settings.db_path)
    assert fetched_count == 1
    assert processable_messages == []
    with database.session_factory() as session:
        analysis = session.query(EmailAnalysisRecord).one()
    assert analysis.priority == "high"


def test_star_high_priority_messages_adds_gmail_star(tmp_path, monkeypatch) -> None:
    settings = _settings(tmp_path)
    database = Database(settings.db_path)
    database.create_all()
    high = _message("high")
    low = _message("low")
    already_starred = EmailMessage(
        gmail_message_id="already-starred",
        gmail_thread_id="thread-starred",
        sender="sender@example.com",
        recipients=("reader@example.com",),
        cc=(),
        subject="Already starred",
        received_at=datetime(2026, 5, 15, 13, 0, tzinfo=UTC),
        snippet="snippet",
        body="body",
        attachments=(),
        label_ids=("INBOX", "STARRED"),
    )
    database.save_email(high)
    database.save_email(low)
    database.save_email(already_starred)
    database.save_analysis("high", json.loads(_analysis_json("high")), llm_model="test")
    database.save_analysis("low", json.loads(_analysis_json("low")), llm_model="test")
    database.save_analysis("already-starred", json.loads(_analysis_json("high")), llm_model="test")
    fake_client = _FakeGmailClient([])
    monkeypatch.setattr(
        workflow.GmailClient,
        "from_oauth",
        lambda _settings, scopes=None: fake_client,
    )

    starred_count = workflow.star_high_priority_messages(settings)

    assert starred_count == 1
    assert fake_client.starred_message_ids == ["high"]


def test_generate_saved_digests_includes_sent_messages_without_analysis(tmp_path, monkeypatch) -> None:
    settings = _settings(tmp_path)
    database = Database(settings.db_path)
    database.create_all()
    sent = EmailMessage(
        gmail_message_id="sent",
        gmail_thread_id="thread-sent",
        sender="Me <me@example.com>",
        recipients=("reader@example.com",),
        cc=(),
        subject="Sent mail",
        received_at=datetime(2026, 5, 15, 13, 0, tzinfo=UTC),
        snippet="sent snippet",
        body="sent body",
        attachments=(),
        label_ids=("SENT",),
    )
    database.save_email(sent)
    fake_client = _FakeLLMClient()
    monkeypatch.setattr(workflow.LLMClient, "from_settings", lambda _settings: fake_client)

    generated_count = workflow.generate_saved_digests(settings, limit=10)

    assert generated_count == 1
    assert len(fake_client.prompts) == 1
    digest = database.get_digest("sent")
    assert digest is not None
    assert digest.summary_ja
