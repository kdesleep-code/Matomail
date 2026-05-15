from datetime import UTC, datetime

from sqlalchemy import inspect
from sqlalchemy import select

from matomail.database import Database, EmailRecord, RulesDatabase
from matomail.database import FILTER_ACTION_ALWAYS_PROCESS
from matomail.database import FILTER_ACTION_PRECLASSIFY
from matomail.database import FILTER_ACTION_SKIP_ANALYSIS
from matomail.models import AttachmentInfo, EmailMessage


def _message(message_id: str = "msg-1") -> EmailMessage:
    return EmailMessage(
        gmail_message_id=message_id,
        gmail_thread_id="thread-1",
        sender="sender@example.com",
        recipients=("reader@example.com",),
        cc=(),
        subject="Database test",
        received_at=datetime(2026, 5, 14, 12, 0, tzinfo=UTC),
        snippet="snippet",
        body="full body for local persistence",
        attachments=(
            AttachmentInfo(
                filename="agenda.pdf",
                mime_type="application/pdf",
                size=42,
                attachment_id="att-1",
            ),
        ),
        size_estimate=2048,
    )


def test_create_all_creates_task_3_tables(tmp_path) -> None:
    database = Database(tmp_path / "matomail.sqlite3")

    database.create_all()

    table_names = set(inspect(database.engine).get_table_names())
    assert {
        "emails",
        "email_analysis",
        "filter_decisions",
        "processing_state",
    } <= table_names
    assert "filter_rules" not in table_names
    assert "llm_instruction_rules" not in table_names


def test_create_all_creates_rules_tables_in_separate_database(tmp_path) -> None:
    rules_database = RulesDatabase(tmp_path / "matomail_rules.sqlite3")

    rules_database.create_all()

    table_names = set(inspect(rules_database.engine).get_table_names())
    assert {
        "filter_rules",
        "llm_instruction_rules",
    } <= table_names
    assert "emails" not in table_names


def test_save_email_upserts_by_gmail_message_id(tmp_path) -> None:
    database = Database(tmp_path / "matomail.sqlite3")
    database.create_all()

    database.save_email(_message())
    database.save_email(_message())

    assert database.count_emails() == 1
    assert database.get_processing_status("msg-1") == "pending"

    with database.session_factory() as session:
        record = session.scalar(select(EmailRecord))
        assert record is not None
        assert record.body == "full body for local persistence"


def test_list_unanalyzed_emails_returns_saved_message_objects(tmp_path) -> None:
    database = Database(tmp_path / "matomail.sqlite3")
    database.create_all()
    database.save_email(_message())

    messages = database.list_unanalyzed_emails(limit=1)

    assert len(messages) == 1
    assert messages[0].gmail_message_id == "msg-1"
    assert messages[0].body == "full body for local persistence"
    assert messages[0].attachments[0].filename == "agenda.pdf"


def test_list_unanalyzed_emails_excludes_analyzed_messages(tmp_path) -> None:
    database = Database(tmp_path / "matomail.sqlite3")
    database.create_all()
    database.save_email(_message())
    database.save_analysis(
        "msg-1",
        {
            "summary_ja": "summary",
            "category": "category",
            "priority": "medium",
            "requires_reply": False,
            "suggested_action_ja": "none",
            "deadline_candidates": [],
            "meeting_candidates": [],
            "reply_draft_ja": "",
            "confidence": 0.5,
        },
        llm_model="test-model",
    )

    assert database.list_unanalyzed_emails(limit=1) == []


def test_final_statuses_are_not_processable_but_pending_is(tmp_path) -> None:
    database = Database(tmp_path / "matomail.sqlite3")
    database.create_all()
    database.save_email(_message())

    assert database.should_process("msg-1") is True

    database.mark_status("msg-1", "processed")
    assert database.should_process("msg-1") is False
    assert database.filter_processable([_message()]) == []

    database.mark_status("msg-1", "pending")
    assert database.should_process("msg-1") is True

    database.mark_status("msg-1", "skipped")
    assert database.should_process("msg-1") is False


def test_can_disable_email_body_persistence(tmp_path) -> None:
    database = Database(tmp_path / "matomail.sqlite3", store_email_body=False)
    database.create_all()

    database.save_email(_message())

    with database.session_factory() as session:
        record = session.scalar(select(EmailRecord))
        assert record is not None
        assert record.body == ""


def test_large_database_is_rotated_to_timestamped_backup_directory(tmp_path) -> None:
    db_path = tmp_path / "matomail.sqlite3"
    backup_dir = tmp_path / "backups"
    db_path.write_bytes(b"x" * 20)

    database = Database(db_path, max_size_bytes=10, backup_dir=backup_dir)
    database.create_all()

    backups = list(backup_dir.glob("*/matomail.sqlite3"))
    assert len(backups) == 1
    assert backups[0].read_bytes() == b"x" * 20
    assert db_path.exists()
    assert database.count_emails() == 0


def test_sender_filter_excludes_matching_messages_from_processable_list(tmp_path) -> None:
    database = RulesDatabase(tmp_path / "matomail_rules.sqlite3")
    database.create_all()

    database.add_sender_filter("Sender <sender@example.com>", note="newsletter")

    blocked = _message("blocked")
    allowed = _message("allowed")
    allowed = EmailMessage(
        gmail_message_id=allowed.gmail_message_id,
        gmail_thread_id=allowed.gmail_thread_id,
        sender="other@example.com",
        recipients=allowed.recipients,
        cc=allowed.cc,
        subject=allowed.subject,
        received_at=allowed.received_at,
        snippet=allowed.snippet,
        body=allowed.body,
        attachments=allowed.attachments,
    )

    assert database.should_skip_analysis(blocked) is True
    assert database.should_skip_analysis(allowed) is False
    assert database.filter_processable([blocked, allowed]) == [allowed]


def test_sender_filter_is_idempotent_and_normalized(tmp_path) -> None:
    database = RulesDatabase(tmp_path / "matomail_rules.sqlite3")
    database.create_all()

    first = database.add_sender_filter("Sender <SENDER@EXAMPLE.COM>", note="first")
    second = database.add_sender_filter("sender@example.com", note="second")

    assert first.id != second.id
    assert second.from_query == "sender@example.com"


def test_filter_rules_can_be_listed_disabled_and_deleted(tmp_path) -> None:
    database = RulesDatabase(tmp_path / "matomail_rules.sqlite3")
    database.create_all()

    first = database.add_sender_filter(
        "sender@example.com",
        name="sender skip",
    )
    second = database.add_filter_rule(
        action=FILTER_ACTION_ALWAYS_PROCESS,
        subject_query="Database",
        name="subject force",
        priority=10,
    )

    rules = database.list_filter_rules()
    assert [rule.id for rule in rules] == [second.id, first.id]

    assert database.set_filter_rule_enabled(first.id, False) is True
    assert database.should_skip_analysis(_message()) is False

    assert database.delete_filter_rule(second.id) is True
    assert [rule.id for rule in database.list_filter_rules()] == [first.id]
    assert database.delete_filter_rule(9999) is False


def test_filter_rules_support_gmail_like_criteria(tmp_path) -> None:
    database = RulesDatabase(tmp_path / "matomail_rules.sqlite3")
    database.create_all()

    database.add_filter_rule(
        action=FILTER_ACTION_SKIP_ANALYSIS,
        from_query="sender@example.com",
        to_query="reader@example.com",
        subject_query="Database",
        has_words="local persistence",
        doesnt_have="do not match",
        has_attachment=True,
        filename_query="agenda",
        size_comparison="larger",
        size_bytes=1024,
        before=datetime(2026, 5, 15, tzinfo=UTC),
    )

    assert database.get_filter_action(_message()) == FILTER_ACTION_SKIP_ANALYSIS


def test_filter_decisions_are_saved_in_mail_database(tmp_path) -> None:
    mail_database = Database(tmp_path / "matomail.sqlite3")
    rules_database = RulesDatabase(tmp_path / "matomail_rules.sqlite3")
    mail_database.create_all()
    rules_database.create_all()
    message = _message()
    mail_database.save_email(message)
    rules_database.add_filter_rule(
        action=FILTER_ACTION_SKIP_ANALYSIS,
        name="newsletter",
        from_query="sender@example.com",
    )

    decision = rules_database.get_filter_decision(message)
    assert decision is not None
    record = mail_database.save_filter_decision(
        message.gmail_message_id,
        action=decision["action"],
        matched_rule_id=decision["matched_rule_id"],
        matched_rule_name=decision["matched_rule_name"],
        reason=decision["reason"],
        rule_snapshot=decision["rule_snapshot"],
    )

    assert record.action == FILTER_ACTION_SKIP_ANALYSIS
    assert record.matched_rule_name == "newsletter"


def test_preclassify_rule_creates_analysis_without_llm(tmp_path) -> None:
    mail_database = Database(tmp_path / "matomail.sqlite3")
    rules_database = RulesDatabase(tmp_path / "matomail_rules.sqlite3")
    mail_database.create_all()
    rules_database.create_all()
    message = _message()
    mail_database.save_email(message)
    rules_database.add_filter_rule(
        action=FILTER_ACTION_PRECLASSIFY,
        name="priority rule",
        subject_query="Database",
        preset_priority="high",
        preset_category="important",
        preset_summary_ja="重要ルールに一致しました。",
        preset_suggested_action_ja="優先的に確認する。",
        preset_requires_reply=True,
    )

    decision = rules_database.get_filter_decision(message)
    assert decision is not None
    mail_database.save_filter_decision(
        message.gmail_message_id,
        action=decision["action"],
        matched_rule_id=decision["matched_rule_id"],
        matched_rule_name=decision["matched_rule_name"],
        reason=decision["reason"],
        rule_snapshot=decision["rule_snapshot"],
    )
    analysis = mail_database.apply_preclassified_analysis(
        message.gmail_message_id,
        decision["preset_analysis"],
    )

    assert analysis.priority == "high"
    assert analysis.category == "important"
    assert analysis.requires_reply is True
    assert mail_database.list_unanalyzed_emails(limit=1) == []


def test_always_process_can_override_lower_priority_skip_filter(tmp_path) -> None:
    database = RulesDatabase(tmp_path / "matomail_rules.sqlite3")
    database.create_all()

    database.add_filter_rule(
        action=FILTER_ACTION_SKIP_ANALYSIS,
        from_query="sender@example.com",
        priority=0,
    )
    database.add_filter_rule(
        action=FILTER_ACTION_ALWAYS_PROCESS,
        subject_query="Database",
        priority=10,
    )

    assert database.get_filter_action(_message()) == FILTER_ACTION_ALWAYS_PROCESS
    assert database.filter_processable([_message()]) == [_message()]


def test_llm_instruction_rules_return_matching_instructions_by_priority(tmp_path) -> None:
    database = RulesDatabase(tmp_path / "matomail_rules.sqlite3")
    database.create_all()

    database.add_llm_instruction_rule(
        name="review request",
        instruction="査読対応の依頼は Abstract を読み，専門との合致率を表示する。",
        subject_query="Database",
        priority=5,
    )
    database.add_llm_instruction_rule(
        name="web update",
        instruction="学位プログラムのWeb更新は優先し，学類のWeb更新は低優先度にする。",
        has_words="local persistence",
        priority=10,
    )
    database.add_llm_instruction_rule(
        name="not matching",
        instruction="これは一致しない。",
        has_words="missing phrase",
        priority=20,
    )

    assert database.get_llm_instructions_for_email(_message()) == [
        "学位プログラムのWeb更新は優先し，学類のWeb更新は低優先度にする。",
        "査読対応の依頼は Abstract を読み，専門との合致率を表示する。",
    ]


def test_llm_instruction_rules_can_be_listed_disabled_and_deleted(tmp_path) -> None:
    database = RulesDatabase(tmp_path / "matomail_rules.sqlite3")
    database.create_all()
    rule = database.add_llm_instruction_rule(
        instruction="専門との合致率を表示する。",
        subject_query="Database",
        name="review instruction",
    )

    assert [item.id for item in database.list_llm_instruction_rules()] == [rule.id]
    assert database.set_llm_instruction_rule_enabled(rule.id, False) is True
    assert database.get_llm_instructions_for_email(_message()) == []
    assert database.delete_llm_instruction_rule(rule.id) is True
    assert database.list_llm_instruction_rules() == []
