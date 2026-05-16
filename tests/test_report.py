from datetime import UTC, datetime

from sqlalchemy import select

from matomail.database import Database, EmailRecord
from matomail.models import AttachmentInfo, EmailMessage
from matomail.report import ReportGenerator


def _message(
    message_id: str,
    subject: str,
    received_at: datetime,
    body: str = "full message body",
    body_html: str = "",
    sender: str = "sender@example.com",
    recipients: tuple[str, ...] = ("reader@example.com",),
    label_ids: tuple[str, ...] = (),
    thread_id: str | None = None,
) -> EmailMessage:
    return EmailMessage(
        gmail_message_id=message_id,
        gmail_thread_id=thread_id or f"thread-{message_id}",
        sender=sender,
        recipients=recipients,
        cc=(),
        subject=subject,
        received_at=received_at,
        snippet=f"snippet for {subject}",
        body=body,
        attachments=(),
        body_html=body_html,
        label_ids=label_ids,
    )


def _analysis(priority: str) -> dict:
    return {
        "summary_ja": "summary",
        "category": "category",
        "priority": priority,
        "requires_reply": False,
        "suggested_action_ja": "none",
        "deadline_candidates": [],
        "meeting_candidates": [],
        "reply_draft_ja": "",
        "confidence": 0.5,
    }


def test_report_generator_groups_by_loaded_date_not_received_date(tmp_path) -> None:
    database = Database(tmp_path / "matomail.sqlite3")
    database.create_all()
    first = _message(
        "msg-1",
        "First mail",
        datetime(2026, 5, 14, 12, 0, tzinfo=UTC),
        body="Body of the first mail.",
        body_html="<main><h1>HTML first mail</h1><p>Rich body.</p></main>",
    )
    second = _message(
        "msg-2",
        "Second mail",
        datetime(2026, 5, 15, 12, 0, tzinfo=UTC),
        body="Body of the second mail.",
    )
    database.save_email(first)
    database.save_email(second)
    with database.session_factory() as session:
        records = session.scalars(select(EmailRecord)).all()
        for record in records:
            record.created_at = datetime(2026, 5, 16, 1, 0, tzinfo=UTC)
        session.commit()
    database.save_analysis("msg-1", _analysis("low"), llm_model="test-model")
    database.save_analysis("msg-2", _analysis("high"), llm_model="test-model")

    report_path = ReportGenerator(
        database=database,
        report_dir=tmp_path / "reports",
        timezone="Asia/Tokyo",
    ).generate_all(open_browser=False)

    assert report_path == tmp_path / "reports" / "2026-05-16" / "index.html"
    index = tmp_path / "reports" / "2026-05-16" / "index.html"
    first_message = tmp_path / "reports" / "2026-05-16" / "messages" / "msg-1.html"
    assert index.exists()
    assert first_message.exists()

    html = index.read_text(encoding="utf-8")
    assert "2026-05-16" in html
    assert "First mail" in html
    assert "Second mail" in html
    assert "sender@example.com" in html
    assert "reader@example.com" in html
    assert 'id="mailSearch"' in html
    assert 'id="mailSort"' in html
    assert "受信: 2026-05-14" in html

    message_html = first_message.read_text(encoding="utf-8")
    assert 'sandbox=""' in message_html
    assert "HTML first mail" in message_html
    assert "メール一覧に戻る" in message_html


def test_report_generator_ranks_top_above_high(tmp_path) -> None:
    database = Database(tmp_path / "matomail.sqlite3")
    database.create_all()
    top = _message("top", "Top mail", datetime(2026, 5, 15, 12, 0, tzinfo=UTC))
    high = _message("high", "High mail", datetime(2026, 5, 15, 12, 5, tzinfo=UTC))
    database.save_email(top)
    database.save_email(high)
    database.save_analysis("top", _analysis("top"), llm_model="test-model")
    database.save_analysis("high", _analysis("high"), llm_model="test-model")

    ReportGenerator(
        database=database,
        report_dir=tmp_path / "reports",
        timezone="Asia/Tokyo",
    ).generate_all(open_browser=False)

    html = (tmp_path / "reports" / "2026-05-15" / "index.html").read_text(
        encoding="utf-8"
    )
    assert 'data-priority="4"' in html
    assert "priority-top" in html


def test_report_generator_shows_summary_digest_and_collapses_original_body(tmp_path) -> None:
    database = Database(tmp_path / "matomail.sqlite3")
    database.create_all()
    message = _message(
        "msg-digest",
        "Digest mail",
        datetime(2026, 5, 15, 12, 0, tzinfo=UTC),
        body="Original body.",
    )
    database.save_email(message)
    database.save_analysis("msg-digest", _analysis("high"), llm_model="test-model")
    database.save_digest(
        "msg-digest",
        summary_ja="重要な要約です。",
        translation_ja="全文和訳です。",
        llm_model="test-model",
    )

    ReportGenerator(
        database=database,
        report_dir=tmp_path / "reports",
        timezone="Asia/Tokyo",
    ).generate_all(open_browser=False)

    message_html = (
        tmp_path / "reports" / "2026-05-15" / "messages" / "msg-digest.html"
    ).read_text(encoding="utf-8")
    assert "重要な要約です。" in message_html
    assert "<h2>翻訳</h2>" not in message_html
    assert "全文和訳です。" in message_html
    assert "和訳全文を開く" in message_html
    assert "原語全文を開く" in message_html
    assert "<details" in message_html


def test_report_generator_can_exclude_configured_sender_addresses(tmp_path) -> None:
    database = Database(tmp_path / "matomail.sqlite3")
    database.create_all()
    received = _message(
        "received",
        "Received mail",
        datetime(2026, 5, 15, 12, 0, tzinfo=UTC),
        sender="other@example.com",
    )
    sent = _message(
        "sent",
        "Sent mail",
        datetime(2026, 5, 15, 12, 5, tzinfo=UTC),
        sender="Me <me@example.com>",
    )
    database.save_email(received)
    database.save_email(sent)

    ReportGenerator(
        database=database,
        report_dir=tmp_path / "reports",
        timezone="Asia/Tokyo",
        excluded_sender_addresses=("me@example.com",),
    ).generate_all(open_browser=False)

    html = (tmp_path / "reports" / "2026-05-15" / "index.html").read_text(
        encoding="utf-8"
    )
    assert "Received mail" in html
    assert "Sent mail" not in html


def test_report_generator_uses_saved_account_email_without_deleting_sent_mail(tmp_path) -> None:
    database = Database(tmp_path / "matomail.sqlite3")
    database.create_all()
    database.save_account_email("me@example.com")
    received = _message(
        "received",
        "Received mail",
        datetime(2026, 5, 15, 12, 0, tzinfo=UTC),
        sender="other@example.com",
        recipients=("me@example.com",),
    )
    sent = _message(
        "sent",
        "Sent mail",
        datetime(2026, 5, 15, 12, 5, tzinfo=UTC),
        sender="Me <me@example.com>",
        recipients=("other@example.com",),
    )
    database.save_email(received)
    database.save_email(sent)

    ReportGenerator(
        database=database,
        report_dir=tmp_path / "reports",
        timezone="Asia/Tokyo",
    ).generate_all(open_browser=False)

    html = (tmp_path / "reports" / "2026-05-15" / "index.html").read_text(
        encoding="utf-8"
    )
    assert database.count_emails() == 2
    assert "Received mail" in html
    assert "Sent mail" not in html


def test_report_generator_excludes_gmail_sent_label(tmp_path) -> None:
    database = Database(tmp_path / "matomail.sqlite3")
    database.create_all()
    database.save_email(
        _message(
            "sent",
            "Sent mail",
            datetime(2026, 5, 15, 12, 5, tzinfo=UTC),
            sender="someone@example.com",
            label_ids=("SENT",),
        )
    )

    report_path = ReportGenerator(
        database=database,
        report_dir=tmp_path / "reports",
        timezone="Asia/Tokyo",
    ).generate_all(open_browser=False)

    assert report_path is None
    assert database.count_emails() == 1


def test_report_generator_merges_duplicate_messages_and_combines_recipients(tmp_path) -> None:
    database = Database(tmp_path / "matomail.sqlite3")
    database.create_all()
    first = _message(
        "duplicate-1",
        "Duplicate mail",
        datetime(2026, 5, 15, 12, 0, tzinfo=UTC),
        body="Same body.",
        recipients=("first@example.com",),
    )
    second = _message(
        "duplicate-2",
        "Duplicate mail",
        datetime(2026, 5, 15, 12, 1, tzinfo=UTC),
        body="Same body.",
        recipients=("second@example.com",),
    )
    database.save_email(first)
    database.save_email(second)
    database.save_analysis("duplicate-1", _analysis("high"), llm_model="test-model")
    database.save_analysis("duplicate-2", _analysis("high"), llm_model="test-model")

    ReportGenerator(
        database=database,
        report_dir=tmp_path / "reports",
        timezone="Asia/Tokyo",
    ).generate_all(open_browser=False)

    html = (tmp_path / "reports" / "2026-05-15" / "index.html").read_text(
        encoding="utf-8"
    )
    assert html.count('class="mail-card"') == 1
    assert "first@example.com" in html
    assert "second@example.com" in html


def test_report_generator_groups_thread_messages_into_one_card_and_detail_page(tmp_path) -> None:
    database = Database(tmp_path / "matomail.sqlite3")
    database.create_all()
    older = _message(
        "thread-old",
        "Thread mail",
        datetime(2026, 5, 15, 11, 0, tzinfo=UTC),
        body="Older body.",
        thread_id="gmail-thread-1",
    )
    newer = _message(
        "thread-new",
        "Re: Thread mail",
        datetime(2026, 5, 15, 12, 0, tzinfo=UTC),
        body="Newer body.",
        thread_id="gmail-thread-1",
    )
    database.save_email(older)
    database.save_email(newer)
    database.save_analysis("thread-old", _analysis("low"), llm_model="test-model")
    database.save_analysis("thread-new", _analysis("high"), llm_model="test-model")

    ReportGenerator(
        database=database,
        report_dir=tmp_path / "reports",
        timezone="Asia/Tokyo",
    ).generate_all(open_browser=False)

    index_html = (tmp_path / "reports" / "2026-05-15" / "index.html").read_text(
        encoding="utf-8"
    )
    assert index_html.count('class="mail-card"') == 1
    assert "2 件" in index_html
    assert "Re: Thread mail" in index_html

    message_html = (
        tmp_path / "reports" / "2026-05-15" / "messages" / "gmail-thread-1.html"
    ).read_text(encoding="utf-8")
    assert message_html.index("Newer body.") < message_html.index("Older body.")


def test_message_detail_edge_nav_uses_priority_list_neighbors(tmp_path) -> None:
    database = Database(tmp_path / "matomail.sqlite3")
    database.create_all()
    top = _message("top", "Top mail", datetime(2026, 5, 15, 12, 0, tzinfo=UTC))
    high = _message("high", "High mail", datetime(2026, 5, 15, 12, 5, tzinfo=UTC))
    low = _message("low", "Low mail", datetime(2026, 5, 15, 12, 10, tzinfo=UTC))
    database.save_email(top)
    database.save_email(high)
    database.save_email(low)
    with database.session_factory() as session:
        for record in session.scalars(select(EmailRecord)).all():
            record.created_at = datetime(2026, 5, 15, 9, 0, tzinfo=UTC)
        session.commit()
    database.save_analysis("top", _analysis("top"), llm_model="test-model")
    database.save_analysis("high", _analysis("high"), llm_model="test-model")
    database.save_analysis("low", _analysis("low"), llm_model="test-model")

    ReportGenerator(
        database=database,
        report_dir=tmp_path / "reports",
        timezone="Asia/Tokyo",
    ).generate_all(open_browser=False)

    top_html = (tmp_path / "reports" / "2026-05-15" / "messages" / "top.html").read_text(
        encoding="utf-8"
    )
    high_html = (tmp_path / "reports" / "2026-05-15" / "messages" / "high.html").read_text(
        encoding="utf-8"
    )
    low_html = (tmp_path / "reports" / "2026-05-15" / "messages" / "low.html").read_text(
        encoding="utf-8"
    )
    assert 'class="edge-nav left disabled"' in top_html
    assert 'class="edge-nav right" href="high.html"' in top_html
    assert 'class="edge-nav left" href="top.html"' in high_html
    assert 'class="edge-nav right" href="low.html"' in high_html
    assert 'class="edge-nav left" href="high.html"' in low_html
    assert 'class="edge-nav right disabled"' in low_html


def test_thread_detail_includes_past_loaded_and_sent_messages_with_divider(tmp_path) -> None:
    database = Database(tmp_path / "matomail.sqlite3")
    database.create_all()
    database.save_account_email("me@example.com")
    current = _message(
        "current",
        "Current thread mail",
        datetime(2026, 5, 15, 12, 0, tzinfo=UTC),
        body="Current body.",
        sender="other@example.com",
        thread_id="thread-with-history",
    )
    past = _message(
        "past",
        "Past thread mail",
        datetime(2026, 5, 14, 12, 0, tzinfo=UTC),
        body="Past body.",
        sender="other@example.com",
        thread_id="thread-with-history",
    )
    sent = _message(
        "sent",
        "Sent thread mail",
        datetime(2026, 5, 14, 13, 0, tzinfo=UTC),
        body="Sent body.",
        sender="Me <me@example.com>",
        recipients=("other@example.com",),
        thread_id="thread-with-history",
    )
    newer = _message(
        "newer",
        "Newer thread mail",
        datetime(2026, 5, 16, 12, 0, tzinfo=UTC),
        body="Newer body.",
        sender="other@example.com",
        thread_id="thread-with-history",
    )
    database.save_email(current)
    database.save_email(past)
    database.save_email(sent)
    database.save_email(newer)
    with database.session_factory() as session:
        records = {
            record.gmail_message_id: record
            for record in session.scalars(select(EmailRecord)).all()
        }
        records["current"].created_at = datetime(2026, 5, 15, 9, 0, tzinfo=UTC)
        records["past"].created_at = datetime(2026, 5, 14, 9, 0, tzinfo=UTC)
        records["sent"].created_at = datetime(2026, 5, 14, 10, 0, tzinfo=UTC)
        records["newer"].created_at = datetime(2026, 5, 16, 9, 0, tzinfo=UTC)
        session.commit()
    database.save_analysis("current", _analysis("high"), llm_model="test-model")
    database.save_analysis("past", _analysis("medium"), llm_model="test-model")
    database.save_analysis("sent", _analysis("medium"), llm_model="test-model")
    database.save_analysis("newer", _analysis("high"), llm_model="test-model")

    ReportGenerator(
        database=database,
        report_dir=tmp_path / "reports",
        timezone="Asia/Tokyo",
    ).generate_all(open_browser=False)

    index_html = (tmp_path / "reports" / "2026-05-15" / "index.html").read_text(
        encoding="utf-8"
    )
    assert index_html.count('class="mail-card"') == 1
    assert "Sent thread mail" not in index_html

    message_html = (
        tmp_path / "reports" / "2026-05-15" / "messages" / "thread-with-history.html"
    ).read_text(encoding="utf-8")
    assert "Newer body." in message_html
    assert "Current body." in message_html
    assert "Past body." in message_html
    assert "Sent body." in message_html
    assert "ーー本日のメールはここからーー" in message_html
    assert "ーー本日のメールはここまでーー" in message_html
    assert message_html.index("Current body.") < message_html.index("ーー本日のメールはここまでーー")
    assert 'style="--sender-bg: hsl(' in message_html
    assert 'style="--sender-bg: #ffffff;"' in message_html
    assert message_html.count("<details") >= 1


def test_thread_detail_includes_sent_message_loaded_on_same_day(tmp_path) -> None:
    database = Database(tmp_path / "matomail.sqlite3")
    database.create_all()
    database.save_account_email("me@example.com")
    received = _message(
        "received",
        "Request for Leave and Signature for University Loan Application",
        datetime(2026, 5, 15, 3, 0, tzinfo=UTC),
        body="Please sign the document.",
        sender="student@example.com",
        thread_id="loan-thread",
    )
    sent = _message(
        "sent",
        "Re: Request for Leave and Signature for University Loan Application",
        datetime(2026, 5, 15, 8, 0, tzinfo=UTC),
        body="I signed it.",
        sender="Me <me@example.com>",
        recipients=("student@example.com",),
        label_ids=("SENT",),
        thread_id="loan-thread",
    )
    database.save_email(received)
    database.save_email(sent)
    database.save_analysis("received", _analysis("high"), llm_model="test-model")
    database.save_analysis("sent", _analysis("medium"), llm_model="test-model")
    database.save_digest(
        "sent",
        summary_ja="署名済みで返信したメールです。",
        translation_ja="",
        llm_model="test-model",
    )

    ReportGenerator(
        database=database,
        report_dir=tmp_path / "reports",
        timezone="Asia/Tokyo",
    ).generate_all(open_browser=False)

    message_html = (
        tmp_path / "reports" / "2026-05-15" / "messages" / "loan-thread.html"
    ).read_text(encoding="utf-8")
    assert "Re: Request for Leave and Signature" in message_html
    assert "I signed it." in message_html
    assert 'style="--sender-bg: #ffffff;"' in message_html
    assert message_html.index("I signed it.") < message_html.index("Please sign the document.")


def test_report_generator_links_attachments_to_manual_download_route(tmp_path) -> None:
    database = Database(tmp_path / "matomail.sqlite3")
    database.create_all()
    message = EmailMessage(
        gmail_message_id="msg-attachment",
        gmail_thread_id="thread-attachment",
        sender="sender@example.com",
        recipients=("reader@example.com",),
        cc=(),
        subject="Attachment mail",
        received_at=datetime(2026, 5, 15, 12, 0, tzinfo=UTC),
        snippet="snippet",
        body="body",
        attachments=(
            AttachmentInfo(
                filename="agenda.pdf",
                mime_type="application/pdf",
                size=123,
                attachment_id="att-1",
            ),
        ),
    )
    database.save_email(message)
    database.save_analysis("msg-attachment", _analysis("medium"), llm_model="test-model")

    ReportGenerator(
        database=database,
        report_dir=tmp_path / "reports",
        timezone="Asia/Tokyo",
    ).generate_all(open_browser=False)

    message_html = (
        tmp_path / "reports" / "2026-05-15" / "messages" / "msg-attachment.html"
    ).read_text(encoding="utf-8")
    assert "/attachments/msg-attachment/0" in message_html
    assert "agenda.pdf" in message_html

    index_html = (tmp_path / "reports" / "2026-05-15" / "index.html").read_text(
        encoding="utf-8"
    )
    assert "添付ファイル: pdf x 1" in index_html
