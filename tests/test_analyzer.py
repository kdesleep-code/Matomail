from datetime import UTC, datetime
import json

from matomail.analyzer import EmailAnalyzer, parse_analysis_json
from matomail.database import Database, RulesDatabase
from matomail.models import EmailMessage


def test_build_prompt_includes_additional_instructions() -> None:
    analyzer = EmailAnalyzer(
        prompt_template=(
            "Instructions:\n{additional_instructions}\n"
            "Subject: {subject}\n"
            "Body: {body}"
        )
    )
    message = EmailMessage(
        gmail_message_id="msg-1",
        gmail_thread_id="thread-1",
        sender="sender@example.com",
        recipients=("reader@example.com",),
        cc=(),
        subject="査読依頼",
        received_at=datetime(2026, 5, 15, tzinfo=UTC),
        snippet="snippet",
        body="Please review this paper abstract.",
        attachments=(),
    )

    prompt = analyzer.build_prompt(
        message,
        additional_instructions=[
            "査読対応の依頼は Abstract を読み，専門との合致率を表示する。",
            "堀江に関係しない場合は優先度を下げる。",
        ],
    )

    assert "1. 査読対応の依頼は Abstract を読み" in prompt
    assert "2. 堀江に関係しない場合は優先度を下げる。" in prompt
    assert "Subject: 査読依頼" in prompt
    assert "Please review this paper abstract." in prompt


def _analysis_json(priority: str = "high") -> str:
    return json.dumps(
        {
            "summary_ja": "査読依頼です。",
            "category": "査読",
            "priority": priority,
            "requires_reply": True,
            "suggested_action_ja": "Abstract を確認する。",
            "deadline_candidates": [],
            "meeting_candidates": [],
            "attachment_action_required": False,
            "reply_recommended": True,
            "reply_draft_ja": "承知しました。",
            "confidence": 0.9,
        },
        ensure_ascii=False,
    )


class _FakeLLMClient:
    def __init__(self, responses: list[str]) -> None:
        self.responses = responses
        self.prompts: list[str] = []

    def generate_text(self, prompt: str) -> str:
        self.prompts.append(prompt)
        return self.responses.pop(0)


def test_parse_analysis_json_validates_required_fields() -> None:
    parsed = parse_analysis_json(f"```json\n{_analysis_json()}\n```")

    assert parsed["summary_ja"] == "査読依頼です。"
    assert parsed["priority"] == "high"


def test_parse_analysis_json_accepts_top_priority() -> None:
    parsed = parse_analysis_json(_analysis_json("top"))

    assert parsed["priority"] == "top"


def test_analyze_retries_once_on_invalid_json() -> None:
    llm_client = _FakeLLMClient(["not json", _analysis_json("medium")])
    analyzer = EmailAnalyzer(llm_client=llm_client, prompt_template="{body}")

    result = analyzer.analyze(_message())

    assert result["priority"] == "medium"
    assert len(llm_client.prompts) == 2
    assert "前回の出力は有効な JSON" in llm_client.prompts[1]


def test_analyze_and_save_uses_rule_instructions(tmp_path) -> None:
    mail_database = Database(tmp_path / "matomail.sqlite3")
    rules_database = RulesDatabase(tmp_path / "matomail_rules.sqlite3")
    mail_database.create_all()
    rules_database.create_all()
    message = _message()
    mail_database.save_email(message)
    rules_database.add_llm_instruction_rule(
        instruction="査読依頼は専門との合致率を表示する。",
        subject_query="査読",
    )
    llm_client = _FakeLLMClient([_analysis_json()])
    analyzer = EmailAnalyzer(
        llm_client=llm_client,
        prompt_template="{additional_instructions}\n{subject}\n{body}",
    )

    analysis = analyzer.analyze_and_save(
        message,
        mail_database=mail_database,
        rules_database=rules_database,
        llm_model="test-model",
    )

    assert analysis["category"] == "査読"
    assert "査読依頼は専門との合致率" in llm_client.prompts[0]


def _message() -> EmailMessage:
    return EmailMessage(
        gmail_message_id="msg-1",
        gmail_thread_id="thread-1",
        sender="sender@example.com",
        recipients=("reader@example.com",),
        cc=(),
        subject="査読依頼",
        received_at=datetime(2026, 5, 15, tzinfo=UTC),
        snippet="snippet",
        body="Please review this paper abstract.",
        attachments=(),
    )
