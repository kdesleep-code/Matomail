from datetime import UTC, datetime

from matomail.digest import EmailDigestGenerator, parse_digest_json
from matomail.models import EmailMessage


class _FakeLLMClient:
    def __init__(self, responses: list[str]) -> None:
        self.responses = responses
        self.prompts: list[str] = []

    def generate_text(self, prompt: str) -> str:
        self.prompts.append(prompt)
        return self.responses.pop(0)


def test_parse_digest_json_accepts_summary_and_optional_translation() -> None:
    parsed = parse_digest_json(
        '{"summary_ja": "日本語の要約です。", "translation_ja": "全文翻訳です。"}'
    )

    assert parsed == {
        "summary_ja": "日本語の要約です。",
        "translation_ja": "全文翻訳です。",
    }


def test_parse_digest_json_defaults_translation_to_empty() -> None:
    parsed = parse_digest_json('{"summary_ja": "日本語の要約です。"}')

    assert parsed == {"summary_ja": "日本語の要約です。", "translation_ja": ""}


def test_digest_generator_asks_for_summary_and_translation_when_needed() -> None:
    client = _FakeLLMClient(
        ['{"summary_ja": "会議への参加依頼です。", "translation_ja": "会議に参加してください。"}']
    )
    generator = EmailDigestGenerator(client)
    message = _message()

    digest = generator.generate(message)

    assert digest["summary_ja"] == "会議への参加依頼です。"
    assert digest["translation_ja"] == "会議に参加してください。"
    assert "メール本文が日本語以外の場合だけ" in client.prompts[0]
    assert "Please attend the meeting." in client.prompts[0]


def test_digest_generator_retries_when_response_is_not_json() -> None:
    client = _FakeLLMClient(
        [
            "要約: 会議です。",
            '{"summary_ja": "会議への参加依頼です。", "translation_ja": "会議に参加してください。"}',
        ]
    )
    generator = EmailDigestGenerator(client)

    digest = generator.generate(_message())

    assert digest["summary_ja"] == "会議への参加依頼です。"
    assert len(client.prompts) == 2
    assert "前回の出力は利用できませんでした" in client.prompts[1]


def test_digest_generator_retries_when_translation_is_not_japanese() -> None:
    client = _FakeLLMClient(
        [
            '{"summary_ja": "会議への参加依頼です。", "translation_ja": "Please attend the meeting."}',
            '{"summary_ja": "会議への参加依頼です。", "translation_ja": "会議に参加してください。"}',
        ]
    )
    generator = EmailDigestGenerator(client)

    digest = generator.generate(_message())

    assert digest["translation_ja"] == "会議に参加してください。"
    assert len(client.prompts) == 2
    assert "自然な日本語の全文翻訳" in client.prompts[1]


def test_digest_generator_rejects_non_japanese_summary_after_retry() -> None:
    client = _FakeLLMClient(
        [
            '{"summary_ja": "Please attend.", "translation_ja": "会議に参加してください。"}',
            '{"summary_ja": "Please attend.", "translation_ja": "会議に参加してください。"}',
        ]
    )
    generator = EmailDigestGenerator(client)

    try:
        generator.generate(_message())
    except ValueError as error:
        assert "summary_ja" in str(error)
    else:
        raise AssertionError("digest generation should fail")


def _message() -> EmailMessage:
    return EmailMessage(
        gmail_message_id="msg-1",
        gmail_thread_id="thread-1",
        sender="sender@example.com",
        recipients=("reader@example.com",),
        cc=(),
        subject="Meeting",
        received_at=datetime(2026, 5, 15, tzinfo=UTC),
        snippet="snippet",
        body="Please attend the meeting.",
        attachments=(),
    )
