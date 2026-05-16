"""Japanese summary and optional full translation generation."""

from __future__ import annotations

import json
from typing import Any

from .database import Database
from .llm_client import TextGenerationClient
from .models import EmailMessage


DIGEST_PRIORITIES = {"high", "medium"}


class EmailDigestGenerator:
    def __init__(self, llm_client: TextGenerationClient) -> None:
        self.llm_client = llm_client

    def generate(self, message: EmailMessage) -> dict[str, str]:
        prompt = _build_digest_prompt(message)
        try:
            digest = parse_digest_json(self.llm_client.generate_text(prompt))
            return _validate_digest_for_message(digest, message)
        except ValueError as error:
            retry_prompt = (
                f"{prompt}\n\n"
                f"前回の出力は利用できませんでした: {error}。"
                "説明文や Markdown を含めず、JSON オブジェクトのみを返してください。"
                "summary_ja は必ず自然な日本語にしてください。"
                "本文が日本語以外の場合、translation_ja には原文ではなく自然な日本語の全文翻訳を入れてください。"
            )
            digest = parse_digest_json(self.llm_client.generate_text(retry_prompt))
            return _validate_digest_for_message(digest, message)

    def generate_and_save(
        self,
        message: EmailMessage,
        mail_database: Database,
        llm_model: str,
    ) -> dict[str, str]:
        digest = self.generate(message)
        mail_database.save_digest(
            message.gmail_message_id,
            summary_ja=digest["summary_ja"],
            translation_ja=digest["translation_ja"],
            llm_model=llm_model,
        )
        return digest


def parse_digest_json(raw_text: str) -> dict[str, str]:
    try:
        parsed: Any = json.loads(_strip_json_fence(raw_text))
    except json.JSONDecodeError as error:
        raise ValueError("LLM digest response was not valid JSON") from error

    if not isinstance(parsed, dict):
        raise ValueError("LLM digest response must be a JSON object")
    if not isinstance(parsed.get("summary_ja"), str) or not parsed["summary_ja"].strip():
        raise ValueError("LLM digest response missing summary_ja")
    return {
        "summary_ja": parsed["summary_ja"].strip(),
        "translation_ja": str(parsed.get("translation_ja") or "").strip(),
    }


def _validate_digest_for_message(
    digest: dict[str, str],
    message: EmailMessage,
) -> dict[str, str]:
    if not _looks_japanese(digest["summary_ja"]):
        raise ValueError("summary_ja must be written in Japanese")

    body = message.body or message.snippet
    if _needs_translation(body):
        translation = digest["translation_ja"]
        if not translation:
            raise ValueError("translation_ja is required for non-Japanese email")
        if not _looks_japanese(translation):
            raise ValueError("translation_ja must be a Japanese translation, not the original text")
        if _normalized_text(translation) == _normalized_text(body):
            raise ValueError("translation_ja must not be identical to the original text")
        return digest

    return {
        **digest,
        "translation_ja": digest["translation_ja"] if _looks_japanese(digest["translation_ja"]) else "",
    }


def _build_digest_prompt(message: EmailMessage) -> str:
    body = message.body or message.snippet
    return (
        "次のメールを処理してください。\n"
        "1. summary_ja: メール内容を日本語で要約してください。"
        "英語など日本語以外のメールでも、自然な日本語の要約にしてください。\n"
        "2. translation_ja: メール本文が日本語以外の場合だけ、本文全体を自然な日本語に翻訳してください。"
        "メール本文が日本語の場合は空文字にしてください。\n"
        "出力は説明文やMarkdownを含めず、JSONオブジェクトだけにしてください。\n"
        "JSON schema: {\"summary_ja\": string, \"translation_ja\": string}\n\n"
        f"From: {message.sender}\n"
        f"To: {', '.join(message.recipients)}\n"
        f"Subject: {message.subject}\n\n"
        f"Body:\n{body[:12000]}"
    )


def _strip_json_fence(raw_text: str) -> str:
    stripped = raw_text.strip()
    if stripped.startswith("```json"):
        stripped = stripped.removeprefix("```json").strip()
    elif stripped.startswith("```"):
        stripped = stripped.removeprefix("```").strip()
    if stripped.endswith("```"):
        stripped = stripped.removesuffix("```").strip()
    return stripped


def _needs_translation(text: str) -> bool:
    sampled = text[:4000]
    japanese_count = _japanese_char_count(sampled)
    ascii_letter_count = sum(1 for character in sampled if character.isascii() and character.isalpha())
    return ascii_letter_count >= 15 and japanese_count < 10


def _looks_japanese(text: str) -> bool:
    if not text:
        return False
    sampled = text[:4000]
    japanese_count = _japanese_char_count(sampled)
    ascii_letter_count = sum(1 for character in sampled if character.isascii() and character.isalpha())
    return japanese_count >= 6 and japanese_count >= max(3, ascii_letter_count // 8)


def _japanese_char_count(text: str) -> int:
    return sum(
        1
        for character in text
        if "\u3040" <= character <= "\u30ff" or "\u4e00" <= character <= "\u9fff"
    )


def _normalized_text(text: str) -> str:
    return " ".join(text.casefold().split())
