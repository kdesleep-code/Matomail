"""Email analysis orchestration."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from .database import Database, RulesDatabase
from .llm_client import TextGenerationClient
from .models import EmailMessage


REQUIRED_ANALYSIS_FIELDS = {
    "summary_ja",
    "category",
    "priority",
    "requires_reply",
    "suggested_action_ja",
    "deadline_candidates",
    "meeting_candidates",
    "attachment_action_required",
    "reply_recommended",
    "reply_draft_ja",
    "confidence",
}
VALID_PRIORITIES = {"top", "high", "medium", "low"}


class EmailAnalyzer:
    """Analyzes email content with an LLM."""

    def __init__(
        self,
        llm_client: TextGenerationClient | None = None,
        prompt_template: str | None = None,
    ) -> None:
        self.llm_client = llm_client
        self.prompt_template = prompt_template or _load_default_prompt()

    def build_prompt(
        self,
        message: EmailMessage,
        additional_instructions: list[str] | None = None,
    ) -> str:
        instruction_block = _format_instruction_block(additional_instructions or [])
        return self.prompt_template.format(
            additional_instructions=instruction_block,
            sender=message.sender,
            recipients=", ".join(message.recipients),
            cc=", ".join(message.cc),
            subject=message.subject,
            snippet=message.snippet,
            body=message.body,
        )

    def analyze(
        self,
        message: EmailMessage,
        additional_instructions: list[str] | None = None,
    ) -> dict[str, Any]:
        if self.llm_client is None:
            raise ValueError("llm_client is required to analyze email")

        prompt = self.build_prompt(message, additional_instructions)
        try:
            return parse_analysis_json(self.llm_client.generate_text(prompt))
        except ValueError:
            retry_prompt = (
                f"{prompt}\n\n"
                "前回の出力は有効な JSON ではありませんでした。"
                "説明文や Markdown を含めず，JSON オブジェクトのみを返してください。"
            )
            return parse_analysis_json(self.llm_client.generate_text(retry_prompt))

    def analyze_and_save(
        self,
        message: EmailMessage,
        mail_database: Database,
        rules_database: RulesDatabase,
        llm_model: str,
    ) -> dict[str, Any]:
        instructions = rules_database.get_llm_instructions_for_email(message)
        analysis = self.analyze(message, additional_instructions=instructions)
        mail_database.save_analysis(message.gmail_message_id, analysis, llm_model)
        return analysis


def parse_analysis_json(raw_text: str) -> dict[str, Any]:
    try:
        parsed = json.loads(_strip_json_fence(raw_text))
    except json.JSONDecodeError as error:
        raise ValueError("LLM response was not valid JSON") from error

    if not isinstance(parsed, dict):
        raise ValueError("LLM response must be a JSON object")

    missing_fields = REQUIRED_ANALYSIS_FIELDS - parsed.keys()
    if missing_fields:
        raise ValueError(f"LLM response missing fields: {sorted(missing_fields)}")

    if parsed["priority"] not in VALID_PRIORITIES:
        raise ValueError("priority must be top, high, medium, or low")

    if not isinstance(parsed["deadline_candidates"], list):
        raise ValueError("deadline_candidates must be a list")
    if not isinstance(parsed["meeting_candidates"], list):
        raise ValueError("meeting_candidates must be a list")

    return parsed


def _strip_json_fence(raw_text: str) -> str:
    stripped = raw_text.strip()
    if stripped.startswith("```json"):
        stripped = stripped.removeprefix("```json").strip()
    elif stripped.startswith("```"):
        stripped = stripped.removeprefix("```").strip()
    if stripped.endswith("```"):
        stripped = stripped.removesuffix("```").strip()
    return stripped


def _load_default_prompt() -> str:
    prompt_path = Path(__file__).parent / "prompts" / "email_analysis.md"
    return prompt_path.read_text(encoding="utf-8")


def _format_instruction_block(instructions: list[str]) -> str:
    if not instructions:
        return "なし"
    return "\n".join(f"{index}. {instruction}" for index, instruction in enumerate(instructions, 1))
