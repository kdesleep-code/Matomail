"""LLM client integration."""

from __future__ import annotations

from typing import Protocol

from openai import OpenAI

from .config import Settings


class TextGenerationClient(Protocol):
    def generate_text(self, prompt: str) -> str:
        ...


class LLMClient:
    """Calls the configured LLM provider."""

    def __init__(self, api_key: str, model: str) -> None:
        if not api_key:
            raise ValueError("OPENAI_API_KEY is required for LLM analysis")
        self.model = model
        self._client = OpenAI(api_key=api_key)

    @classmethod
    def from_settings(cls, settings: Settings | None = None) -> "LLMClient":
        settings = settings or Settings()
        return cls(api_key=settings.openai_api_key, model=settings.llm_model)

    def generate_text(self, prompt: str) -> str:
        response = self._client.responses.create(
            model=self.model,
            input=prompt,
        )
        return response.output_text
