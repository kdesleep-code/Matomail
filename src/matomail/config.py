"""Runtime configuration."""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path


def _read_dotenv(path: Path = Path(".env")) -> dict[str, str]:
    if not path.exists():
        return {}

    values: dict[str, str] = {}
    for line in path.read_text(encoding="utf-8").splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#") or "=" not in stripped:
            continue
        key, value = stripped.split("=", 1)
        values[key.strip()] = value.strip().strip('"').strip("'")
    return values


_DOTENV = _read_dotenv()


def _env(name: str, default: str) -> str:
    return os.getenv(name) or _DOTENV.get(name, default)


def _env_bool(name: str, default: bool) -> bool:
    raw_value = os.getenv(name) or _DOTENV.get(name)
    if raw_value is None:
        return default
    return raw_value.strip().lower() in {"1", "true", "yes", "on"}


def _env_int(name: str, default: int) -> int:
    raw_value = os.getenv(name) or _DOTENV.get(name)
    if raw_value is None:
        return default
    value = int(raw_value)
    if value < 1:
        raise ValueError(f"{name} must be at least 1")
    return value


def _env_float(name: str, default: float) -> float:
    raw_value = os.getenv(name) or _DOTENV.get(name)
    if raw_value is None:
        return default
    value = float(raw_value)
    if value <= 0:
        raise ValueError(f"{name} must be greater than 0")
    return value


@dataclass(frozen=True)
class Settings:
    lookback_days: int = field(
        default_factory=lambda: _env_int("MATOMAIL_LOOKBACK_DAYS", 7)
    )
    report_dir: Path = field(
        default_factory=lambda: Path(_env("MATOMAIL_REPORT_DIR", "./reports"))
    )
    db_path: Path = field(
        default_factory=lambda: Path(_env("MATOMAIL_DB_PATH", "./data/matomail.sqlite3"))
    )
    rules_db_path: Path = field(
        default_factory=lambda: Path(_env("MATOMAIL_RULES_DB_PATH", "./data/matomail_rules.sqlite3"))
    )
    db_backup_dir: Path = field(
        default_factory=lambda: Path(_env("MATOMAIL_DB_BACKUP_DIR", "./data/backups"))
    )
    db_max_size_mb: float = field(
        default_factory=lambda: _env_float("MATOMAIL_DB_MAX_SIZE_MB", 512.0)
    )
    store_email_body: bool = field(
        default_factory=lambda: _env_bool("MATOMAIL_STORE_EMAIL_BODY", True)
    )
    timezone: str = field(default_factory=lambda: _env("MATOMAIL_TIMEZONE", "Asia/Tokyo"))
    max_emails_per_run: int = field(
        default_factory=lambda: _env_int("MATOMAIL_MAX_EMAILS_PER_RUN", 30)
    )
    auto_open_report: bool = field(
        default_factory=lambda: _env_bool("MATOMAIL_AUTO_OPEN_REPORT", True)
    )
    download_attachments: bool = field(
        default_factory=lambda: _env_bool("MATOMAIL_DOWNLOAD_ATTACHMENTS", False)
    )
    send_email_without_confirmation: bool = field(
        default_factory=lambda: _env_bool(
            "MATOMAIL_SEND_EMAIL_WITHOUT_CONFIRMATION", False
        )
    )
    create_calendar_without_confirmation: bool = field(
        default_factory=lambda: _env_bool(
            "MATOMAIL_CREATE_CALENDAR_WITHOUT_CONFIRMATION", False
        )
    )
    google_client_secrets_file: Path = field(
        default_factory=lambda: Path(_env("GOOGLE_CLIENT_SECRETS_FILE", "./credentials.json"))
    )
    google_token_file: Path = field(
        default_factory=lambda: Path(_env("GOOGLE_TOKEN_FILE", "./token.json"))
    )
    google_oauth_port: int = field(default_factory=lambda: _env_int("GOOGLE_OAUTH_PORT", 8080))
    openai_api_key: str = field(default_factory=lambda: _env("OPENAI_API_KEY", ""))
    llm_model: str = field(default_factory=lambda: _env("MATOMAIL_LLM_MODEL", "gpt-5.4-mini"))
