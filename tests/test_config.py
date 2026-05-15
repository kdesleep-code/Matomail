from matomail import __version__
from matomail.config import Settings


def test_version_is_defined() -> None:
    assert __version__ == "0.1.0"


def test_default_settings_match_initial_requirements() -> None:
    settings = Settings()

    assert settings.lookback_days == 7
    assert str(settings.rules_db_path).replace("\\", "/") == "data/matomail_rules.sqlite3"
    assert settings.timezone == "Asia/Tokyo"
    assert settings.max_emails_per_run == 30
    assert settings.download_attachments is False
    assert settings.send_email_without_confirmation is False
    assert settings.create_calendar_without_confirmation is False
