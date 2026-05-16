"""Local web app for Matomail."""

from __future__ import annotations

import mimetypes
import json
import threading
from contextlib import asynccontextmanager
from dataclasses import dataclass
from datetime import UTC, datetime, timedelta, timezone as fixed_timezone
from email.utils import getaddresses
from pathlib import Path
from urllib.parse import parse_qs
from uuid import uuid4
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from fastapi import BackgroundTasks, FastAPI, Request
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse, Response
from googleapiclient.errors import HttpError
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from sqlalchemy import select

from .calendar_client import CALENDAR_EVENTS_SCOPE, CalendarClient
from .config import Settings
from .database import FILTER_ACTION_IGNORE
from .database import FILTER_ACTION_PRECLASSIFY, FILTER_ACTION_SKIP_ANALYSIS
from .database import ProcessingStateRecord
from .gmail_client import GMAIL_READONLY_SCOPE, GMAIL_SEND_SCOPE, GmailClient
from .llm_client import LLMClient
from .models import EmailMessage
from .report import PRIORITY_RANK, ReportEmail, ReportGenerator, _active_list_emails
from .workflow import create_mail_database, create_rules_database, generate_report
from .workflow import LoadMailResult, load_today_mail


GENERATION_INSTRUCTION_KEYS = {
    "compose": "llm_instruction.compose",
    "reply": "llm_instruction.reply",
}
GENERATION_INSTRUCTION_LABELS = {
    "compose": "新規メールの文面を生成するとき",
    "reply": "返信メールの文面を生成するとき",
}


settings = Settings()
settings.report_dir.mkdir(parents=True, exist_ok=True)


def refresh_report_on_startup() -> None:
    try:
        generate_report(settings, open_browser=False)
    except Exception:
        pass


@asynccontextmanager
async def lifespan(app: FastAPI):
    refresh_report_on_startup()
    yield


app = FastAPI(title="Matomail", lifespan=lifespan)
templates = Jinja2Templates(directory=str(Path(__file__).parent / "templates"))
app.mount("/reports", StaticFiles(directory=str(settings.report_dir)), name="reports")


@dataclass
class RunState:
    run_id: str
    percent: int = 0
    stage: str = "queued"
    message: str = "待機しています"
    done: bool = False
    error: str = ""
    report_url: str = ""
    fetched_count: int = 0
    processable_count: int = 0
    analyzed_count: int = 0


@dataclass
class InstructionFormSnapshot:
    id: int = 0
    name: str = ""
    instruction: str = ""
    from_query: str = ""
    to_query: str = ""
    subject_query: str = ""
    has_words: str = ""
    doesnt_have: str = ""
    note: str = ""
    enabled: bool = True


@dataclass(frozen=True)
class PriorityRuleView:
    key: str
    id: int
    name: str
    enabled: bool
    rule_type: str
    action: str
    instruction: str
    priority_value: str
    from_query: str
    to_query: str
    subject_query: str
    has_words: str
    doesnt_have: str
    note: str
    edit_url: str
    delete_url: str


_runs: dict[str, RunState] = {}
_runs_lock = threading.Lock()


@app.get("/", response_class=HTMLResponse)
def home(request: Request) -> HTMLResponse:
    return templates.TemplateResponse(
        request,
        "web_home.html.j2",
        {
            "latest_report_url": _latest_report_url(),
        },
    )


@app.post("/runs")
def start_run(background_tasks: BackgroundTasks) -> RedirectResponse:
    run_id = uuid4().hex
    with _runs_lock:
        _runs[run_id] = RunState(run_id=run_id)
    background_tasks.add_task(_run_load_mail, run_id)
    return RedirectResponse(f"/runs/{run_id}", status_code=303)


@app.get("/runs/{run_id}", response_class=HTMLResponse)
def run_page(request: Request, run_id: str) -> HTMLResponse:
    state = _get_run(run_id)
    return templates.TemplateResponse(
        request,
        "web_run.html.j2",
        {
            "run": state,
        },
    )


@app.get("/runs/{run_id}/status")
def run_status(run_id: str) -> JSONResponse:
    state = _get_run(run_id)
    return JSONResponse(
        {
            "run_id": state.run_id,
            "percent": state.percent,
            "stage": state.stage,
            "message": state.message,
            "done": state.done,
            "error": state.error,
            "report_url": state.report_url,
            "fetched_count": state.fetched_count,
            "processable_count": state.processable_count,
            "analyzed_count": state.analyzed_count,
        }
    )


@app.post("/reply-drafts")
async def create_reply_draft(request: Request) -> JSONResponse:
    try:
        payload = await request.json()
    except Exception:
        return JSONResponse({"error": "invalid JSON"}, status_code=400)

    message_id = str(payload.get("message_id", "")).strip()
    policy = str(payload.get("policy", "")).strip()
    if not message_id:
        return JSONResponse({"error": "message_id is required"}, status_code=400)

    database = create_mail_database(settings)
    message = database.get_email(message_id)
    if message is None:
        return JSONResponse({"error": "message not found"}, status_code=404)
    if "SENT" in set(message.label_ids):
        return JSONResponse(
            {"error": "reply drafts are only available for received mail"},
            status_code=400,
        )

    prompt = _build_reply_draft_prompt(message, policy)
    try:
        draft = LLMClient.from_settings(settings).generate_text(prompt).strip()
    except ValueError as error:
        return JSONResponse({"error": str(error)}, status_code=500)
    except Exception as error:
        return JSONResponse(
            {"error": "failed to generate reply draft", "detail": str(error)},
            status_code=502,
        )
    if not draft:
        return JSONResponse({"error": "empty reply draft"}, status_code=502)
    return JSONResponse({"draft": draft})


@app.get("/replies/new", response_class=HTMLResponse)
def new_reply_page(request: Request, message_id: str) -> HTMLResponse:
    database = create_mail_database(settings)
    message = database.get_email(message_id)
    return_url = _safe_return_url(
        request.query_params.get("return_url") or request.headers.get("referer", "")
    )
    if message is None:
        return templates.TemplateResponse(
            request,
            "web_reply_form.html.j2",
            {
                "message": None,
                "error": "指定されたメールが見つかりません。",
                "sent": False,
                "return_url": return_url,
            },
            status_code=404,
        )
    if "SENT" in set(message.label_ids):
        return templates.TemplateResponse(
            request,
            "web_reply_form.html.j2",
            {
                "message": message,
                "error": "送信メールには返信できません。",
                "sent": False,
                "return_url": return_url,
            },
            status_code=400,
        )

    defaults = _reply_defaults(message, database)
    draft = str(request.query_params.get("draft", "")).strip()
    return templates.TemplateResponse(
        request,
        "web_reply_form.html.j2",
        {
            "message": message,
            "error": "",
            "sent": False,
            "to_value": defaults["to"],
            "cc_value": defaults["cc"],
            "bcc_value": "",
            "body_value": _reply_body_prefill(message, draft),
            "return_url": return_url,
        },
    )


@app.post("/replies/new", response_class=HTMLResponse)
async def new_reply_page_from_draft(request: Request) -> HTMLResponse:
    form = parse_qs((await request.body()).decode("utf-8"))
    message_id = _form_value(form, "message_id")
    database = create_mail_database(settings)
    message = database.get_email(message_id)
    return_url = _safe_return_url(_form_value(form, "return_url"))
    if message is None:
        return templates.TemplateResponse(
            request,
            "web_reply_form.html.j2",
            {
                "message": None,
                "error": "指定されたメールが見つかりません。",
                "sent": False,
                "return_url": return_url,
            },
            status_code=404,
        )
    defaults = _reply_defaults(message, database)
    return templates.TemplateResponse(
        request,
        "web_reply_form.html.j2",
        {
            "message": message,
            "error": "",
            "sent": False,
            "to_value": defaults["to"],
            "cc_value": defaults["cc"],
            "bcc_value": "",
            "body_value": _reply_body_prefill(message, _form_value(form, "draft")),
            "return_url": return_url,
        },
    )


@app.post("/replies/send", response_class=HTMLResponse)
async def send_reply(request: Request) -> HTMLResponse:
    form = parse_qs((await request.body()).decode("utf-8"))
    message_id = _form_value(form, "message_id")
    database = create_mail_database(settings)
    message = database.get_email(message_id)
    if message is None:
        return templates.TemplateResponse(
            request,
            "web_reply_form.html.j2",
            {
                "message": None,
                "error": "指定されたメールが見つかりません。",
                "sent": False,
            },
            status_code=404,
        )

    to_value = _form_value(form, "to")
    cc_value = _form_value(form, "cc")
    bcc_value = _form_value(form, "bcc")
    body_value = _form_value(form, "body")
    return_url = _safe_return_url(_form_value(form, "return_url"))
    to_addresses = _address_list(to_value)
    if not to_addresses:
        return templates.TemplateResponse(
            request,
            "web_reply_form.html.j2",
            {
                "message": message,
                "error": "相手アドレスを入力してください。",
                "sent": False,
                "to_value": to_value,
                "cc_value": cc_value,
                "bcc_value": bcc_value,
                "body_value": body_value,
                "return_url": return_url,
            },
            status_code=400,
        )
    if not body_value.strip():
        return templates.TemplateResponse(
            request,
            "web_reply_form.html.j2",
            {
                "message": message,
                "error": "本文を入力してください。",
                "sent": False,
                "to_value": to_value,
                "cc_value": cc_value,
                "bcc_value": bcc_value,
                "body_value": body_value,
                "return_url": return_url,
            },
            status_code=400,
        )

    try:
        try:
            _send_reply_via_gmail(
                message=message,
                to_addresses=tuple(to_addresses),
                cc_addresses=tuple(_address_list(cc_value)),
                bcc_addresses=tuple(_address_list(bcc_value)),
                body=body_value,
            )
        except HttpError as error:
            if not _is_insufficient_auth_scope(error):
                raise
            _send_reply_via_gmail(
                message=message,
                to_addresses=tuple(to_addresses),
                cc_addresses=tuple(_address_list(cc_value)),
                bcc_addresses=tuple(_address_list(bcc_value)),
                body=body_value,
                force_consent=True,
            )
    except Exception as error:
        return templates.TemplateResponse(
            request,
            "web_reply_form.html.j2",
            {
                "message": message,
                "error": _reply_send_error_message(error),
                "sent": False,
                "to_value": to_value,
                "cc_value": cc_value,
                "bcc_value": bcc_value,
                "body_value": body_value,
                "return_url": return_url,
            },
            status_code=200,
        )

    return RedirectResponse(return_url or f"/replies/new?message_id={message_id}", status_code=303)


@app.post("/replies/schedule", response_class=HTMLResponse)
async def schedule_reply_page(request: Request) -> HTMLResponse:
    form = parse_qs((await request.body()).decode("utf-8"))
    database = create_mail_database(settings)
    message = database.get_email(_form_value(form, "message_id"))
    return templates.TemplateResponse(
        request,
        "web_reply_schedule.html.j2",
        {
            "message": message,
            "to_value": _form_value(form, "to"),
            "cc_value": _form_value(form, "cc"),
            "bcc_value": _form_value(form, "bcc"),
            "body_value": _form_value(form, "body"),
            "error": "" if message else "指定されたメールが見つかりません。",
        },
        status_code=200 if message else 404,
    )


@app.get("/compose", response_class=HTMLResponse)
def compose_page(request: Request) -> HTMLResponse:
    return templates.TemplateResponse(
        request,
        "web_compose.html.j2",
        {"error": "", "sent": False, "to_value": "", "cc_value": "", "bcc_value": "", "subject_value": "", "body_value": ""},
    )


@app.post("/compose", response_class=HTMLResponse)
async def send_new_message(request: Request) -> HTMLResponse:
    form = parse_qs((await request.body()).decode("utf-8"))
    values = {
        "to_value": _form_value(form, "to"),
        "cc_value": _form_value(form, "cc"),
        "bcc_value": _form_value(form, "bcc"),
        "subject_value": _form_value(form, "subject"),
        "body_value": _form_value(form, "body"),
    }
    to_addresses = _address_list(values["to_value"])
    if not to_addresses or not values["body_value"].strip():
        return templates.TemplateResponse(
            request,
            "web_compose.html.j2",
            {**values, "error": "宛先と本文を入力してください。", "sent": False},
            status_code=400,
        )
    try:
        GmailClient.from_oauth(
            settings,
            scopes=[GMAIL_READONLY_SCOPE, GMAIL_SEND_SCOPE],
        ).send_message(
            to=tuple(to_addresses),
            cc=tuple(_address_list(values["cc_value"])),
            bcc=tuple(_address_list(values["bcc_value"])),
            subject=values["subject_value"],
            body=values["body_value"],
        )
    except Exception as error:
        return templates.TemplateResponse(
            request,
            "web_compose.html.j2",
            {**values, "error": _reply_send_error_message(error), "sent": False},
            status_code=200,
        )
    return RedirectResponse(_latest_report_url() or "/", status_code=303)


@app.get("/singletons")
def singleton_messages_page() -> RedirectResponse:
    return RedirectResponse("/unhandled", status_code=303)


@app.get("/completed")
def completed_messages_page() -> RedirectResponse:
    latest_url = _latest_report_url() or "/"
    if latest_url == "/":
        return RedirectResponse(latest_url, status_code=303)
    return RedirectResponse(f"{latest_url}#completed", status_code=303)


@app.get("/unhandled", response_class=HTMLResponse)
def unhandled_messages_page(request: Request) -> HTMLResponse:
    database = create_mail_database(settings)
    messages = _unhandled_priority_messages(database)
    return templates.TemplateResponse(
        request,
        "web_unhandled.html.j2",
        {"messages": messages},
    )


@app.post("/messages/{message_id}/opened")
def mark_message_opened(message_id: str) -> JSONResponse:
    database = create_mail_database(settings)
    if database.get_email(message_id) is None:
        return JSONResponse({"ok": False, "error": "message not found"}, status_code=404)
    database.mark_web_opened(message_id)
    generate_report(settings, open_browser=False)
    return JSONResponse({"ok": True})


@app.post("/messages/{message_id}/resolved")
async def set_message_resolved(request: Request, message_id: str) -> JSONResponse:
    database = create_mail_database(settings)
    if database.get_email(message_id) is None:
        return JSONResponse({"ok": False, "error": "message not found"}, status_code=404)
    try:
        payload = await request.json()
    except Exception:
        payload = {}
    database.set_resolved(message_id, bool(payload.get("resolved")))
    generate_report(settings, open_browser=False)
    return JSONResponse({"ok": True})


@app.post("/calendar-events")
async def create_calendar_event(request: Request) -> Response:
    form = parse_qs((await request.body()).decode("utf-8"))
    message_id = _form_value(form, "message_id")
    return_url = _safe_return_url(_form_value(form, "return_url")) or "/reports/index.html"
    database = create_mail_database(settings)
    message = database.get_email(message_id)
    if message is None:
        return HTMLResponse("message not found", status_code=404)

    title = _form_value(form, "title").strip() or message.subject or "Matomail event"
    start_time = _form_value(form, "start_time").strip()
    end_time = _form_value(form, "end_time").strip()
    timezone = _form_value(form, "timezone").strip() or getattr(settings, "timezone", "Asia/Tokyo")
    location = _form_value(form, "location").strip()
    attendees = tuple(_address_list(_form_value(form, "attendees")))
    description = _form_value(form, "description").strip()
    description = _calendar_description_with_source_url(description, message)
    if not start_time or not end_time:
        return HTMLResponse("start_time and end_time are required", status_code=400)

    try:
        event = _create_calendar_event_via_google(
            title=title,
            start_time=start_time,
            end_time=end_time,
            timezone=timezone,
            location=location,
            attendees=attendees,
            description=description,
        )
    except HttpError as error:
        if _is_insufficient_auth_scope(error):
            try:
                event = _create_calendar_event_via_google(
                    title=title,
                    start_time=start_time,
                    end_time=end_time,
                    timezone=timezone,
                    location=location,
                    attendees=attendees,
                    description=description,
                    force_consent=True,
                )
            except Exception as retry_error:
                return HTMLResponse(_calendar_error_message(retry_error), status_code=502)
        else:
            return HTMLResponse(_calendar_error_message(error), status_code=502)
    except Exception as error:
        return HTMLResponse(_calendar_error_message(error), status_code=500)

    database.save_calendar_event(
        message_id,
        title=title,
        start_time=_parse_calendar_datetime(start_time),
        end_time=_parse_calendar_datetime(end_time),
        timezone=timezone,
        location=location,
        attendees=attendees,
        calendar_event_id=str(event.get("id", "")),
    )
    database.set_meeting_candidates(
        message_id,
        [
            {
                "title": title,
                "start_time": start_time,
                "end_time": end_time,
                "timezone": timezone,
                "location": location,
                "attendees": list(attendees),
                "description": description,
            }
        ],
        hidden=False,
    )
    generate_report(settings, open_browser=False)
    return RedirectResponse(return_url, status_code=303)


@app.post("/calendar-candidates/extract")
async def extract_calendar_candidates(request: Request) -> Response:
    form = parse_qs((await request.body()).decode("utf-8"))
    message_id = _form_value(form, "message_id")
    return_url = _safe_return_url(_form_value(form, "return_url")) or "/reports/index.html"
    force = _form_value(form, "force") == "1"
    database = create_mail_database(settings)
    message = database.get_email(message_id)
    if message is None:
        return HTMLResponse("message not found", status_code=404)
    if not force and database.get_meeting_candidates(message_id):
        database.set_calendar_candidates_hidden(message_id, False)
        generate_report(settings, open_browser=False)
        return RedirectResponse(return_url, status_code=303)
    thread_messages = [
        item
        for item in database.list_saved_emails()
        if (item.gmail_thread_id or item.gmail_message_id)
        == (message.gmail_thread_id or message.gmail_message_id)
    ]
    prompt = _calendar_candidate_prompt(message, thread_messages)
    try:
        raw = LLMClient.from_settings(settings).generate_text(prompt)
        candidates = _parse_calendar_candidates(raw)
    except Exception as error:
        return HTMLResponse(f"カレンダー候補の抽出に失敗しました: {error}", status_code=500)
    database.set_meeting_candidates(message_id, candidates, hidden=False)
    generate_report(settings, open_browser=False)
    return RedirectResponse(return_url, status_code=303)


@app.post("/calendar-candidates/clear")
async def clear_calendar_candidates(request: Request) -> Response:
    form = parse_qs((await request.body()).decode("utf-8"))
    message_id = _form_value(form, "message_id")
    return_url = _safe_return_url(_form_value(form, "return_url")) or "/reports/index.html"
    database = create_mail_database(settings)
    if database.get_email(message_id) is None:
        return HTMLResponse("message not found", status_code=404)
    database.set_calendar_candidates_hidden(message_id, True)
    generate_report(settings, open_browser=False)
    return RedirectResponse(return_url, status_code=303)


@app.post("/calendar-events/{event_record_id}/cancel")
async def cancel_calendar_event(request: Request, event_record_id: int) -> Response:
    form = parse_qs((await request.body()).decode("utf-8"))
    return_url = _safe_return_url(_form_value(form, "return_url")) or "/reports/index.html"
    database = create_mail_database(settings)
    event_record = database.get_calendar_event(event_record_id)
    if event_record is not None and event_record.calendar_event_id:
        try:
            _delete_calendar_event_via_google(event_record.calendar_event_id)
        except HttpError as error:
            if not _is_missing_calendar_event(error):
                return HTMLResponse(_calendar_error_message(error), status_code=502)
        except Exception as error:
            return HTMLResponse(_calendar_error_message(error), status_code=500)
    cancelled_message_id = database.mark_calendar_event_cancelled(event_record_id)
    if cancelled_message_id:
        database.set_calendar_candidates_hidden(cancelled_message_id, True)
    generate_report(settings, open_browser=False)
    return RedirectResponse(return_url, status_code=303)


@app.get("/generation-instructions", response_class=HTMLResponse)
def generation_instructions_page(request: Request) -> HTMLResponse:
    database = create_mail_database(settings)
    return templates.TemplateResponse(
        request,
        "web_generation_instructions.html.j2",
        {
            "items": [
                {
                    "kind": kind,
                    "label": GENERATION_INSTRUCTION_LABELS[kind],
                    "instruction": database.get_app_setting(key),
                    "edit_url": f"/generation-instructions/{kind}/edit",
                }
                for kind, key in GENERATION_INSTRUCTION_KEYS.items()
            ],
        },
    )


@app.get("/generation-instructions/{kind}/edit", response_class=HTMLResponse)
def edit_generation_instruction_page(request: Request, kind: str) -> HTMLResponse:
    if kind not in GENERATION_INSTRUCTION_KEYS:
        return templates.TemplateResponse(
            request,
            "web_generation_instruction_edit.html.j2",
            {
                "kind": kind,
                "label": "",
                "instruction": "",
                "error": "指定された追加指示が見つかりません。",
            },
            status_code=404,
        )
    database = create_mail_database(settings)
    return templates.TemplateResponse(
        request,
        "web_generation_instruction_edit.html.j2",
        {
            "kind": kind,
            "label": GENERATION_INSTRUCTION_LABELS[kind],
            "instruction": database.get_app_setting(GENERATION_INSTRUCTION_KEYS[kind]),
            "error": "",
        },
    )


@app.post("/generation-instructions/{kind}/edit", response_class=HTMLResponse)
async def update_generation_instruction(request: Request, kind: str) -> HTMLResponse:
    form = parse_qs((await request.body()).decode("utf-8"))
    if kind not in GENERATION_INSTRUCTION_KEYS:
        return templates.TemplateResponse(
            request,
            "web_generation_instruction_edit.html.j2",
            {
                "kind": kind,
                "label": "",
                "instruction": _form_value(form, "instruction"),
                "error": "指定された追加指示が見つかりません。",
            },
            status_code=404,
        )
    database = create_mail_database(settings)
    database.set_app_setting(
        GENERATION_INSTRUCTION_KEYS[kind],
        _form_value(form, "instruction").strip(),
    )
    return RedirectResponse("/generation-instructions", status_code=303)


@app.post("/replies/schedule/confirm", response_class=HTMLResponse)
async def schedule_reply_confirm(request: Request) -> HTMLResponse:
    form = parse_qs((await request.body()).decode("utf-8"))
    database = create_mail_database(settings)
    message = database.get_email(_form_value(form, "message_id"))
    return templates.TemplateResponse(
        request,
        "web_reply_schedule.html.j2",
        {
            "message": message,
            "to_value": _form_value(form, "to"),
            "cc_value": _form_value(form, "cc"),
            "bcc_value": _form_value(form, "bcc"),
            "body_value": _form_value(form, "body"),
            "scheduled_at": _form_value(form, "scheduled_at"),
            "error": "Gmail APIはGmail本体の送信予約機能の作成に対応していません。現時点ではこの画面から予約送信はできません。",
        },
        status_code=501,
    )


@app.get("/attachments/{message_id}/{attachment_index}")
def download_attachment(message_id: str, attachment_index: int) -> Response:
    database = create_mail_database(settings)
    message = database.get_email(message_id)
    if message is None:
        return JSONResponse({"error": "message not found"}, status_code=404)
    if attachment_index < 0 or attachment_index >= len(message.attachments):
        return JSONResponse({"error": "attachment not found"}, status_code=404)

    attachment = message.attachments[attachment_index]
    filename = attachment.filename or f"attachment-{attachment_index + 1}"
    cache_path = _attachment_cache_path(message_id, attachment_index, filename)
    if not cache_path.exists():
        if not attachment.attachment_id:
            return JSONResponse({"error": "attachment body is not available"}, status_code=404)
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        try:
            data = GmailClient.from_oauth(settings).download_attachment(
                message_id=message_id,
                attachment_id=attachment.attachment_id,
            )
        except HttpError as error:
            return JSONResponse(
                {
                    "error": "failed to download attachment from Gmail",
                    "detail": str(error),
                },
                status_code=502,
            )
        except Exception as error:
            return JSONResponse(
                {
                    "error": "failed to download attachment",
                    "detail": str(error),
                },
                status_code=500,
            )
        cache_path.write_bytes(data)

    media_type = attachment.mime_type or mimetypes.guess_type(filename)[0]
    return FileResponse(cache_path, media_type=media_type, filename=filename)


@app.get("/rules/new", response_class=HTMLResponse)
def new_rule_page(request: Request, message_id: str) -> HTMLResponse:
    database = create_mail_database(settings)
    message = database.get_email(message_id)
    if message is None:
        return templates.TemplateResponse(
            request,
            "web_rule_form.html.j2",
            {
                "message": None,
                "error": "指定されたメールが見つかりません。",
                "saved": False,
            },
            status_code=404,
        )
    return templates.TemplateResponse(
        request,
        "web_rule_form.html.j2",
        {
            "message": message,
            "sender_rule_value": _sender_rule_value(message),
            "body_excerpt": _body_rule_excerpt(message.body),
            "error": "",
            "saved": False,
        },
    )


@app.get("/rules", response_class=HTMLResponse)
def rules_page(request: Request) -> HTMLResponse:
    rules_database = create_rules_database(settings)
    return templates.TemplateResponse(
        request,
        "web_rules.html.j2",
        {
            "rules": _priority_rule_views(rules_database),
        },
    )


@app.get("/instructions", response_class=HTMLResponse)
def instructions_page() -> RedirectResponse:
    return RedirectResponse("/rules", status_code=303)


@app.get("/instructions/new", response_class=HTMLResponse)
def new_instruction_page(request: Request) -> HTMLResponse:
    return templates.TemplateResponse(
        request,
        "web_instruction_form.html.j2",
        {
            "rule": None,
            "error": "",
            "mode": "new",
        },
    )


@app.post("/instructions/new", response_class=HTMLResponse)
async def create_instruction(request: Request) -> HTMLResponse:
    form = parse_qs((await request.body()).decode("utf-8"), keep_blank_values=True)
    rules_database = create_rules_database(settings)
    try:
        rules_database.add_llm_instruction_rule(
            name=_form_value(form, "name"),
            instruction=_form_value(form, "instruction"),
            from_query=_form_value(form, "from_query"),
            to_query=_form_value(form, "to_query"),
            subject_query=_form_value(form, "subject_query"),
            has_words=_form_value(form, "has_words"),
            doesnt_have=_form_value(form, "doesnt_have"),
            note=_form_value(form, "note"),
            enabled="enabled" in form,
        )
    except ValueError as error:
        return templates.TemplateResponse(
            request,
            "web_instruction_form.html.j2",
            {
                "rule": _instruction_form_snapshot(form),
                "error": str(error),
                "mode": "new",
            },
            status_code=400,
        )
    return RedirectResponse("/rules", status_code=303)


@app.get("/instructions/{rule_id}/edit", response_class=HTMLResponse)
def edit_instruction_page(request: Request, rule_id: int) -> HTMLResponse:
    rules_database = create_rules_database(settings)
    rule = rules_database.get_llm_instruction_rule(rule_id)
    return templates.TemplateResponse(
        request,
        "web_instruction_form.html.j2",
        {
            "rule": rule,
            "error": "" if rule else "指定された追加指示が見つかりません。",
            "mode": "edit",
        },
        status_code=200 if rule else 404,
    )


@app.post("/instructions/{rule_id}/edit", response_class=HTMLResponse)
async def update_instruction(request: Request, rule_id: int) -> HTMLResponse:
    form = parse_qs((await request.body()).decode("utf-8"), keep_blank_values=True)
    rules_database = create_rules_database(settings)
    existing_rule = rules_database.get_llm_instruction_rule(rule_id)
    if existing_rule is None:
        return templates.TemplateResponse(
            request,
            "web_instruction_form.html.j2",
            {
                "rule": None,
                "error": "指定された追加指示が見つかりません。",
                "mode": "edit",
            },
            status_code=404,
        )
    try:
        ok = rules_database.update_llm_instruction_rule(
            rule_id,
            name=_form_value(form, "name"),
            instruction=_form_value(form, "instruction"),
            priority=existing_rule.priority,
            from_query=_form_value(form, "from_query"),
            to_query=_form_value(form, "to_query"),
            subject_query=_form_value(form, "subject_query"),
            has_words=_form_value(form, "has_words"),
            doesnt_have=_form_value(form, "doesnt_have"),
            note=_form_value(form, "note"),
            enabled="enabled" in form,
        )
    except ValueError as error:
        snapshot = _instruction_form_snapshot(form)
        snapshot.id = rule_id
        return templates.TemplateResponse(
            request,
            "web_instruction_form.html.j2",
            {
                "rule": snapshot,
                "error": str(error),
                "mode": "edit",
            },
            status_code=400,
        )
    if not ok:
        return RedirectResponse("/rules", status_code=303)
    return RedirectResponse("/rules", status_code=303)


@app.post("/instructions/{rule_id}/delete")
def delete_instruction(rule_id: int) -> RedirectResponse:
    rules_database = create_rules_database(settings)
    rules_database.delete_llm_instruction_rule(rule_id)
    return RedirectResponse("/rules", status_code=303)


@app.post("/instructions/reorder")
async def reorder_instructions(request: Request) -> JSONResponse:
    payload = await request.json()
    ordered_ids = [int(rule_id) for rule_id in payload.get("rule_ids", [])]
    rules_database = create_rules_database(settings)
    ok = rules_database.reorder_llm_instruction_rules(ordered_ids)
    return JSONResponse({"ok": ok}, status_code=200 if ok else 400)


@app.get("/rules/{rule_id}/edit", response_class=HTMLResponse)
def edit_rule_page(request: Request, rule_id: int) -> HTMLResponse:
    rules_database = create_rules_database(settings)
    rule = rules_database.get_filter_rule(rule_id)
    return templates.TemplateResponse(
        request,
        "web_rule_edit.html.j2",
        {
            "rule": rule,
            "error": "" if rule else "指定されたルールが見つかりません。",
        },
        status_code=200 if rule else 404,
    )


@app.post("/rules/{rule_id}/edit", response_class=HTMLResponse)
async def update_rule(request: Request, rule_id: int) -> HTMLResponse:
    form = parse_qs((await request.body()).decode("utf-8"), keep_blank_values=True)
    rules_database = create_rules_database(settings)
    existing_rule = rules_database.get_filter_rule(rule_id)
    action_choice = _form_value(form, "action", FILTER_ACTION_SKIP_ANALYSIS)
    if action_choice == "preclassify_top":
        action = FILTER_ACTION_PRECLASSIFY
        preset_priority = "top"
        preset_category = "top"
    elif action_choice == "preclassify" or action_choice == "preclassify_high":
        action = FILTER_ACTION_PRECLASSIFY
        preset_priority = "high"
        preset_category = "important"
    elif action_choice in {"preclassify_middle", "preclassify_medium"}:
        action = FILTER_ACTION_PRECLASSIFY
        preset_priority = "medium"
        preset_category = "middle"
    elif action_choice == "preclassify_low":
        action = FILTER_ACTION_PRECLASSIFY
        preset_priority = "low"
        preset_category = "low"
    else:
        action = FILTER_ACTION_SKIP_ANALYSIS
        preset_priority = ""
        preset_category = ""
    ok = rules_database.update_filter_rule(
        rule_id,
        action=action,
        name=_form_value(form, "name"),
        priority=existing_rule.priority if existing_rule else 0,
        preset_priority=preset_priority,
        preset_category=preset_category,
        preset_summary_ja="Web画面から編集した高優先度ルールです。"
        if action == FILTER_ACTION_PRECLASSIFY
        else "",
        preset_suggested_action_ja="優先して確認する。"
        if action == FILTER_ACTION_PRECLASSIFY
        else "",
        from_query=_form_value(form, "from_query"),
        subject_query=_form_value(form, "subject_query"),
        has_words=_form_value(form, "has_words"),
        note=_form_value(form, "note"),
        enabled="enabled" in form,
    )
    _refresh_rules_and_report()
    if not ok:
        return templates.TemplateResponse(
            request,
            "web_rule_edit.html.j2",
            {
                "rule": None,
                "error": "指定されたルールが見つかりません。",
            },
            status_code=404,
        )
    return RedirectResponse("/rules", status_code=303)


@app.post("/rules/{rule_id}/delete")
def delete_rule(rule_id: int) -> RedirectResponse:
    rules_database = create_rules_database(settings)
    rules_database.delete_filter_rule(rule_id)
    _refresh_rules_and_report()
    return RedirectResponse("/rules", status_code=303)


@app.post("/rules/{rule_id}/move")
async def move_rule(request: Request, rule_id: int) -> RedirectResponse:
    form = parse_qs((await request.body()).decode("utf-8"), keep_blank_values=True)
    direction = _form_value(form, "direction")
    rules_database = create_rules_database(settings)
    rules_database.move_filter_rule(rule_id, direction)
    _refresh_rules_and_report()
    return RedirectResponse("/rules", status_code=303)


@app.post("/rules/reorder")
async def reorder_rules(request: Request) -> JSONResponse:
    payload = await request.json()
    ordered_keys = [
        raw if ":" in raw else f"filter:{raw}"
        for raw in (str(rule_id) for rule_id in payload.get("rule_ids", []))
    ]
    rules_database = create_rules_database(settings)
    ok = rules_database.reorder_priority_rules(ordered_keys)
    if ok:
        _refresh_rules_and_report()
    return JSONResponse({"ok": ok}, status_code=200 if ok else 400)


@app.post("/rules/new", response_class=HTMLResponse)
async def create_rule(request: Request) -> HTMLResponse:
    form = parse_qs((await request.body()).decode("utf-8"), keep_blank_values=True)
    message_id = _form_value(form, "message_id")
    database = create_mail_database(settings)
    message = database.get_email(message_id)
    if message is None:
        return templates.TemplateResponse(
            request,
            "web_rule_form.html.j2",
            {
                "message": None,
                "error": "指定されたメールが見つかりません。",
                "saved": False,
            },
            status_code=404,
        )

    from_query = _form_value(form, "from_query").strip() if "use_from" in form else ""
    subject_query = (
        _form_value(form, "subject_query").strip() if "use_subject" in form else ""
    )
    has_words = _form_value(form, "body_query").strip() if "use_body" in form else ""
    rule_mode = _form_value(form, "rule_mode", "skip")
    instruction = _form_value(form, "instruction").strip()

    if not any([from_query, subject_query, has_words]):
        return templates.TemplateResponse(
            request,
            "web_rule_form.html.j2",
            {
                "message": message,
                "sender_rule_value": _sender_rule_value(message),
                "body_excerpt": _body_rule_excerpt(message.body),
                "error": "少なくとも1つの条件にチェックを入れてください。",
                "saved": False,
            },
            status_code=400,
        )

    rules_database = create_rules_database(settings)
    if rule_mode == "instruction":
        if not instruction:
            return templates.TemplateResponse(
                request,
                "web_rule_form.html.j2",
                {
                    "message": message,
                    "sender_rule_value": _sender_rule_value(message),
                    "body_excerpt": _body_rule_excerpt(message.body),
                    "error": "LLMへの追加指示を入力してください。",
                    "saved": False,
                },
                status_code=400,
            )
        rule = rules_database.add_llm_instruction_rule(
            name=f"LLM instruction: {message.subject[:60]}",
            instruction=instruction,
            from_query=from_query,
            subject_query=subject_query,
            has_words=has_words,
            note=f"Created from message {message.gmail_message_id}",
        )
        generate_report(settings, open_browser=False)
        return templates.TemplateResponse(
            request,
            "web_rule_form.html.j2",
            {
                "message": message,
                "sender_rule_value": _sender_rule_value(message),
                "body_excerpt": _body_rule_excerpt(message.body),
                "error": "",
                "saved": True,
                "rule": rule,
                "affected_count": 0,
            },
        )

    if rule_mode == "top":
        action = FILTER_ACTION_PRECLASSIFY
        preset_priority = "top"
        preset_category = "top"
        rule_name = f"Top: {message.subject[:60]}"
    elif rule_mode == "high":
        action = FILTER_ACTION_PRECLASSIFY
        preset_priority = "high"
        preset_category = "important"
        rule_name = f"High priority: {message.subject[:60]}"
    elif rule_mode in {"middle", "medium"}:
        action = FILTER_ACTION_PRECLASSIFY
        preset_priority = "medium"
        preset_category = "middle"
        rule_name = f"Middle priority: {message.subject[:60]}"
    elif rule_mode == "low":
        action = FILTER_ACTION_PRECLASSIFY
        preset_priority = "low"
        preset_category = "low"
        rule_name = f"Low priority: {message.subject[:60]}"
    else:
        action = FILTER_ACTION_SKIP_ANALYSIS
        preset_priority = ""
        preset_category = ""
        rule_name = f"Skip: {message.subject[:60]}"

    rule = rules_database.add_filter_rule(
        action=action,
        name=rule_name,
        priority=0,
        preset_priority=preset_priority,
        preset_category=preset_category,
        preset_summary_ja=(
            "Web画面から追加した最優先ルールです。"
            if rule_mode == "top"
            else "Web画面から追加した高優先度ルールです。"
            if rule_mode == "high"
            else ""
        ),
        preset_suggested_action_ja=(
            "一覧の最上位で確認する。"
            if rule_mode == "top"
            else "優先して確認する。"
            if rule_mode == "high"
            else ""
        ),
        from_query=from_query,
        subject_query=subject_query,
        has_words=has_words,
        note=f"Created from message {message.gmail_message_id}",
    )
    affected_count = _apply_rule_to_loaded_day(
        database,
        rules_database,
        message.gmail_message_id,
    )
    generate_report(settings, open_browser=False)
    return templates.TemplateResponse(
        request,
        "web_rule_form.html.j2",
        {
            "message": message,
            "sender_rule_value": _sender_rule_value(message),
            "body_excerpt": _body_rule_excerpt(message.body),
            "error": "",
            "saved": True,
            "rule": rule,
            "affected_count": affected_count,
        },
    )


def _run_load_mail(run_id: str) -> None:
    def progress(percent: int, stage: str, message: str) -> None:
        _update_run(run_id, percent=percent, stage=stage, message=message)

    try:
        result = load_today_mail(settings, progress=progress)
        _complete_run(run_id, result)
    except Exception as error:  # pragma: no cover - surfaced in the web UI
        _update_run(
            run_id,
            percent=100,
            stage="error",
            message="エラーが発生しました",
            error=str(error),
            done=True,
        )


def _complete_run(run_id: str, result: LoadMailResult) -> None:
    report_url = _report_url(result.report_path) if result.report_path else ""
    _update_run(
        run_id,
        percent=100,
        stage="done",
        message="本日のメール一覧を生成しました",
        done=True,
        report_url=report_url,
        fetched_count=result.fetched_count,
        processable_count=result.processable_count,
        analyzed_count=result.analyzed_count,
    )


def _get_run(run_id: str) -> RunState:
    with _runs_lock:
        return _runs.get(
            run_id,
            RunState(
                run_id=run_id,
                percent=100,
                stage="missing",
                message="指定された処理は見つかりません",
                done=True,
                error="run not found",
            ),
        )


def _update_run(run_id: str, **changes: object) -> None:
    with _runs_lock:
        state = _runs[run_id]
        for key, value in changes.items():
            setattr(state, key, value)


def _report_url(report_path: Path) -> str:
    try:
        relative = report_path.resolve().relative_to(settings.report_dir.resolve())
    except ValueError:
        relative = Path(report_path.name)
    return "/reports/" + relative.as_posix()


def _latest_report_url() -> str:
    index_path = settings.report_dir / "index.html"
    if index_path.exists():
        return "/reports/index.html"
    return ""


def _attachment_cache_path(message_id: str, attachment_index: int, filename: str) -> Path:
    safe_message_id = _safe_path_part(message_id)
    safe_filename = _safe_path_part(filename) or f"attachment-{attachment_index + 1}"
    return settings.attachment_cache_dir / safe_message_id / f"{attachment_index}-{safe_filename}"


def _safe_path_part(value: str) -> str:
    return "".join(
        character if character.isalnum() or character in "._-" else "_"
        for character in value
    ).strip("._")


def _body_rule_excerpt(body: str) -> str:
    lines = [line.strip() for line in body.splitlines() if line.strip()]
    text = " ".join(lines)
    return text[:500]


def _sender_rule_value(message) -> str:
    for _, address in getaddresses([*message.sender_candidates, message.sender]):
        if address:
            return address
    return message.sender.strip()


def _apply_rule_to_loaded_day(
    database,
    rules_database,
    gmail_message_id: str,
) -> int:
    affected_count = 0
    for message in database.list_emails_loaded_on_same_day(gmail_message_id):
        decision = rules_database.get_filter_decision(message)
        if decision is None:
            continue
        if decision["action"] not in {
            FILTER_ACTION_SKIP_ANALYSIS,
            FILTER_ACTION_PRECLASSIFY,
            FILTER_ACTION_IGNORE,
        }:
            continue
        database.save_filter_decision(
            message.gmail_message_id,
            action=decision["action"],
            matched_rule_id=decision["matched_rule_id"],
            matched_rule_name=decision["matched_rule_name"],
            reason=decision["reason"],
            rule_snapshot=decision["rule_snapshot"],
        )
        if decision["action"] == FILTER_ACTION_PRECLASSIFY:
            database.apply_preclassified_analysis(
                message.gmail_message_id,
                decision["preset_analysis"],
            )
        affected_count += 1
    return affected_count


def _refresh_rules_and_report() -> None:
    database = create_mail_database(settings)
    rules_database = create_rules_database(settings)
    database.clear_filter_decisions()
    for message in database.list_saved_emails():
        decision = rules_database.get_filter_decision(message)
        if decision is None:
            continue
        if decision["action"] not in {
            FILTER_ACTION_SKIP_ANALYSIS,
            FILTER_ACTION_PRECLASSIFY,
            FILTER_ACTION_IGNORE,
        }:
            continue
        database.save_filter_decision(
            message.gmail_message_id,
            action=decision["action"],
            matched_rule_id=decision["matched_rule_id"],
            matched_rule_name=decision["matched_rule_name"],
            reason=decision["reason"],
            rule_snapshot=decision["rule_snapshot"],
        )
        if decision["action"] == FILTER_ACTION_PRECLASSIFY:
            database.apply_preclassified_analysis(
                message.gmail_message_id,
                decision["preset_analysis"],
            )
    generate_report(settings, open_browser=False)


def _build_reply_draft_prompt(message: EmailMessage, policy: str) -> str:
    base_prompt = _load_reply_draft_prompt()
    policy_text = policy or "特に指定なし。元メールに対して自然で丁寧に返信してください。"
    additional_instruction = _generation_instruction_text("reply")
    recipients = ", ".join(message.recipients)
    cc = ", ".join(message.cc)
    body = (message.body or message.snippet or "").strip()
    if len(body) > 12000:
        body = body[:12000] + "\n\n[本文はここで省略されています]"
    return f"""{base_prompt}

以下の元メールに対する返信案を作成してください。
ユーザーが入力した返信方針に必ず従ってください。
出力はそのままメール本文として貼れる返信文だけにしてください。
件名、説明、Markdown、前置き、後書きは不要です。

返信の方針:
{policy_text}

元メール:
LLMへの追加指示:
{additional_instruction or "なし"}

From: {message.sender}
To: {recipients}
Cc: {cc}
Subject: {message.subject}
Received: {message.received_at.isoformat() if message.received_at else ""}
Snippet: {message.snippet}

本文:
{body}
"""


def _generation_instruction_text(kind: str) -> str:
    key = GENERATION_INSTRUCTION_KEYS.get(kind)
    if key is None:
        return ""
    database = create_mail_database(settings)
    return database.get_app_setting(key).strip()


def _load_reply_draft_prompt() -> str:
    prompt_path = Path(__file__).parent / "prompts" / "reply_draft.md"
    return prompt_path.read_text(encoding="utf-8").strip()


def _reply_send_error_message(error: Exception) -> str:
    if isinstance(error, HttpError):
        detail = ""
        try:
            payload = json.loads(error.content.decode("utf-8"))
            detail = payload.get("error", {}).get("message", "")
        except Exception:
            detail = error.content.decode("utf-8", errors="replace")
        if detail:
            return f"送信に失敗しました: {detail}"
    return f"送信に失敗しました: {error}"


def _is_insufficient_auth_scope(error: HttpError) -> bool:
    content = error.content.decode("utf-8", errors="replace").lower()
    return (
        "insufficient authentication scopes" in content
        or "access_token_scope_insufficient" in content
        or "insufficientpermissions" in content
    )


def _send_reply_via_gmail(
    *,
    message: EmailMessage,
    to_addresses: tuple[str, ...],
    cc_addresses: tuple[str, ...],
    bcc_addresses: tuple[str, ...],
    body: str,
    force_consent: bool = False,
) -> None:
    GmailClient.from_oauth(
        settings,
        scopes=[GMAIL_READONLY_SCOPE, GMAIL_SEND_SCOPE],
        force_consent=force_consent,
    ).send_reply(
        to=to_addresses,
        cc=cc_addresses,
        bcc=bcc_addresses,
        subject=message.subject,
        body=body,
        thread_id=message.gmail_thread_id,
        source_message_id=message.gmail_message_id,
    )


def _create_calendar_event_via_google(
    *,
    title: str,
    start_time: str,
    end_time: str,
    timezone: str,
    location: str,
    attendees: tuple[str, ...],
    description: str,
    force_consent: bool = False,
) -> dict:
    return CalendarClient.from_oauth(
        settings,
        scopes=[CALENDAR_EVENTS_SCOPE],
        force_consent=force_consent,
    ).create_event(
        title=title,
        start_time=start_time,
        end_time=end_time,
        timezone=timezone,
        location=location,
        attendees=attendees,
        description=description,
    )


def _delete_calendar_event_via_google(event_id: str) -> None:
    CalendarClient.from_oauth(
        settings,
        scopes=[CALENDAR_EVENTS_SCOPE],
    ).delete_event(event_id)


def _calendar_description_with_source_url(description: str, message: EmailMessage) -> str:
    gmail_id = message.gmail_thread_id or message.gmail_message_id
    if not gmail_id:
        return description
    source_url = f"https://mail.google.com/mail/u/0/#all/{gmail_id}"
    if source_url in description:
        return description
    prefix = description.strip()
    suffix = f"元メール: {source_url}"
    return f"{prefix}\n\n{suffix}" if prefix else suffix


def _is_missing_calendar_event(error: HttpError) -> bool:
    status = getattr(getattr(error, "resp", None), "status", None)
    return status in {404, 410}


def _calendar_candidate_prompt(
    target_message: EmailMessage,
    thread_messages: list[EmailMessage],
) -> str:
    ordered = sorted(
        thread_messages or [target_message],
        key=lambda item: item.received_at or datetime.min.replace(tzinfo=UTC),
    )
    context = "\n\n".join(
        f"""---
From: {message.sender}
To: {", ".join(message.recipients)}
Cc: {", ".join(message.cc)}
Subject: {message.subject}
Received: {message.received_at.isoformat() if message.received_at else ""}
Body:
{(message.body or message.snippet)[:5000]}"""
        for message in ordered
    )
    timezone = getattr(settings, "timezone", "Asia/Tokyo")
    return f"""あなたはメール本文からGoogleカレンダー登録候補を抽出するアシスタントです。
メールスレッドを読み、予定・面談・会議・締切ではなくカレンダーに入れるべき予定だけを抽出してください。

基準日時は対象メールの受信日時です:
{target_message.received_at.isoformat() if target_message.received_at else ""}

相対表現（例: 水曜日、明日、来週火曜）は、基準日時と文脈から具体的な日時へ解決してください。
時刻が明示されていて終了時刻がない場合は、通常の会議として1時間後を終了時刻にしてください。
タイムゾーンが不明な場合は {timezone} としてください。
予定がない場合は [] を返してください。

必ずJSON配列のみを返してください。説明文やMarkdownは不要です。
各要素は次のキーを持つJSONオブジェクトにしてください:
- title
- start_time（ISO 8601）
- end_time（ISO 8601）
- timezone
- location
- attendees（メールアドレス配列）
- description
- confidence（0から1）

メールスレッド:
{context}
"""


def _parse_calendar_candidates(raw: str) -> list[dict]:
    text = raw.strip()
    if text.startswith("```"):
        text = text.strip("`")
        if text.lower().startswith("json"):
            text = text[4:].strip()
    start = text.find("[")
    end = text.rfind("]")
    if start >= 0 and end >= start:
        text = text[start : end + 1]
    parsed = json.loads(text)
    if not isinstance(parsed, list):
        raise ValueError("calendar candidate response must be a JSON array")
    candidates = []
    for item in parsed:
        if not isinstance(item, dict):
            continue
        candidates.append(
            {
                "title": str(item.get("title") or ""),
                "start_time": str(item.get("start_time") or ""),
                "end_time": str(item.get("end_time") or ""),
                "timezone": str(item.get("timezone") or getattr(settings, "timezone", "Asia/Tokyo")),
                "location": str(item.get("location") or ""),
                "attendees": [
                    str(address)
                    for address in item.get("attendees", [])
                    if str(address).strip()
                ]
                if isinstance(item.get("attendees", []), list)
                else [],
                "description": str(item.get("description") or ""),
                "confidence": float(item.get("confidence") or 0.0),
            }
        )
    return candidates


def _parse_calendar_datetime(value: str):
    if not value:
        return None
    normalized = value.strip().replace("Z", "+00:00")
    try:
        return datetime.fromisoformat(normalized)
    except ValueError:
        return None


def _calendar_error_message(error: Exception) -> str:
    if isinstance(error, HttpError):
        detail = ""
        try:
            payload = json.loads(error.content.decode("utf-8"))
            detail = payload.get("error", {}).get("message", "")
        except Exception:
            detail = error.content.decode("utf-8", errors="replace")
        if detail:
            return f"カレンダー登録に失敗しました: {detail}"
    return f"カレンダー登録に失敗しました: {error}"


def _safe_return_url(value: str) -> str:
    stripped = value.strip()
    if not stripped:
        return ""
    if stripped.startswith("/"):
        return stripped
    if stripped.startswith("http://127.0.0.1") or stripped.startswith("http://localhost"):
        marker = stripped.find("/", stripped.find("//") + 2)
        return stripped[marker:] if marker >= 0 else "/"
    return ""


def _completed_thread_latest_messages(database) -> list[EmailMessage]:
    messages = database.list_saved_emails()
    resolved_ids = _resolved_message_ids(database)
    groups: dict[str, list[EmailMessage]] = {}
    for message in messages:
        key = message.gmail_thread_id or message.gmail_message_id
        groups.setdefault(key, []).append(message)
    completed = []
    for group in groups.values():
        latest = max(
            group,
            key=lambda message: (
                message.received_at.timestamp() if message.received_at else 0,
                message.gmail_message_id,
            ),
        )
        if "SENT" in set(latest.label_ids) or latest.gmail_message_id in resolved_ids:
            completed.append(latest)
    completed.sort(
        key=lambda message: message.received_at.timestamp() if message.received_at else 0,
        reverse=True,
    )
    return completed


def _unhandled_priority_messages(database) -> list[dict[str, object]]:
    generator = ReportGenerator(
        database=database,
        report_dir=settings.report_dir,
        timezone=getattr(settings, "timezone", "Asia/Tokyo"),
        excluded_sender_addresses=tuple(getattr(settings, "account_emails", ()) or ()),
    )
    emails = _active_list_emails(generator._list_report_emails())
    groups: dict[str, list[ReportEmail]] = {}
    for email in emails:
        key = email.gmail_thread_id or email.gmail_message_id
        groups.setdefault(key, []).append(email)

    rows: list[dict[str, object]] = []
    for group in groups.values():
        latest = max(
            group,
            key=lambda email: (
                email.received_at or datetime.min.replace(tzinfo=UTC),
                email.loaded_at,
            ),
        )
        if latest.analysis.priority not in {"high", "medium"}:
            continue
        rows.append(
            {
                "email": latest,
                "detail_href": f"/reports/{latest.report_date}/{latest.message_href}",
                "gmail_href": f"https://mail.google.com/mail/u/0/#all/{latest.gmail_thread_id or latest.gmail_message_id}",
            }
        )

    rows.sort(
        key=lambda row: (
            -PRIORITY_RANK.get(row["email"].analysis.priority, 2),  # type: ignore[union-attr]
            row["email"].received_at or datetime.max.replace(tzinfo=UTC),  # type: ignore[union-attr]
            row["email"].loaded_at,  # type: ignore[union-attr]
        )
    )
    return rows


def _resolved_message_ids(database) -> set[str]:
    with database.session_factory() as session:
        return set(
            session.scalars(
                select(ProcessingStateRecord.gmail_message_id).where(
                    ProcessingStateRecord.resolved.is_(True)
                )
            ).all()
        )


def _reply_defaults(message: EmailMessage, database) -> dict[str, str]:
    own_addresses = set(getattr(settings, "account_emails", ()) or ())
    own_addresses.update(database.list_account_emails())
    own_addresses = {address.lower() for address in own_addresses if address}

    to_addresses = _address_list(", ".join([*message.sender_candidates, message.sender]))
    if not to_addresses:
        to_addresses = _address_list(message.sender)

    cc_candidates = [*message.recipients, *message.cc]
    cc_addresses = [
        address
        for address in _dedupe_addresses(cc_candidates)
        if address.lower() not in own_addresses
        and address.lower() not in {item.lower() for item in to_addresses}
    ]
    return {
        "to": ", ".join(to_addresses),
        "cc": ", ".join(cc_addresses),
    }


def _reply_body_prefill(message: EmailMessage, draft: str = "") -> str:
    quoted = _quoted_original_message(message)
    draft = draft.strip()
    if draft and quoted:
        return f"{draft}\n\n\n{quoted}"
    if draft:
        return draft
    return quoted


def _quoted_original_message(message: EmailMessage) -> str:
    original = (message.body or message.snippet or "").strip()
    if not original:
        return ""
    return f"{_gmail_quote_header(message)}\n\n{original}"


def _gmail_quote_header(message: EmailMessage) -> str:
    timestamp = _gmail_quote_timestamp(message)
    sender = message.sender.strip()
    if timestamp and sender:
        return f"{timestamp} {sender}:"
    if timestamp:
        return f"{timestamp}:"
    return f"{sender}:" if sender else "元のメール:"


def _gmail_quote_timestamp(message: EmailMessage) -> str:
    if message.received_at is None:
        return ""
    try:
        timezone = ZoneInfo(str(getattr(settings, "timezone", "Asia/Tokyo")))
    except ZoneInfoNotFoundError:
        timezone = fixed_timezone(timedelta(hours=9))
    source_received_at = message.received_at
    if source_received_at.tzinfo is None:
        source_received_at = source_received_at.replace(tzinfo=UTC)
    received_at = source_received_at.astimezone(timezone)
    weekdays = ("月", "火", "水", "木", "金", "土", "日")
    weekday = weekdays[received_at.weekday()]
    return (
        f"{received_at.year}年{received_at.month}月{received_at.day}日"
        f"({weekday}) {received_at.hour:02d}:{received_at.minute:02d}"
    )


def _address_list(value: str) -> list[str]:
    return _dedupe_addresses(address for _, address in getaddresses([value]) if address)


def _dedupe_addresses(addresses) -> list[str]:
    seen: set[str] = set()
    deduped: list[str] = []
    for address in addresses:
        normalized = str(address).strip()
        key = normalized.lower()
        if not normalized or key in seen:
            continue
        seen.add(key)
        deduped.append(normalized)
    return deduped


def _form_value(form: dict[str, list[str]], key: str, default: str = "") -> str:
    values = form.get(key)
    if not values:
        return default
    return values[0]


def _instruction_form_snapshot(
    form: dict[str, list[str]],
) -> InstructionFormSnapshot:
    return InstructionFormSnapshot(
        name=_form_value(form, "name"),
        instruction=_form_value(form, "instruction"),
        from_query=_form_value(form, "from_query"),
        to_query=_form_value(form, "to_query"),
        subject_query=_form_value(form, "subject_query"),
        has_words=_form_value(form, "has_words"),
        doesnt_have=_form_value(form, "doesnt_have"),
        note=_form_value(form, "note"),
        enabled="enabled" in form,
    )


def _priority_rule_views(rules_database) -> list[PriorityRuleView]:
    items: list[PriorityRuleView] = []
    for rule in rules_database.list_filter_rules():
        items.append(
            PriorityRuleView(
                key=f"filter:{rule.id}",
                id=rule.id,
                name=rule.name or f"Rule #{rule.id}",
                enabled=rule.enabled,
                rule_type="固定Priority",
                action=rule.action,
                instruction="",
                priority_value=rule.preset_priority,
                from_query=rule.from_query,
                to_query=rule.to_query,
                subject_query=rule.subject_query,
                has_words=rule.has_words,
                doesnt_have=rule.doesnt_have,
                note=rule.note,
                edit_url=f"/rules/{rule.id}/edit",
                delete_url=f"/rules/{rule.id}/delete",
            )
        )
    for rule in rules_database.list_llm_instruction_rules():
        items.append(
            PriorityRuleView(
                key=f"instruction:{rule.id}",
                id=rule.id,
                name=rule.name or f"Instruction #{rule.id}",
                enabled=rule.enabled,
                rule_type="LLM追加指示",
                action="instruction",
                instruction=rule.instruction,
                priority_value="",
                from_query=rule.from_query,
                to_query=rule.to_query,
                subject_query=rule.subject_query,
                has_words=rule.has_words,
                doesnt_have=rule.doesnt_have,
                note=rule.note,
                edit_url=f"/instructions/{rule.id}/edit",
                delete_url=f"/instructions/{rule.id}/delete",
            )
        )
    items.sort(key=lambda item: (_priority_for_view(rules_database, item), item.key))
    return items


def _priority_for_view(rules_database, item: PriorityRuleView) -> int:
    if item.key.startswith("filter:"):
        rule = rules_database.get_filter_rule(item.id)
    else:
        rule = rules_database.get_llm_instruction_rule(item.id)
    return int(rule.priority) if rule is not None else 0
