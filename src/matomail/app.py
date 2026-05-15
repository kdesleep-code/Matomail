"""Application entry point for Matomail."""

from __future__ import annotations

import argparse

from googleapiclient.errors import HttpError

from .analyzer import EmailAnalyzer
from .config import Settings
from .database import Database, RulesDatabase
from .database import FILTER_ACTION_ALWAYS_PROCESS, FILTER_ACTION_PRECLASSIFY
from .database import FILTER_ACTION_SKIP_ANALYSIS
from .gmail_client import GmailClient
from .llm_client import LLMClient


def main() -> int:
    parser = argparse.ArgumentParser(prog="matomail")
    subparsers = parser.add_subparsers(dest="command")
    subparsers.add_parser("fetch", help="Fetch recent Gmail messages and print a summary.")
    analyze_parser = subparsers.add_parser(
        "analyze",
        help="Analyze saved emails with the configured LLM.",
    )
    analyze_parser.add_argument(
        "--limit",
        type=int,
        default=1,
        help="Maximum number of saved emails to analyze. Defaults to 1.",
    )
    rules_parser = subparsers.add_parser("rules", help="Manage filter rules.")
    rules_subparsers = rules_parser.add_subparsers(dest="rules_command")

    rules_subparsers.add_parser("list", help="List filter rules.")

    add_sender_skip_parser = rules_subparsers.add_parser(
        "add-sender-skip",
        help="Skip analysis for messages from a sender address.",
    )
    add_sender_skip_parser.add_argument("email_address")
    add_sender_skip_parser.add_argument("--name", default="")
    add_sender_skip_parser.add_argument("--note", default="")

    add_subject_skip_parser = rules_subparsers.add_parser(
        "add-subject-skip",
        help="Skip analysis for messages whose subject contains the given text.",
    )
    add_subject_skip_parser.add_argument("subject")
    add_subject_skip_parser.add_argument("--name", default="")
    add_subject_skip_parser.add_argument("--note", default="")

    add_subject_preclassify_parser = rules_subparsers.add_parser(
        "add-subject-preclassify",
        help="Create a no-LLM preclassification rule for matching subjects.",
    )
    add_subject_preclassify_parser.add_argument("subject")
    add_subject_preclassify_parser.add_argument(
        "--priority",
        choices=["high", "medium", "low"],
        required=True,
    )
    add_subject_preclassify_parser.add_argument("--category", default="preclassified")
    add_subject_preclassify_parser.add_argument("--summary", default="")
    add_subject_preclassify_parser.add_argument("--suggested-action", default="")
    add_subject_preclassify_parser.add_argument("--requires-reply", action="store_true")
    add_subject_preclassify_parser.add_argument("--name", default="")
    add_subject_preclassify_parser.add_argument("--note", default="")

    for command_name in ["enable", "disable", "delete"]:
        command_parser = rules_subparsers.add_parser(command_name)
        command_parser.add_argument("rule_id", type=int)

    instructions_parser = subparsers.add_parser(
        "instructions",
        help="Manage conditional LLM instruction rules.",
    )
    instructions_subparsers = instructions_parser.add_subparsers(
        dest="instructions_command"
    )
    instructions_subparsers.add_parser("list", help="List LLM instruction rules.")
    add_instruction_parser = instructions_subparsers.add_parser(
        "add",
        help="Add a conditional LLM instruction rule.",
    )
    add_instruction_parser.add_argument("instruction")
    add_instruction_parser.add_argument("--name", default="")
    add_instruction_parser.add_argument("--priority", type=int, default=0)
    add_instruction_parser.add_argument("--from-query", default="")
    add_instruction_parser.add_argument("--to-query", default="")
    add_instruction_parser.add_argument("--subject-query", default="")
    add_instruction_parser.add_argument("--has-words", default="")
    add_instruction_parser.add_argument("--doesnt-have", default="")
    add_instruction_parser.add_argument("--note", default="")

    for command_name in ["enable", "disable", "delete"]:
        command_parser = instructions_subparsers.add_parser(command_name)
        command_parser.add_argument("rule_id", type=int)
    args = parser.parse_args()

    settings = Settings()

    if args.command == "fetch":
        database = Database(
            settings.db_path,
            max_size_bytes=int(settings.db_max_size_mb * 1024 * 1024),
            backup_dir=settings.db_backup_dir,
            store_email_body=settings.store_email_body,
        )
        rules_database = RulesDatabase(settings.rules_db_path)
        database.create_all()
        rules_database.create_all()
        client = GmailClient.from_oauth(settings)
        try:
            messages = client.fetch_recent_messages(
                lookback_days=settings.lookback_days,
                max_results=settings.max_emails_per_run,
            )
        except HttpError as error:
            print(f"Gmail API request failed: {error.reason}")
            print("If Gmail API is disabled, enable it in Google Cloud Console and retry.")
            return 1
        database.save_emails(messages)
        processable_messages = []
        for message in database.filter_processable(messages):
            decision = rules_database.get_filter_decision(message)
            if decision is None:
                processable_messages.append(message)
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
            elif decision["action"] == FILTER_ACTION_ALWAYS_PROCESS:
                processable_messages.append(message)

        messages = processable_messages
        for index, message in enumerate(messages, start=1):
            attachment_note = " attachments" if message.has_attachments else ""
            print(f"{index}. [{message.received_at}] {message.subject}{attachment_note}")
        print(f"Fetched {len(messages)} processable message(s).")
        return 0

    if args.command == "analyze":
        if args.limit < 1:
            print("--limit must be at least 1.")
            return 1

        database = Database(
            settings.db_path,
            max_size_bytes=int(settings.db_max_size_mb * 1024 * 1024),
            backup_dir=settings.db_backup_dir,
            store_email_body=settings.store_email_body,
        )
        rules_database = RulesDatabase(settings.rules_db_path)
        database.create_all()
        rules_database.create_all()

        messages = database.list_unanalyzed_emails(limit=args.limit)
        messages = rules_database.filter_processable(messages)
        if not messages:
            print("No saved emails need analysis.")
            return 0

        analyzer = EmailAnalyzer(llm_client=LLMClient.from_settings(settings))
        for index, message in enumerate(messages, start=1):
            analysis = analyzer.analyze_and_save(
                message,
                mail_database=database,
                rules_database=rules_database,
                llm_model=settings.llm_model,
            )
            print(
                f"{index}. {message.subject} -> "
                f"{analysis['priority']} / {analysis['category']}"
            )
        print(f"Analyzed {len(messages)} message(s).")
        return 0

    if args.command == "rules":
        rules_database = RulesDatabase(settings.rules_db_path)
        rules_database.create_all()
        return _handle_rules_command(args, rules_database)

    if args.command == "instructions":
        rules_database = RulesDatabase(settings.rules_db_path)
        rules_database.create_all()
        return _handle_instructions_command(args, rules_database)

    print(f"Matomail is configured to write reports to {settings.report_dir}")
    return 0


def _handle_rules_command(args: argparse.Namespace, rules_database: RulesDatabase) -> int:
    if args.rules_command == "list":
        rules = rules_database.list_filter_rules()
        if not rules:
            print("No filter rules.")
            return 0
        for rule in rules:
            status = "enabled" if rule.enabled else "disabled"
            condition = _format_rule_condition(rule)
            print(
                f"{rule.id}. [{status}] {rule.action} "
                f"name={rule.name!r} condition={condition}"
            )
        return 0

    if args.rules_command == "add-sender-skip":
        rule = rules_database.add_sender_filter(
            args.email_address,
            name=args.name,
            note=args.note,
            action=FILTER_ACTION_SKIP_ANALYSIS,
        )
        print(f"Created sender skip rule id={rule.id}")
        return 0

    if args.rules_command == "add-subject-skip":
        rule = rules_database.add_filter_rule(
            action=FILTER_ACTION_SKIP_ANALYSIS,
            name=args.name,
            subject_query=args.subject,
            note=args.note,
        )
        print(f"Created subject skip rule id={rule.id}")
        return 0

    if args.rules_command == "add-subject-preclassify":
        rule = rules_database.add_filter_rule(
            action=FILTER_ACTION_PRECLASSIFY,
            name=args.name,
            subject_query=args.subject,
            preset_priority=args.priority,
            preset_category=args.category,
            preset_summary_ja=args.summary,
            preset_suggested_action_ja=args.suggested_action,
            preset_requires_reply=args.requires_reply,
            note=args.note,
        )
        print(f"Created subject preclassify rule id={rule.id}")
        return 0

    if args.rules_command == "enable":
        return _print_update_result(
            rules_database.set_filter_rule_enabled(args.rule_id, True),
            "Enabled",
            "filter rule",
            args.rule_id,
        )

    if args.rules_command == "disable":
        return _print_update_result(
            rules_database.set_filter_rule_enabled(args.rule_id, False),
            "Disabled",
            "filter rule",
            args.rule_id,
        )

    if args.rules_command == "delete":
        return _print_update_result(
            rules_database.delete_filter_rule(args.rule_id),
            "Deleted",
            "filter rule",
            args.rule_id,
        )

    print("Missing rules subcommand.")
    return 1


def _handle_instructions_command(
    args: argparse.Namespace,
    rules_database: RulesDatabase,
) -> int:
    if args.instructions_command == "list":
        rules = rules_database.list_llm_instruction_rules()
        if not rules:
            print("No LLM instruction rules.")
            return 0
        for rule in rules:
            status = "enabled" if rule.enabled else "disabled"
            condition = _format_rule_condition(rule)
            print(
                f"{rule.id}. [{status}] priority={rule.priority} "
                f"name={rule.name!r} condition={condition} "
                f"instruction={rule.instruction!r}"
            )
        return 0

    if args.instructions_command == "add":
        rule = rules_database.add_llm_instruction_rule(
            instruction=args.instruction,
            name=args.name,
            priority=args.priority,
            from_query=args.from_query,
            to_query=args.to_query,
            subject_query=args.subject_query,
            has_words=args.has_words,
            doesnt_have=args.doesnt_have,
            note=args.note,
        )
        print(f"Created LLM instruction rule id={rule.id}")
        return 0

    if args.instructions_command == "enable":
        return _print_update_result(
            rules_database.set_llm_instruction_rule_enabled(args.rule_id, True),
            "Enabled",
            "LLM instruction rule",
            args.rule_id,
        )

    if args.instructions_command == "disable":
        return _print_update_result(
            rules_database.set_llm_instruction_rule_enabled(args.rule_id, False),
            "Disabled",
            "LLM instruction rule",
            args.rule_id,
        )

    if args.instructions_command == "delete":
        return _print_update_result(
            rules_database.delete_llm_instruction_rule(args.rule_id),
            "Deleted",
            "LLM instruction rule",
            args.rule_id,
        )

    print("Missing instructions subcommand.")
    return 1


def _format_rule_condition(rule: object) -> str:
    parts = []
    for field_name in [
        "from_query",
        "to_query",
        "subject_query",
        "has_words",
        "doesnt_have",
    ]:
        value = getattr(rule, field_name, "")
        if value:
            parts.append(f"{field_name}={value!r}")
    return ", ".join(parts) if parts else "always"


def _print_update_result(
    ok: bool,
    verb: str,
    item_name: str,
    item_id: int,
) -> int:
    if not ok:
        print(f"No {item_name} found with id={item_id}")
        return 1
    print(f"{verb} {item_name} id={item_id}")
    return 0




if __name__ == "__main__":
    raise SystemExit(main())
