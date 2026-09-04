#!/usr/bin/env python3
"""Leader control plane for three bounded source-management modules.

The modules are deterministic workflow boundaries, not autonomous AI agents.
The Leader owns scheduling, retries, state, human tasks and reports.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sqlite3
from collections import Counter
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable

import source_updater as updater
import staged_pipeline as pipeline
import human_operations
from system1.provider_extensions import collect_provider_extensions


MODULES = (
    "discovery_and_intake",
    "retrieval_and_monitoring",
    "governance_and_qa",
)
DEFAULT_INTERVAL_DAYS = {
    "discovery_and_intake": 30,
    "retrieval_and_monitoring": 0,
    "governance_and_qa": 0,
}
REVIEW_THRESHOLD_DAYS = 60


def selection_review_fingerprint(record: dict[str, Any]) -> str:
    fields = (
        "official_url", "retrieval_url", "issuer", "jurisdiction", "authoritative_language",
        "source_family", "document_type", "requirement_role", "inclusion_rationale",
        *updater.SCORE_FIELDS, "automation_readiness", "operator_selection_decision",
        "acquisition_channel", "provenance_status", "update_model", "applicability_reference",
    )
    payload = {field: str(record.get(field) or "").strip() for field in fields}
    return hashlib.sha256(json.dumps(payload, sort_keys=True).encode("utf-8")).hexdigest()[:16]


def selection_review_trigger(
    record: dict[str, Any],
    today: date,
    threshold_days: int,
    review_history: dict[tuple[str, str], dict[str, Any]],
) -> tuple[str, date | None, str, str, str, str] | None:
    if str(record.get("source_status") or "CURRENT").strip().upper() != "CURRENT":
        return None
    if "operator_selection_decision" not in record and not any(field in record for field in updater.SCORE_FIELDS):
        return None
    operator_decision = str(record.get("operator_selection_decision") or "PENDING").strip().upper()
    if operator_decision == "EXCLUDE":
        return None
    scores = {field: str(record.get(field) or "").strip().upper() for field in updater.SCORE_FIELDS}
    unresolved = [field for field, value in scores.items() if value != "HIGH"]
    if operator_decision == "INCLUDE" and not unresolved:
        return None
    fingerprint = selection_review_fingerprint(record)
    source_id = str(record.get("source_id") or "").strip()
    history = review_history.get((source_id, "SELECTION_PENDING"), {})
    completed_at = as_date(history.get("completed_at"))
    same_state = str(history.get("fingerprint") or "") == fingerprint
    if same_state and completed_at and (today - completed_at).days <= threshold_days:
        return None
    reasons: list[str] = []
    if operator_decision != "INCLUDE":
        reasons.append(f"operator_selection_decision is {operator_decision or 'missing'}")
    if unresolved:
        reasons.append("mandatory assessment is not HIGH: " + ", ".join(unresolved))
    since = completed_at if same_state and completed_at else (
        as_date(record.get("current_snapshot_date")) or as_date(record.get("last_success_at"))
    )
    return (
        "SELECTION_PENDING",
        since,
        "; ".join(reasons) + ".",
        "Review the human-governed source fields and choose INCLUDE, PENDING or EXCLUDE in Human Operation Desktop.",
        fingerprint,
        completed_at.isoformat() if same_state and completed_at else "",
    )


def completed_exit_code(overall_status: str) -> int:
    """Return success for any completed business outcome, including REVIEW."""
    return 0


def as_date(value: Any) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str) and value.strip():
        try:
            return datetime.fromisoformat(value.strip()).date()
        except ValueError:
            return None
    return None


def completed_manual_fallback(record: dict[str, Any]) -> bool:
    """Return whether a failed automated attempt has a later valid manual copy."""
    reviewed_at = as_date(record.get("manual_update_date"))
    attempted_at = as_date(record.get("last_attempt_at"))
    snapshot = str(record.get("snapshot_status") or "").strip().upper()
    origin = str(record.get("current_origin") or "").strip().upper()
    updated_by = str(record.get("manual_updated_by") or "").strip()
    return bool(
        reviewed_at
        and attempted_at
        and reviewed_at >= attempted_at
        and snapshot == "STORED"
        and origin == "MANUAL"
        and updated_by
    )


def completed_impact_review(record: dict[str, Any]) -> bool:
    """Return whether a rolling-content impact review was recorded after change."""
    reviewed_at = as_date(record.get("manual_update_date"))
    changed_at = as_date(record.get("last_success_at")) or as_date(record.get("last_attempt_at"))
    reviewed_by = str(record.get("manual_updated_by") or "").strip()
    return bool(reviewed_at and changed_at and reviewed_at >= changed_at and reviewed_by)


def review_trigger(record: dict[str, Any]) -> tuple[str, date | None, str, str] | None:
    source_status = str(record.get("source_status") or "").strip().upper()
    if source_status and source_status != "CURRENT":
        return None
    operator_selection = str(record.get("operator_selection_decision") or "").strip().upper()
    if operator_selection == "EXCLUDE" or (not operator_selection and updater.selection_from_scores(record) == "EXCLUDE"):
        return None

    download = str(record.get("download_status") or "").strip().upper()
    snapshot = str(record.get("snapshot_status") or "").strip().upper()
    change = str(record.get("content_change") or "").strip().upper()
    update_model = str(record.get("update_model") or "").strip().upper()
    provenance = str(record.get("provenance_status") or "").strip().upper()
    document_type = str(record.get("document_type") or "").strip()
    requirement_role = str(record.get("requirement_role") or "").strip()

    if snapshot == "MISSING":
        return (
            "MISSING_FILE",
            as_date(record.get("last_attempt_at")) or as_date(record.get("current_snapshot_date")),
            "The registered current file is missing from the expected folder.",
            "Restore or re-collect the file, then run the consistency audit.",
        )
    if download == "FAIL" and not completed_manual_fallback(record):
        reason = str(record.get("failure_reason") or "Automated retrieval failed.").strip()
        stage = str(record.get("failure_stage") or "UNKNOWN").strip().upper()
        action = {
            "HTTP": "Confirm the official URL or provide an authorised manual copy.",
            "CONNECTION": "Retry after checking network access and URL availability.",
            "CONTENT": "Inspect the response and choose an alternative official retrieval method or manual import.",
            "FORMAT": "Confirm the original file format and retrieval URL.",
            "STORAGE": "Resolve the workbook or file-storage conflict before retrying.",
        }.get(stage, "Inspect the failure reason and choose a safe recovery action.")
        return "DOWNLOAD_FAILURE", as_date(record.get("last_attempt_at")), reason, action
    if download == "PAYWALL_BLOCKED" and not completed_manual_fallback(record):
        return (
            "PAYWALL_BLOCKED",
            as_date(record.get("last_attempt_at")),
            str(record.get("failure_reason") or "The official file is restricted by a paywall or licence."),
            "Place an authorised copy in Data/00_Human_Intake or record an authorised retrieval route.",
        )
    if "official_url" in record and not str(record.get("official_url") or "").strip():
        return (
            "MISSING_OFFICIAL_URL",
            as_date(record.get("current_snapshot_date")),
            "The source has no official identity or publication URL.",
            "Confirm the official source identity in Human Operation Desktop.",
        )
    if "retrieval_url" in record and not str(record.get("retrieval_url") or "").strip():
        return (
            "MISSING_RETRIEVAL_URL",
            as_date(record.get("current_snapshot_date")),
            "The source has no exact retrieval target.",
            "Provide the complete official HTML target when available, otherwise the original official file URL.",
        )
    if update_model == "ROLLING" and change == "CHANGED" and not completed_impact_review(record):
        return (
            "ROLLING_CONTENT_CHANGED",
            as_date(record.get("last_success_at")),
            "Rolling source content changed after the previous validated snapshot.",
            "Review downstream Requirement impact before clearing the review item.",
        )
    if provenance == "UNVERIFIED":
        return (
            "PROVENANCE_UNVERIFIED",
            as_date(record.get("current_snapshot_date")),
            "The authenticity or authorisation of the current copy is not verified.",
            "Confirm provenance and authorised use, or retain the source as PENDING.",
        )
    if document_type == "Equipment manual / OEM instruction" or requirement_role == "Asset-specific normative":
        if not str(record.get("applicability_reference") or "").strip():
            return (
                "APPLICABILITY_UNCONFIRMED",
                as_date(record.get("current_snapshot_date")),
                "The equipment, model, site or asset applicability is not confirmed.",
                "Record applicability_reference before applying extracted Requirements.",
            )
    return None


def build_human_review_queue(
    records: list[dict[str, Any]],
    today: date | None = None,
    threshold_days: int = REVIEW_THRESHOLD_DAYS,
    review_history: dict[tuple[str, str], dict[str, Any]] | None = None,
) -> list[dict[str, Any]]:
    today = today or date.today()
    review_history = review_history or {}
    items: list[dict[str, Any]] = []
    official_url_counts = Counter(
        str(record.get("official_url") or "").strip()
        for record in records
        if str(record.get("official_url") or "").strip()
        and str(record.get("source_status") or "CURRENT").strip().upper() == "CURRENT"
        and str(record.get("operator_selection_decision") or "").strip().upper() != "EXCLUDE"
    )
    for record in records:
        triggers: list[tuple[str, date | None, str, str, str, str]] = []
        technical = review_trigger(record)
        if technical:
            triggers.append((*technical, "", ""))
        official_url = str(record.get("official_url") or "").strip()
        if official_url and official_url_counts[official_url] > 1 and str(record.get("operator_selection_decision") or "").strip().upper() != "EXCLUDE":
            triggers.append((
                "DUPLICATE_OFFICIAL_URL",
                as_date(record.get("current_snapshot_date")),
                f"The official URL is shared by {official_url_counts[official_url]} current source records.",
                "Confirm whether this record is a distinct source, correct its identity, or exclude it.",
                "",
                "",
            ))
        selection = selection_review_trigger(record, today, threshold_days, review_history)
        if selection:
            triggers.append(selection)
        for trigger_code, since, reason, action, fingerprint, cycle_key in triggers:
            days_open = max(0, (today - since).days) if since else None
            severity = "OVERDUE" if days_open is not None and days_open > threshold_days else "OPEN"
            source_id = str(record.get("source_id") or "").strip()
            stable = f"{source_id}|{trigger_code}|{since or 'undated'}|{fingerprint}|{cycle_key}"
            items.append({
                "review_id": hashlib.sha256(stable.encode("utf-8")).hexdigest()[:16],
                "source_id": source_id,
                "source_title": str(record.get("source_title") or "").strip(),
                "trigger": trigger_code,
                "review_since": since.isoformat() if since else None,
                "days_open": days_open,
                "severity": severity,
                "reason": reason,
                "recommended_action": action,
                "review_fingerprint": fingerprint,
            })
    return consolidate_review_queue(items)


def consolidate_review_queue(items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Create one current review task per source while preserving every issue code."""
    grouped: dict[str, list[dict[str, Any]]] = {}
    for item in items:
        grouped.setdefault(str(item.get("source_id") or ""), []).append(item)
    consolidated: list[dict[str, Any]] = []
    for source_id, group in grouped.items():
        group = sorted(group, key=lambda item: str(item.get("trigger") or ""))
        triggers: list[str] = []
        for item in group:
            codes = item.get("issue_codes") or [code.strip() for code in str(item.get("trigger") or "").split(";") if code.strip()]
            triggers.extend(str(code) for code in codes)
        triggers = list(dict.fromkeys(triggers))
        proposed_updates: dict[str, Any] = {}
        for item in group:
            proposed_updates.update(item.get("proposed_updates") or {})
        since_values = [as_date(item.get("review_since")) for item in group]
        since = min((value for value in since_values if value), default=None)
        days_values = [item.get("days_open") for item in group if item.get("days_open") is not None]
        fingerprint = hashlib.sha256(
            "|".join(str(item.get("review_fingerprint") or "") for item in group).encode("utf-8")
        ).hexdigest()[:16]
        stable = f"{source_id}|{'|'.join(triggers)}|{since or 'undated'}|{fingerprint}"
        consolidated.append({
            "review_id": hashlib.sha256(stable.encode("utf-8")).hexdigest()[:16],
            "source_id": source_id,
            "source_title": str(group[0].get("source_title") or ""),
            "trigger": "; ".join(triggers),
            "issue_codes": triggers,
            "review_since": since.isoformat() if since else None,
            "days_open": max(days_values) if days_values else None,
            "severity": "OVERDUE" if any(item.get("severity") == "OVERDUE" for item in group) else "OPEN",
            "reason": " | ".join(dict.fromkeys(str(item.get("reason") or "") for item in group)),
            "recommended_action": " | ".join(dict.fromkeys(str(item.get("recommended_action") or "") for item in group)),
            "review_fingerprint": fingerprint,
            "proposed_updates": proposed_updates,
        })
    return sorted(consolidated, key=lambda item: (item["days_open"] is None, -(item["days_open"] or 0), item["source_id"]))


def metadata_audit_review_items(
    audit_issues: list[str],
    records: list[dict[str, Any]],
    today: date,
    threshold_days: int,
) -> list[dict[str, Any]]:
    """Convert source-specific registry/file audit findings into human-visible tasks."""
    sources = {str(record.get("source_id") or "").strip(): record for record in records}
    items: list[dict[str, Any]] = []
    for issue in audit_issues:
        match = re.match(r"^Row\s+\d+\s+([A-Z]{2}\d{3}):\s*(.+)$", str(issue))
        if not match:
            continue
        source_id, reason = match.groups()
        source = sources.get(source_id, {})
        since = as_date(source.get("last_attempt_at")) or as_date(source.get("current_snapshot_date"))
        days_open = max(0, (today - since).days) if since else None
        stable = f"{source_id}|REGISTRY_FILE_MISMATCH|{reason}"
        items.append({
            "review_id": hashlib.sha256(stable.encode("utf-8")).hexdigest()[:16],
            "source_id": source_id,
            "source_title": str(source.get("source_title") or "").strip(),
            "trigger": "REGISTRY_FILE_MISMATCH",
            "review_since": since.isoformat() if since else None,
            "days_open": days_open,
            "severity": "OVERDUE" if days_open is not None and days_open > threshold_days else "OPEN",
            "reason": reason,
            "recommended_action": "Inspect the registered file and accept a controlled retry or provide an authorised manual replacement.",
            "review_fingerprint": "",
        })
    return items


def initialise_state(path: Path) -> sqlite3.Connection:
    path.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(path)
    connection.executescript(
        """
        CREATE TABLE IF NOT EXISTS leader_runs (
            run_id TEXT PRIMARY KEY,
            started_at TEXT NOT NULL,
            completed_at TEXT,
            overall_status TEXT NOT NULL,
            report_path TEXT
        );
        CREATE TABLE IF NOT EXISTS agent_runs (
            run_id TEXT NOT NULL,
            agent_name TEXT NOT NULL,
            status TEXT NOT NULL,
            completed_at TEXT NOT NULL,
            result_json TEXT NOT NULL,
            PRIMARY KEY (run_id, agent_name)
        );
        CREATE TABLE IF NOT EXISTS review_tasks (
            review_id TEXT PRIMARY KEY,
            source_id TEXT NOT NULL,
            trigger TEXT NOT NULL,
            severity TEXT NOT NULL,
            opened_at TEXT,
            last_seen_at TEXT NOT NULL,
            reason TEXT NOT NULL,
            recommended_action TEXT NOT NULL,
            status TEXT NOT NULL
        );
        """
    )
    connection.commit()
    return connection


def agent_is_due(
    connection: sqlite3.Connection,
    agent_name: str,
    now: datetime,
    force_all: bool,
    interval_days: dict[str, int] | None = None,
) -> bool:
    intervals = interval_days or DEFAULT_INTERVAL_DAYS
    if force_all or int(intervals.get(agent_name, DEFAULT_INTERVAL_DAYS[agent_name])) == 0:
        return True
    row = connection.execute(
        "SELECT MAX(completed_at) FROM agent_runs WHERE agent_name = ? AND status IN ('PASS','REVIEW','DRY_RUN')",
        (agent_name,),
    ).fetchone()
    if not row or not row[0]:
        return True
    try:
        completed = datetime.fromisoformat(row[0])
    except ValueError:
        return True
    return (now - completed).days >= int(intervals.get(agent_name, DEFAULT_INTERVAL_DAYS[agent_name]))


def persist_review_queue(connection: sqlite3.Connection, queue: list[dict[str, Any]], now: datetime) -> None:
    seen = {item["review_id"] for item in queue}
    for item in queue:
        connection.execute(
            """
            INSERT INTO review_tasks
              (review_id, source_id, trigger, severity, opened_at, last_seen_at, reason, recommended_action, status)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'OPEN')
            ON CONFLICT(review_id) DO UPDATE SET
              severity=excluded.severity,
              last_seen_at=excluded.last_seen_at,
              reason=excluded.reason,
              recommended_action=excluded.recommended_action,
              status='OPEN'
            """,
            (
                item["review_id"], item["source_id"], item["trigger"], item["severity"],
                item["review_since"], now.isoformat(timespec="seconds"), item["reason"], item["recommended_action"],
            ),
        )
    if seen:
        placeholders = ",".join("?" for _ in seen)
        connection.execute(
            f"UPDATE review_tasks SET status='CLOSED' WHERE status='OPEN' AND review_id NOT IN ({placeholders})",
            tuple(seen),
        )
    else:
        connection.execute("UPDATE review_tasks SET status='CLOSED' WHERE status='OPEN'")
    connection.commit()


def write_leader_report(config: dict[str, Any], report: dict[str, Any]) -> Path:
    report_dir: Path = config["log_root"] / "leader_reports"
    report_dir.mkdir(parents=True, exist_ok=True)
    path = report_dir / f"leader_{datetime.now():%Y%m%d_%H%M%S_%f}.json"
    path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def _module_status(results: list[dict[str, Any]]) -> str:
    statuses = {str(result.get("status") or "") for result in results}
    if "STOP" in statuses:
        return "STOP"
    if "REVIEW" in statuses:
        return "REVIEW"
    return "PASS"


def _discovery_module(records: list[dict[str, Any]]) -> dict[str, Any]:
    discovery = pipeline.discovery_stage(records)
    return {
        "status": discovery["status"],
        "purpose": "Coverage checks and controlled discovery intake",
        "discovery": discovery,
    }


def _retrieval_module(
    records: list[dict[str, Any]],
    config: dict[str, Any],
    include_pending: bool,
    available_only: bool,
    execute: bool,
) -> dict[str, Any]:
    monitoring = pipeline.monitoring_stage(records, config, include_pending)
    collection = pipeline.collection_stage(config, include_pending, available_only, execute)
    return {
        "status": _module_status([monitoring, collection]),
        "purpose": "Monitor, retrieve, validate and preserve source snapshots",
        "monitoring": monitoring,
        "collection": collection,
    }


def _governance_module(records: list[dict[str, Any]], config: dict[str, Any], config_path: Path) -> dict[str, Any]:
    assessment = pipeline.assessment_stage(records)
    metadata = pipeline.metadata_stage(config)
    monthly_qa = human_operations.generate_random_qa(config_path)
    qa_status = "REVIEW" if monthly_qa.get("status") == "STOP" else "PASS"
    return {
        "status": _module_status([assessment, metadata, {"status": qa_status}]),
        "purpose": "Selection governance, metadata audit and monthly quality assurance",
        "assessment": assessment,
        "metadata": metadata,
        "monthly_qa": monthly_qa,
    }


def run_leader(
    config_path: Path,
    execute: bool = False,
    include_pending: bool = False,
    available_only: bool = False,
    force_all: bool = False,
    providers: dict[str, object] | None = None,
) -> tuple[int, dict[str, Any]]:
    config = updater.read_config(config_path)
    human_cycle = human_operations.run_human_cycle(config_path)
    records = pipeline.read_records(config)
    provider_extensions = collect_provider_extensions(config_path, records, providers)
    now = datetime.now().astimezone()
    run_id = now.strftime("%Y%m%dT%H%M%S%f%z")
    state_path = config["log_root"] / "leader_state.sqlite"
    leader_config = config.get("leader", {})
    threshold_days = int(leader_config.get("review_threshold_days", REVIEW_THRESHOLD_DAYS))
    interval_days = leader_config.get("module_interval_days", DEFAULT_INTERVAL_DAYS)
    connection = initialise_state(state_path)
    connection.execute(
        "INSERT INTO leader_runs (run_id, started_at, overall_status) VALUES (?, ?, 'RUNNING')",
        (run_id, now.isoformat(timespec="seconds")),
    )
    connection.commit()

    module_results: list[dict[str, Any]] = []

    def wake(module_name: str, callable_stage) -> dict[str, Any]:
        if not agent_is_due(connection, module_name, now, force_all, interval_days):
            result = {"module": module_name, "status": "SLEEPING", "reason": "Not due under the current schedule"}
            module_results.append(result)
            return result
        stage_result = callable_stage()
        result = {"module": module_name, **stage_result}
        module_results.append(result)
        connection.execute(
            "INSERT OR REPLACE INTO agent_runs (run_id, agent_name, status, completed_at, result_json) VALUES (?, ?, ?, ?, ?)",
            (run_id, module_name, result["status"], datetime.now().astimezone().isoformat(timespec="seconds"), json.dumps(result, ensure_ascii=False, default=str)),
        )
        connection.commit()
        return result

    wake("discovery_and_intake", lambda: _discovery_module(records))
    wake("retrieval_and_monitoring", lambda: _retrieval_module(records, config, include_pending, available_only, execute))
    wake("governance_and_qa", lambda: _governance_module(records, config, config_path))

    # Re-read the discovery handoff after the specialist stage so findings
    # created during this Leader run reach the human gate immediately.
    post_agent_discovery_gate = human_operations.sync_discovery_candidates(config_path)
    post_agent_review_reconciliation = human_operations.reconcile_resolved_reviews(config_path)

    current_records = pipeline.read_records(config)
    review_history = human_operations.completed_review_history(config_path)
    review_queue = build_human_review_queue(current_records, now.date(), threshold_days, review_history)
    review_queue.extend(provider_extensions["review_items"])
    governance_result = next((result for result in module_results if result.get("module") == "governance_and_qa"), {})
    metadata_result = governance_result.get("metadata") or {}
    review_queue.extend(metadata_audit_review_items(metadata_result.get("audit_issues") or [], current_records, now.date(), threshold_days))
    review_queue = consolidate_review_queue(review_queue)
    desktop_sync = human_operations.sync_review_queue(config_path, review_queue)
    desktop_state = human_operations.open_operation_summary(config_path, now.date())
    persist_review_queue(connection, review_queue, now)
    overdue = [item for item in review_queue if item["severity"] == "OVERDUE"]
    oldest_days = max((item["days_open"] or 0 for item in review_queue), default=0)
    if provider_extensions["status"] == "DEGRADED":
        user_message = "The deterministic cycle completed, but an optional API provider extension failed. Review the provider error in the Leader report."
    elif desktop_state["oldest_days"] > threshold_days:
        user_message = f"Human review is overdue. {desktop_state['open_count']} open operation(s); the oldest has waited {desktop_state['oldest_days']} days."
    elif desktop_state["open_count"]:
        user_message = f"{desktop_state['open_count']} operation(s) require human attention in Human Operation Desktop."
    else:
        user_message = "No human review is currently required."

    actionable_statuses = {"REVIEW", "STOP"}
    overall = "REVIEW" if (
        desktop_state["open_count"]
        or provider_extensions["status"] == "DEGRADED"
        or any(result.get("status") in actionable_statuses for result in module_results)
    ) else "PASS"
    report = {
        "run_id": run_id,
        "generated_at": now.isoformat(timespec="seconds"),
        "architecture": "1 Leader + 3 bounded modules",
        "execute_collection": execute,
        "capability_mode": provider_extensions["mode"],
        "overall_status": overall,
        "user_message": user_message,
        "awakened_modules": [result["module"] for result in module_results if result["status"] != "SLEEPING"],
        "sleeping_modules": [result["module"] for result in module_results if result["status"] == "SLEEPING"],
        "module_results": module_results,
        "provider_extensions": provider_extensions,
        "human_operation_desktop": {
            "cycle": human_cycle,
            "post_agent_discovery_gate": post_agent_discovery_gate,
            "post_agent_review_reconciliation": post_agent_review_reconciliation,
            "review_queue_sync": desktop_sync,
            "open_operations": desktop_state,
        },
        "human_review": {
            "open_count": len(review_queue),
            "overdue_count": len(overdue),
            "oldest_days": oldest_days,
            "threshold_days": threshold_days,
            "items": review_queue,
        },
    }
    report_path = write_leader_report(config, report)
    connection.execute(
        "UPDATE leader_runs SET completed_at=?, overall_status=?, report_path=? WHERE run_id=?",
        (datetime.now().astimezone().isoformat(timespec="seconds"), overall, str(report_path), run_id),
    )
    connection.commit()
    connection.close()
    report["report_path"] = str(report_path)
    report["state_path"] = str(state_path)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    # REVIEW is a successful business outcome that asks for human attention,
    # not a process failure. Unhandled exceptions still produce a non-zero exit.
    return completed_exit_code(overall), report


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Run the Leader control plane for five requirement-source specialist agents.")
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--execute", action="store_true", help="Allow the collection agent to download and update files")
    parser.add_argument("--include-pending", action="store_true")
    parser.add_argument("--available-only", action="store_true", help="Skip records without retrieval_url instead of classifying their collection state")
    parser.add_argument("--all-sources", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--force-all", action="store_true", help="Wake all scheduled specialist agents regardless of interval")
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    code, _ = run_leader(
        Path(args.config).expanduser().resolve(),
        execute=args.execute,
        include_pending=args.include_pending,
        available_only=args.available_only and not args.all_sources,
        force_all=args.force_all,
    )
    return code


if __name__ == "__main__":
    raise SystemExit(main())
