from __future__ import annotations

import argparse
import json
from datetime import datetime
from pathlib import Path
from typing import Iterable
from zoneinfo import ZoneInfo

import leader_orchestrator
import source_updater

from .doctor import run_doctor
from .platform.notifications import notify
from .scheduling import due_jobs, load_state, read_schedule, register_schedule, save_state, unregister_schedule
from .workbook_guard import WorkbookBusyError, excel_appears_open, exclusive_process_lock


def code_root() -> Path:
    return Path(__file__).resolve().parents[2]


def default_config() -> Path:
    return code_root() / "config/config.json"


def default_schedule() -> Path:
    return code_root() / "config/schedule.json"


def latest_report(config_path: Path) -> dict:
    config = source_updater.read_config(config_path)
    reports = sorted((config["log_root"] / "leader_reports").glob("leader_*.json"), reverse=True)
    if not reports:
        return {"status": "NOT_RUN", "message": "No Leader report exists yet."}
    return json.loads(reports[0].read_text(encoding="utf-8"))


def run_cycle(config_path: Path, full: bool) -> tuple[int, dict]:
    config = source_updater.read_config(config_path)
    if excel_appears_open(config["workbook"]):
        return 2, {"status": "DEFERRED_WORKBOOK_OPEN", "message": "Save and close the workbook before running System1."}
    try:
        with exclusive_process_lock(config["log_root"] / ".system1-run.lock", float(config.get("lock_stale_hours", 12))):
            code, report = leader_orchestrator.run_leader(
                config_path,
                execute=full,
                include_pending=True,
                available_only=False,
                force_all=full,
            )
    except (source_updater.UpdaterError, WorkbookBusyError) as exc:
        lowered = str(exc).lower()
        status = "DEFERRED_ALREADY_RUNNING" if "already running" in lowered else ("DEFERRED_WORKBOOK_OPEN" if "open" in lowered else "FAIL")
        return 2, {"status": status, "message": str(exc)}
    return code, {"status": report["overall_status"], "message": report["user_message"], "report_path": report.get("report_path")}


def run_scheduled(config_path: Path, schedule_path: Path) -> tuple[int, dict]:
    schedule = read_schedule(schedule_path)
    config = source_updater.read_config(config_path)
    state_path = config["log_root"] / "scheduler_state.json"
    state = load_state(state_path)
    now = datetime.now(ZoneInfo(schedule["timezone"]))
    jobs = due_jobs(schedule, state, now)
    if not jobs:
        return 0, {"status": "NOT_DUE", "message": "No scheduled job is due."}
    job, key = jobs[0]
    state.setdefault("last_attempt", {})[job] = now.isoformat(timespec="seconds")
    code, result = run_cycle(config_path, full=job == "full_check")
    result.update({"job": job, "due_key": key})
    if code == 0:
        state.setdefault("last_success", {})[job] = key
        if job == "full_check":
            routine = schedule["routine"]
            routine_jobs = due_jobs(schedule, state, now)
            for pending, pending_key in routine_jobs:
                if pending == "routine":
                    state["last_success"]["routine"] = pending_key
    state["last_result"] = result
    save_state(state_path, state)
    if schedule.get("notifications", True) and result["status"] not in {"NOT_DUE"}:
        notify("System1 Source Management", result["message"])
    return code, result


def interactive_menu(config_path: Path, schedule_path: Path) -> int:
    print("System1 Source Management")
    print("1. Routine Cycle")
    print("2. Full Source Check")
    print("3. Status Only")
    print("4. Validate Environment")
    print("5. View Last Run")
    choice = input("Choose 1-5: ").strip()
    if choice == "1":
        code, result = run_cycle(config_path, False)
    elif choice == "2":
        code, result = run_cycle(config_path, True)
    elif choice in {"3", "5"}:
        code, result = 0, latest_report(config_path)
    elif choice == "4":
        result = run_doctor(config_path, schedule_path)
        code = 0 if result["status"] == "PASS" else 1
    else:
        print("Invalid choice.")
        return 2
    print(json.dumps(result, indent=2, default=str))
    return code


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="System1 cross-platform source-management runner")
    parser.add_argument("command", nargs="?", choices=["routine", "full", "status", "doctor", "scheduled", "register-schedule", "unregister-schedule", "menu"], default="menu")
    parser.add_argument("--config", type=Path, default=default_config())
    parser.add_argument("--schedule", type=Path, default=default_schedule())
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    config_path = args.config.expanduser().resolve()
    schedule_path = args.schedule.expanduser().resolve()
    if args.command == "menu":
        return interactive_menu(config_path, schedule_path)
    if args.command == "routine":
        code, result = run_cycle(config_path, False)
    elif args.command == "full":
        code, result = run_cycle(config_path, True)
    elif args.command == "status":
        code, result = 0, latest_report(config_path)
    elif args.command == "doctor":
        result = run_doctor(config_path, schedule_path)
        code = 0 if result["status"] == "PASS" else 1
    elif args.command == "scheduled":
        code, result = run_scheduled(config_path, schedule_path)
    elif args.command == "register-schedule":
        code, result = 0, {"status": "REGISTERED", "target": register_schedule(code_root(), schedule_path)}
    else:
        code, result = 0, {"status": "UNREGISTERED", "target": unregister_schedule()}
    print(json.dumps(result, indent=2, default=str))
    return code
