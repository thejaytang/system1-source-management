"""Portable due-time logic with OS-specific registration adapters."""

from __future__ import annotations

import json
import os
import platform
import plistlib
import subprocess
import sys
import tempfile
from datetime import datetime, time, timedelta
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo


TASK_NAME = "System1 Source Management"
PLIST_LABEL = "com.smartercompliance.system1-source-management"
WEEKDAYS = {"MON": 0, "TUE": 1, "WED": 2, "THU": 3, "FRI": 4, "SAT": 5, "SUN": 6}


def read_schedule(path: Path) -> dict[str, Any]:
    data = json.loads(path.read_text(encoding="utf-8"))
    data.setdefault("enabled", False)
    data.setdefault("timezone", "Europe/Oslo")
    data.setdefault("catch_up_enabled", True)
    data.setdefault("poll_interval_minutes", 30)
    data.setdefault("notifications", True)
    data.setdefault("routine", {"days": ["MON", "TUE", "WED", "THU", "FRI"], "time": "09:00"})
    data.setdefault("full_check", {"days": ["SUN"], "time": "09:00"})
    validate_schedule(data)
    return data


def validate_schedule(data: dict[str, Any]) -> None:
    ZoneInfo(str(data["timezone"]))
    for name in ("routine", "full_check"):
        section = data.get(name) or {}
        days = section.get("days") or []
        if not days or any(str(day).upper() not in WEEKDAYS for day in days):
            raise ValueError(f"{name}.days must contain valid three-letter weekdays.")
        datetime.strptime(str(section.get("time") or ""), "%H:%M")
    interval = int(data.get("poll_interval_minutes", 30))
    if interval < 5:
        raise ValueError("poll_interval_minutes must be at least 5.")


def latest_due_key(now: datetime, days: list[str], clock: str, catch_up: bool) -> str | None:
    hour, minute = map(int, clock.split(":"))
    candidates: list[datetime] = []
    horizon = 8 if catch_up else 1
    allowed = {WEEKDAYS[day.upper()] for day in days}
    for offset in range(horizon):
        day = (now - timedelta(days=offset)).date()
        if day.weekday() not in allowed:
            continue
        candidate = datetime.combine(day, time(hour, minute), tzinfo=now.tzinfo)
        if candidate <= now:
            candidates.append(candidate)
    if not candidates:
        return None
    return max(candidates).isoformat(timespec="minutes")


def due_jobs(schedule: dict[str, Any], state: dict[str, Any], now: datetime | None = None) -> list[tuple[str, str]]:
    if not schedule.get("enabled", False):
        return []
    now = now or datetime.now(ZoneInfo(schedule["timezone"]))
    catch_up = bool(schedule.get("catch_up_enabled", True))
    jobs: list[tuple[str, str]] = []
    for name in ("routine", "full_check"):
        section = schedule[name]
        key = latest_due_key(now, list(section["days"]), str(section["time"]), catch_up)
        if key and state.get("last_success", {}).get(name) != key:
            jobs.append((name, key))
    jobs.sort(key=lambda item: 0 if item[0] == "full_check" else 1)
    if jobs and jobs[0][0] == "full_check":
        return [jobs[0]]
    return jobs


def load_state(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {"last_success": {}, "last_attempt": {}}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {"last_success": {}, "last_attempt": {}}


def save_state(path: Path, state: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, name = tempfile.mkstemp(prefix="scheduler-state-", suffix=".json", dir=path.parent)
    os.close(fd)
    temporary = Path(name)
    try:
        temporary.write_text(json.dumps(state, indent=2, sort_keys=True), encoding="utf-8")
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def register_schedule(code_root: Path, schedule_path: Path) -> str:
    schedule = read_schedule(schedule_path)
    if not schedule.get("enabled", False):
        raise RuntimeError("Automation is disabled in schedule.json. Confirm the schedule and set enabled to true before registration.")
    python = (code_root / ".venv" / ("Scripts/python.exe" if os.name == "nt" else "bin/python")).resolve()
    if not python.exists():
        raise RuntimeError("The project environment is missing. Run the platform setup script first.")
    interval = int(schedule["poll_interval_minutes"]) * 60
    system = platform.system()
    if system == "Darwin":
        destination = Path.home() / "Library/LaunchAgents" / f"{PLIST_LABEL}.plist"
        destination.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "Label": PLIST_LABEL,
            "ProgramArguments": [str(python), "-m", "system1", "scheduled", "--config", str((code_root / "config/config.json").resolve()), "--schedule", str(schedule_path.resolve())],
            "WorkingDirectory": str(code_root.resolve()),
            "EnvironmentVariables": {"PYTHONPATH": str((code_root / "src").resolve())},
            "StartInterval": interval,
            "RunAtLoad": True,
            "StandardOutPath": str((code_root / "runtime/logs/scheduler.stdout.log").resolve()),
            "StandardErrorPath": str((code_root / "runtime/logs/scheduler.stderr.log").resolve()),
        }
        destination.write_bytes(plistlib.dumps(payload))
        subprocess.run(["launchctl", "unload", str(destination)], check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        subprocess.run(["launchctl", "load", str(destination)], check=True)
        return str(destination)
    if system == "Windows":
        command = f'"{python}" -m system1 scheduled --config "{(code_root / "config/config.json").resolve()}" --schedule "{schedule_path.resolve()}"'
        subprocess.run(["schtasks", "/Create", "/F", "/TN", TASK_NAME, "/SC", "MINUTE", "/MO", str(max(1, interval // 60)), "/TR", command], check=True)
        return TASK_NAME
    raise RuntimeError("Formal schedule registration is supported only on macOS and Windows.")


def unregister_schedule() -> str:
    system = platform.system()
    if system == "Darwin":
        destination = Path.home() / "Library/LaunchAgents" / f"{PLIST_LABEL}.plist"
        if destination.exists():
            subprocess.run(["launchctl", "unload", str(destination)], check=False)
            destination.unlink()
        return str(destination)
    if system == "Windows":
        subprocess.run(["schtasks", "/Delete", "/F", "/TN", TASK_NAME], check=False)
        return TASK_NAME
    raise RuntimeError("Formal schedule registration is supported only on macOS and Windows.")
