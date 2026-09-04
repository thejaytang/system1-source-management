"""Environment diagnostics for user-configured macOS and Windows installs."""

from __future__ import annotations

import importlib.metadata
import json
import os
import platform
import subprocess
import sys
from dataclasses import dataclass, asdict
from pathlib import Path
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from .workbook_guard import excel_appears_open, validate_windows_path


@dataclass
class Check:
    name: str
    status: str
    detail: str


def _dependency_check(package: str, expected: str | None = None) -> Check:
    try:
        version = importlib.metadata.version(package)
    except importlib.metadata.PackageNotFoundError:
        return Check(package, "FAIL", "Not installed")
    if expected and version != expected:
        return Check(package, "WARN", f"Installed {version}; expected {expected}")
    return Check(package, "PASS", version)


def run_doctor(config_path: Path, schedule_path: Path | None = None) -> dict:
    checks: list[Check] = []
    checks.append(Check("python", "PASS" if sys.version_info >= (3, 11) else "FAIL", platform.python_version()))
    checks.extend([
        _dependency_check("openpyxl", "3.1.5"),
        _dependency_check("portalocker", "4.3.0"),
        _dependency_check("tzdata", "2026.3"),
    ])
    if not config_path.exists():
        checks.append(Check("config", "FAIL", f"Missing {config_path}"))
        return _result(checks)
    try:
        raw = json.loads(config_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        checks.append(Check("config", "FAIL", str(exc)))
        return _result(checks)
    checks.append(Check("config", "PASS", str(config_path)))
    for key in ("workbook", "source_root", "backup_root", "log_root"):
        value = str(raw.get(key) or "")
        if not value:
            checks.append(Check(f"config.{key}", "FAIL", "Missing value"))
            continue
        candidate = Path(value)
        status = "PASS" if not candidate.is_absolute() else "FAIL"
        checks.append(Check(f"config.{key}", status, "Relative path" if status == "PASS" else "Absolute paths are not portable"))
    base = config_path.parent.resolve()
    workbook = (base / str(raw.get("workbook") or "")).resolve()
    source_root = (base / str(raw.get("source_root") or "")).resolve()
    checks.append(Check("workbook", "PASS" if workbook.is_file() else "FAIL", str(workbook)))
    checks.append(Check("excel_lock", "FAIL" if excel_appears_open(workbook) else "PASS", "Workbook is open" if excel_appears_open(workbook) else "Workbook is available"))
    checks.append(Check("data_directory", "PASS" if source_root.is_dir() else "FAIL", str(source_root)))
    if source_root.is_dir():
        portable_issues: list[str] = []
        casefolded: dict[str, str] = {}
        root = config_path.parents[2]
        for item in source_root.rglob("*"):
            relative = item.relative_to(root)
            portable_issues.extend(validate_windows_path(relative))
            key = str(relative).casefold()
            if key in casefolded and casefolded[key] != str(relative):
                portable_issues.append(f"Case-insensitive path collision: {casefolded[key]} and {relative}")
            casefolded[key] = str(relative)
        checks.append(Check("data_windows_compatibility", "PASS" if not portable_issues else "FAIL", "; ".join(portable_issues[:10]) or "Portable filenames and paths"))
    for key in ("backup_root", "log_root"):
        path = (base / str(raw.get(key) or "")).resolve()
        parent = path if path.exists() else path.parent
        checks.append(Check(f"{key}_writable", "PASS" if parent.exists() and os.access(parent, os.W_OK) else "FAIL", str(path)))
    timezone = str(raw.get("timezone") or "Europe/Oslo")
    try:
        ZoneInfo(timezone)
        checks.append(Check("timezone", "PASS", timezone))
    except ZoneInfoNotFoundError:
        checks.append(Check("timezone", "FAIL", timezone))
    root_issues = validate_windows_path(config_path.parents[2])
    checks.append(Check("windows_path_compatibility", "PASS" if not root_issues else "FAIL", "; ".join(root_issues) or "Portable"))
    if schedule_path:
        if schedule_path.exists():
            try:
                schedule = json.loads(schedule_path.read_text(encoding="utf-8"))
                checks.append(Check("schedule", "PASS", f"enabled={bool(schedule.get('enabled', False))}"))
                if bool(schedule.get("enabled", False)):
                    if platform.system() == "Darwin":
                        target = Path.home() / "Library/LaunchAgents/com.smartercompliance.system1-source-management.plist"
                        registered = target.is_file()
                    elif platform.system() == "Windows":
                        result = subprocess.run(["schtasks", "/Query", "/TN", "System1 Source Management"], capture_output=True, text=True, check=False)
                        registered = result.returncode == 0
                        target = Path("Windows Task Scheduler: System1 Source Management")
                    else:
                        registered = False
                        target = Path("Unsupported scheduler platform")
                    checks.append(Check("scheduler_registration", "PASS" if registered else "WARN", str(target)))
                else:
                    checks.append(Check("scheduler_registration", "PASS", "Automation is intentionally disabled"))
            except (OSError, json.JSONDecodeError) as exc:
                checks.append(Check("schedule", "FAIL", str(exc)))
        else:
            checks.append(Check("schedule", "WARN", "No active schedule.json; manual operation remains available"))
    return _result(checks)


def _result(checks: list[Check]) -> dict:
    return {
        "status": "FAIL" if any(check.status == "FAIL" for check in checks) else "PASS",
        "checks": [asdict(check) for check in checks],
    }
