from __future__ import annotations

import json
import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import patch
from zoneinfo import ZoneInfo

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from system1.doctor import run_doctor
from system1.scheduling import due_jobs, latest_due_key
from system1.workbook_guard import validate_windows_path
import leader_orchestrator as leader


class SchedulingTests(unittest.TestCase):
    def setUp(self):
        self.schedule = {
            "enabled": True,
            "timezone": "Europe/Oslo",
            "catch_up_enabled": True,
            "routine": {"days": ["MON", "TUE", "WED", "THU", "FRI"], "time": "09:00"},
            "full_check": {"days": ["SUN"], "time": "09:00"},
        }

    def test_catches_up_latest_missed_routine(self):
        now = datetime(2026, 9, 1, 8, 0, tzinfo=ZoneInfo("Europe/Oslo"))
        key = latest_due_key(now, self.schedule["routine"]["days"], "09:00", True)
        self.assertEqual(key[:10], "2026-08-31")

    def test_full_check_takes_precedence_when_both_are_due(self):
        now = datetime(2026, 9, 7, 10, 0, tzinfo=ZoneInfo("Europe/Oslo"))
        jobs = due_jobs(self.schedule, {"last_success": {}}, now)
        self.assertEqual(jobs[0][0], "full_check")
        self.assertEqual(len(jobs), 1)

    def test_completed_period_is_not_repeated(self):
        now = datetime(2026, 9, 8, 10, 0, tzinfo=ZoneInfo("Europe/Oslo"))
        first = due_jobs(self.schedule, {"last_success": {}}, now)
        key = first[0][1]
        state = {"last_success": {first[0][0]: key}}
        remaining = due_jobs(self.schedule, state, now)
        self.assertNotIn(first[0], remaining)


class PortabilityTests(unittest.TestCase):
    def test_windows_path_validation_accepts_spaces(self):
        issues = validate_windows_path(Path("C:/Team Files/System1/Requirement Sources"))
        self.assertEqual(issues, [])

    def test_windows_path_validation_rejects_reserved_name(self):
        issues = validate_windows_path(Path("System1/CON/source.pdf"))
        self.assertTrue(any("reserved" in issue.lower() for issue in issues))

    def test_doctor_accepts_relative_config_contract(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            code = root / "Code"
            config_dir = code / "config"
            data = root / "Data"
            runtime = code / "runtime"
            config_dir.mkdir(parents=True)
            data.mkdir()
            runtime.mkdir()
            workbook = root / "Requirement_Source_Registry.xlsx"
            workbook.write_bytes(b"test")
            config = config_dir / "config.json"
            config.write_text(json.dumps({
                "workbook": "../../Requirement_Source_Registry.xlsx",
                "source_root": "../../Data",
                "backup_root": "../runtime/backups",
                "log_root": "../runtime/logs",
                "timezone": "Europe/Oslo",
            }), encoding="utf-8")
            versions = {"openpyxl": "3.1.5", "portalocker": "4.3.0", "tzdata": "2026.3"}
            with patch("system1.doctor.importlib.metadata.version", side_effect=lambda name: versions[name]):
                result = run_doctor(config)
            failed_names = {check["name"] for check in result["checks"] if check["status"] == "FAIL"}
            self.assertNotIn("config", failed_names)
            self.assertNotIn("config.workbook", failed_names)


class ConsolidationTests(unittest.TestCase):
    def test_multiple_issues_become_one_source_task(self):
        items = [
            {"review_id": "1", "source_id": "PA001", "source_title": "Law", "trigger": "MISSING_RETRIEVAL_URL", "reason": "No URL", "recommended_action": "Add URL", "severity": "OPEN", "days_open": 2},
            {"review_id": "2", "source_id": "PA001", "source_title": "Law", "trigger": "SELECTION_PENDING", "reason": "Pending", "recommended_action": "Review", "severity": "OPEN", "days_open": 1},
        ]
        result = leader.consolidate_review_queue(items)
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["issue_codes"], ["MISSING_RETRIEVAL_URL", "SELECTION_PENDING"])


if __name__ == "__main__":
    unittest.main()

