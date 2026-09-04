from __future__ import annotations

import sys
import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
import leader_orchestrator as leader


class LeaderOrchestratorTests(unittest.TestCase):
    def test_review_queue_prioritises_overdue_download_failure(self):
        records = [
            {
                "source_id": "PA010",
                "source_title": "Animal Health Law",
                "download_status": "FAIL",
                "failure_stage": "CONTENT",
                "failure_reason": "HTML content is too small: 0 bytes.",
                "last_attempt_at": datetime(2026, 5, 1, 10, 0),
                "snapshot_status": "NOT_COLLECTED",
            },
            {
                "source_id": "PA011",
                "source_title": "Rolling guidance",
                "download_status": "SUCCESS",
                "content_change": "CHANGED",
                "update_model": "ROLLING",
                "last_success_at": datetime(2026, 8, 20, 10, 0),
                "snapshot_status": "STORED",
            },
        ]
        queue = leader.build_human_review_queue(records, date(2026, 8, 28))
        self.assertEqual([item["source_id"] for item in queue], ["PA010", "PA011"])
        self.assertEqual(queue[0]["severity"], "OVERDUE")
        self.assertEqual(queue[1]["trigger"], "ROLLING_CONTENT_CHANGED")

    def test_equipment_without_applicability_is_sent_to_human_review(self):
        records = [{
            "source_id": "MS001",
            "source_title": "OEM manual",
            "document_type": "Equipment manual / OEM instruction",
            "requirement_role": "Asset-specific normative",
            "applicability_reference": "",
            "download_status": "SUCCESS",
            "snapshot_status": "STORED",
        }]
        queue = leader.build_human_review_queue(records, date(2026, 8, 28))
        self.assertEqual(queue[0]["trigger"], "APPLICABILITY_UNCONFIRMED")
        self.assertIn("applicability_reference", queue[0]["recommended_action"])

    def test_clean_record_does_not_create_review_task(self):
        records = [{
            "source_id": "PA001",
            "source_title": "Stable source",
            "download_status": "SUCCESS",
            "content_change": "UNCHANGED",
            "update_model": "ROLLING",
            "snapshot_status": "STORED",
            "provenance_status": "OFFICIAL_ORIGINAL",
        }]
        self.assertEqual(leader.build_human_review_queue(records, date(2026, 8, 28)), [])

    def test_pending_selection_is_sent_to_human_operation_desktop(self):
        record = {
            "source_id": "PA006",
            "source_title": "Conditional regulation",
            "source_status": "CURRENT",
            "download_status": "SUCCESS",
            "snapshot_status": "STORED",
            "official_url": "https://example.test/conditional",
            "retrieval_url": "https://example.test/conditional",
            "provenance_status": "OFFICIAL_ORIGINAL",
            "operator_selection_decision": "PENDING",
            "authority_quality": "HIGH",
            "scope_relevance": "MEDIUM",
            "version_currency": "HIGH",
            "traceability": "HIGH",
            "access_permission": "HIGH",
        }
        queue = leader.build_human_review_queue([record], date(2026, 8, 31))
        self.assertEqual(len(queue), 1)
        self.assertEqual(queue[0]["trigger"], "SELECTION_PENDING")

    def test_recently_confirmed_pending_selection_is_temporarily_suppressed(self):
        record = {
            "source_id": "PA006",
            "source_status": "CURRENT",
            "operator_selection_decision": "PENDING",
            "authority_quality": "HIGH",
            "scope_relevance": "MEDIUM",
            "version_currency": "HIGH",
            "traceability": "HIGH",
            "access_permission": "HIGH",
        }
        fingerprint = leader.selection_review_fingerprint(record)
        history = {("PA006", "SELECTION_PENDING"): {"completed_at": date(2026, 8, 1), "fingerprint": fingerprint}}
        self.assertEqual(leader.build_human_review_queue([record], date(2026, 8, 31), 60, history), [])
        overdue = leader.build_human_review_queue([record], date(2026, 10, 2), 60, history)
        self.assertEqual(overdue[0]["trigger"], "SELECTION_PENDING")

    def test_withdrawn_or_excluded_source_does_not_create_review_task(self):
        withdrawn = {
            "source_id": "PA029",
            "download_status": "FAIL",
            "source_status": "WITHDRAWN",
        }
        excluded = {
            "source_id": "PA099",
            "download_status": "FAIL",
            "source_status": "CURRENT",
            "issuer": "Test authority",
            "official_url": "https://example.test/source",
            "acquisition_channel": "OFFICIAL_WEBSITE",
            "provenance_status": "OFFICIAL_ORIGINAL",
            "authority_quality": "HIGH",
            "scope_relevance": "LOW",
            "version_currency": "HIGH",
            "traceability": "HIGH",
            "access_permission": "HIGH",
            "operator_selection_decision": "EXCLUDE",
        }
        self.assertEqual(leader.build_human_review_queue([withdrawn, excluded], date(2026, 8, 31)), [])

    def test_completed_manual_fallback_closes_download_review(self):
        records = [{
            "source_id": "PA010",
            "source_title": "Manually recovered source",
            "download_status": "FAIL",
            "last_attempt_at": datetime(2026, 8, 28, 10, 0),
            "snapshot_status": "STORED",
            "current_origin": "MANUAL",
            "manual_update_date": date(2026, 8, 28),
            "manual_updated_by": "Reviewer",
        }]
        self.assertEqual(leader.build_human_review_queue(records, date(2026, 8, 28)), [])

    def test_recorded_rolling_impact_review_closes_change_review(self):
        records = [{
            "source_id": "PA011",
            "source_title": "Rolling guidance",
            "download_status": "SUCCESS",
            "content_change": "CHANGED",
            "update_model": "ROLLING",
            "last_success_at": datetime(2026, 8, 20, 10, 0),
            "manual_update_date": date(2026, 8, 21),
            "manual_updated_by": "Reviewer",
            "snapshot_status": "STORED",
        }]
        self.assertEqual(leader.build_human_review_queue(records, date(2026, 8, 28)), [])

    def test_review_outcome_is_not_a_process_failure(self):
        self.assertEqual(leader.completed_exit_code("REVIEW"), 0)
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            config = {
                "log_root": root,
                "leader": {"review_threshold_days": 60},
            }
            report = {"overall_status": "REVIEW"}
            path = leader.write_leader_report(config, report)
            self.assertTrue(path.exists())
            self.assertIn("REVIEW", path.read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
