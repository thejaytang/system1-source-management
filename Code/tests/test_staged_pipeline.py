from __future__ import annotations

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
import staged_pipeline as pipeline


class StagedPipelineTests(unittest.TestCase):
    def test_discovery_and_assessment_contracts(self):
        record = {
            "source_id": "PA001",
            "official_url": "https://example.test/source",
            "issuer": "Test Authority",
            "source_family": "Public authority",
            "document_type": "Law",
            "requirement_role": "Primary normative",
            "inclusion_rationale": "Official source selected for an in-scope duty.",
            "authority_quality": "HIGH",
            "scope_relevance": "HIGH",
            "version_currency": "HIGH",
            "traceability": "HIGH",
            "access_permission": "HIGH",
            "automation_readiness": "MEDIUM",
            "operator_selection_decision": "INCLUDE",
            "download_status": "SUCCESS",
            "snapshot_status": "STORED",
            "acquisition_channel": "OFFICIAL_WEBSITE",
            "provenance_status": "OFFICIAL_ORIGINAL",
            "update_model": "ROLLING",
        }
        discovery = pipeline.discovery_stage([record])
        assessment = pipeline.assessment_stage([record])
        self.assertEqual(discovery["status"], "PASS")
        self.assertEqual(discovery["source_count"], 1)
        self.assertEqual(assessment["status"], "PASS")
        self.assertEqual(assessment["selection_status_counts"], {"INCLUDE": 1})

    def test_assessment_stage_rejects_incomplete_scores(self):
        record = {
            "source_id": "PA002",
            "official_url": "https://example.test/source-2",
            "issuer": "Test Authority",
            "source_family": "Public authority",
            "document_type": "Regulation",
            "requirement_role": "Primary normative",
            "inclusion_rationale": "",
            "authority_quality": "HIGH",
            "scope_relevance": "MEDIUM",
            "version_currency": "",
            "traceability": "HIGH",
            "access_permission": "HIGH",
            "automation_readiness": "HIGH",
            "acquisition_channel": "OFFICIAL_WEBSITE",
            "provenance_status": "OFFICIAL_ORIGINAL",
            "update_model": "ROLLING",
        }
        result = pipeline.assessment_stage([record])
        self.assertEqual(result["status"], "REVIEW")
        self.assertTrue(any(issue["field"] == "inclusion_rationale" for issue in result["issues"]))
        self.assertTrue(any(issue["field"] == "version_currency" for issue in result["issues"]))

    def test_discovery_accepts_private_source_with_equivalent_provenance(self):
        record = {
            "source_id": "MS001",
            "official_url": "",
            "acquisition_channel": "CLIENT_OR_SITE",
            "provenance_status": "AUTHORISED_COPY",
        }
        result = pipeline.discovery_stage([record])
        self.assertEqual(result["status"], "PASS")
        self.assertEqual(result["no_official_url"], ["MS001"])
        self.assertEqual(result["equivalent_provenance_gaps"], [])

    def test_discovery_rejects_official_channel_without_official_url(self):
        record = {
            "source_id": "PA003",
            "official_url": "",
            "acquisition_channel": "OFFICIAL_WEBSITE",
            "provenance_status": "OFFICIAL_ORIGINAL",
        }
        result = pipeline.discovery_stage([record])
        self.assertEqual(result["status"], "REVIEW")
        self.assertEqual(result["equivalent_provenance_gaps"], ["PA003"])

    def test_equipment_assessment_requires_applicability(self):
        record = {
            "source_id": "MS002",
            "issuer": "OEM Test",
            "source_family": "Manufacturer / supplier",
            "document_type": "Equipment manual / OEM instruction",
            "requirement_role": "Asset-specific normative",
            "inclusion_rationale": "The installed equipment must operate within OEM limits.",
            "authority_quality": "HIGH",
            "scope_relevance": "HIGH",
            "version_currency": "HIGH",
            "traceability": "HIGH",
            "access_permission": "HIGH",
            "automation_readiness": "MEDIUM",
            "official_url": "",
            "acquisition_channel": "CLIENT_OR_SITE",
            "provenance_status": "AUTHORISED_COPY",
            "update_model": "VERSIONED",
            "applicability_reference": "",
        }
        result = pipeline.assessment_stage([record])
        self.assertEqual(result["status"], "REVIEW")
        self.assertEqual(result["selection_status_counts"], {"PENDING": 1})
        self.assertTrue(any(issue["field"] == "applicability_reference" for issue in result["issues"]))


if __name__ == "__main__":
    unittest.main()
