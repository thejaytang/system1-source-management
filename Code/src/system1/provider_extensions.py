"""Optional provider handoff for API-connected System1 extensions.

The shared project does not configure or call a real API.  This module defines
the narrow boundary that a future search or assessment provider may use.  A
provider can propose candidates and governed field updates, but every proposal
is routed through ``Human Operation Desktop`` before it can affect the source
register.
"""

from __future__ import annotations

import hashlib
import json
import tempfile
from pathlib import Path
from typing import Any, Protocol

import human_operations
import source_updater


class DiscoveryProvider(Protocol):
    name: str

    def discover(self, records: list[dict[str, Any]]) -> list[dict[str, Any]]: ...


class AssessmentProvider(Protocol):
    name: str

    def assess(self, records: list[dict[str, Any]]) -> list[dict[str, Any]]: ...


def _provider_name(provider: object, fallback: str) -> str:
    return str(getattr(provider, "name", fallback) or fallback).strip()


def _candidate_key(candidate: dict[str, Any]) -> str:
    stable = str(candidate.get("content_hash") or "").strip().lower()
    if not stable:
        stable = "|".join((
            str(candidate.get("official_url") or "").strip().lower(),
            str(candidate.get("source_title") or "").strip().lower(),
        ))
    return hashlib.sha256(stable.encode("utf-8")).hexdigest()


def _append_candidates(config_path: Path, candidates: list[dict[str, Any]]) -> int:
    config = source_updater.read_config(config_path)
    inbox = human_operations.discovery_inbox_path(config_path, config)
    existing = human_operations.load_discovery_candidates(config_path, config)
    seen = {_candidate_key(item) for item in existing}
    added = 0
    for candidate in candidates:
        key = _candidate_key(candidate)
        if key in seen:
            continue
        existing.append(candidate)
        seen.add(key)
        added += 1
    if not added:
        return 0
    inbox.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(
        "w", encoding="utf-8", delete=False, dir=inbox.parent,
        prefix=".provider-candidates-", suffix=".json",
    ) as handle:
        json.dump(existing, handle, ensure_ascii=False, indent=2)
        temporary = Path(handle.name)
    temporary.replace(inbox)
    return added


def _normalise_candidates(provider: object, raw: Any) -> list[dict[str, Any]]:
    if not isinstance(raw, list):
        raise ValueError("Discovery provider output must be a list of candidate objects.")
    name = _provider_name(provider, "discovery_provider")
    candidates: list[dict[str, Any]] = []
    for index, item in enumerate(raw, 1):
        if not isinstance(item, dict):
            raise ValueError(f"Discovery candidate {index} is not an object.")
        title = str(item.get("source_title") or "").strip()
        if not title:
            raise ValueError(f"Discovery candidate {index} has no source_title.")
        candidate = dict(item)
        candidate["candidate_origin"] = "API_DISCOVERY"
        candidate["provider_name"] = name
        candidate.setdefault(
            "notes",
            f"Draft candidate proposed by {name}; human acceptance is required.",
        )
        candidates.append(candidate)
    return candidates


def _normalise_assessments(
    provider: object,
    raw: Any,
    records: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    if not isinstance(raw, list):
        raise ValueError("Assessment provider output must be a list of review suggestions.")
    name = _provider_name(provider, "assessment_provider")
    source_titles = {
        str(record.get("source_id") or "").strip(): str(record.get("source_title") or "").strip()
        for record in records
    }
    items: list[dict[str, Any]] = []
    for index, suggestion in enumerate(raw, 1):
        if not isinstance(suggestion, dict):
            raise ValueError(f"Assessment suggestion {index} is not an object.")
        source_id = str(suggestion.get("source_id") or "").strip()
        if source_id not in source_titles:
            raise ValueError(f"Assessment suggestion {index} references an unknown source_id: {source_id or '<blank>'}.")
        updates = suggestion.get("proposed_updates") or {}
        if not isinstance(updates, dict):
            raise ValueError(f"Assessment suggestion {index} proposed_updates must be an object.")
        forbidden = sorted(set(updates) - human_operations.SOURCE_UPDATE_FIELDS)
        if forbidden:
            raise ValueError(
                f"Assessment suggestion {index} contains non-governed fields: {', '.join(forbidden)}."
            )
        reason = str(suggestion.get("reason") or "").strip()
        if not reason:
            raise ValueError(f"Assessment suggestion {index} has no reason.")
        stable = json.dumps(
            {"provider": name, "source_id": source_id, "reason": reason, "updates": updates},
            sort_keys=True,
        )
        items.append({
            "review_id": "api" + hashlib.sha256(stable.encode("utf-8")).hexdigest()[:13],
            "source_id": source_id,
            "source_title": source_titles[source_id],
            "trigger": "API_ASSESSMENT_SUGGESTION",
            "review_since": None,
            "days_open": None,
            "severity": "OPEN",
            "reason": f"{name}: {reason}",
            "recommended_action": str(suggestion.get("recommended_action") or "Review the proposed governed field updates."),
            "proposed_updates": updates,
            "review_fingerprint": "",
        })
    return items


def collect_provider_extensions(
    config_path: Path,
    records: list[dict[str, Any]],
    providers: dict[str, object] | None = None,
) -> dict[str, Any]:
    """Collect optional provider proposals without granting write authority.

    Provider errors are isolated.  The deterministic Leader can continue and
    reports a degraded optional extension instead of losing local operations.
    """
    providers = providers or {}
    discovery = providers.get("discovery")
    assessment = providers.get("assessment")
    if discovery is None and assessment is None:
        return {
            "mode": "NO_API",
            "status": "DISABLED",
            "candidate_count": 0,
            "queued_candidate_count": 0,
            "review_items": [],
            "errors": [],
        }

    result: dict[str, Any] = {
        "mode": "API_CONNECTED",
        "status": "PASS",
        "candidate_count": 0,
        "queued_candidate_count": 0,
        "review_items": [],
        "errors": [],
    }
    if discovery is not None:
        try:
            candidates = _normalise_candidates(discovery, discovery.discover(records))
            result["candidate_count"] = len(candidates)
            result["queued_candidate_count"] = _append_candidates(config_path, candidates)
        except Exception as exc:
            result["errors"].append({
                "provider": _provider_name(discovery, "discovery_provider"),
                "stage": "discovery",
                "error": f"{type(exc).__name__}: {exc}",
            })
    if assessment is not None:
        try:
            result["review_items"] = _normalise_assessments(assessment, assessment.assess(records), records)
        except Exception as exc:
            result["errors"].append({
                "provider": _provider_name(assessment, "assessment_provider"),
                "stage": "assessment",
                "error": f"{type(exc).__name__}: {exc}",
            })
    if result["errors"]:
        result["status"] = "DEGRADED"
    return result
