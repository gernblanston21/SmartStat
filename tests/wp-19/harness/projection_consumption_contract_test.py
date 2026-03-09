#!/usr/bin/env python3
"""WP-19 Target-05 projection consumption contract consolidation harness."""

from __future__ import annotations

import hashlib
import json
import unittest
from pathlib import Path
from typing import Any

PROJECTION_CONTRACT = "wp19.viewer_projection.v1"
PROJECTION_KIND = "read_only_view_model"

AUTHORITATIVE_TOP_LEVEL_ORDER = [
    "projection_contract",
    "projection_kind",
    "input_artifact",
    "input_identity",
    "status_summary",
    "issues_summary",
    "rule_evaluation_summary",
    "deterministic_identity_summary",
    "semantic_interpretation_summary",
]

DISPLAY_SUMMARY_FIELDS = {
    "status_summary",
    "issues_summary",
    "rule_evaluation_summary",
    "semantic_interpretation_summary",
}

TRACEABILITY_FIELDS = {
    "projection_contract",
    "projection_kind",
    "input_artifact",
    "input_identity",
    "deterministic_identity_summary",
}

EXPECTED_PHASE_ORDER = ["STRUCTURAL", "SEMANTIC", "DETERMINISM", "BOUNDARY"]
EXPECTED_STATUS_VALUES = {"PASS", "REFUSE"}
EXPECTED_SCOPE_RESOLUTION_VALUES = {
    "explicit",
    "implicit_default",
    "not_applicable",
    "unknown",
}

INPUT_IDENTITY_KEYS = ["artifact_path", "input_fingerprint_sha256"]
STATUS_SUMMARY_KEYS = ["status", "error_count", "warning_count"]
ISSUES_SUMMARY_KEYS = ["errors", "warnings"]
RULE_SUMMARY_KEYS = ["category", "rule_id", "outcome"]
RULE_EVALUATION_SUMMARY_KEYS = ["phase_order", "ordered_rules"]
DETERMINISTIC_IDENTITY_KEYS = [
    "normalized_plan_hash",
    "replay_identity",
    "validator_run_identity",
]
SEMANTIC_INTERPRETATION_KEYS = [
    "scope_resolution",
    "effective_scope",
    "evidence_source",
]

FORBIDDEN_FIELDS = {
    "runtime_result",
    "apply_result",
    "bridge_result",
    "engine_result",
    "trio_command",
    "execution_trace",
    "runtime_side_effects",
}


def find_repo_root(start: Path) -> Path:
    current = start.resolve()
    for candidate in [current, *current.parents]:
        if (candidate / "AGENTS.md").exists():
            return candidate
    raise FileNotFoundError(f"Unable to locate repo root from: {start}")


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        digest.update(handle.read())
    return digest.hexdigest()


def load_json(path: Path) -> Any:
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def collect_object_keys(value: Any, keys: set[str]) -> None:
    if isinstance(value, dict):
        keys.update(value.keys())
        for nested in value.values():
            collect_object_keys(nested, keys)
    elif isinstance(value, list):
        for nested in value:
            collect_object_keys(nested, keys)


class ProjectionConsumptionContractTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.repo_root = find_repo_root(Path(__file__).resolve().parent)
        fixture_dir = cls.repo_root / "tests" / "wp-19" / "target-03" / "fixtures"
        cls.fixture_paths = [
            fixture_dir / "projection_pass_case.json",
            fixture_dir / "projection_refuse_case.json",
            fixture_dir / "projection_stats_implicit_default_scope_case.json",
            fixture_dir / "projection_stats_explicit_scope_case.json",
        ]

    def test_projection_consumption_contract_is_consolidated_and_stable(self) -> None:
        before_hashes = {path: file_sha256(path) for path in self.fixture_paths}

        self.assertTrue(DISPLAY_SUMMARY_FIELDS.isdisjoint(TRACEABILITY_FIELDS))
        self.assertEqual(
            DISPLAY_SUMMARY_FIELDS | TRACEABILITY_FIELDS,
            set(AUTHORITATIVE_TOP_LEVEL_ORDER),
        )

        for fixture_path in self.fixture_paths:
            projection = load_json(fixture_path)
            self.assert_projection_contract_shape(projection)
            self.assert_consumption_surface_split(projection)
            self.assert_non_goal_boundaries(projection)

        after_hashes = {path: file_sha256(path) for path in self.fixture_paths}
        self.assertEqual(before_hashes, after_hashes)

    def assert_projection_contract_shape(self, projection: dict[str, Any]) -> None:
        self.assertEqual(list(projection.keys()), AUTHORITATIVE_TOP_LEVEL_ORDER)
        self.assertEqual(projection["projection_contract"], PROJECTION_CONTRACT)
        self.assertEqual(projection["projection_kind"], PROJECTION_KIND)
        self.assertEqual(
            list(projection["input_identity"].keys()),
            INPUT_IDENTITY_KEYS,
        )
        self.assertEqual(
            list(projection["status_summary"].keys()),
            STATUS_SUMMARY_KEYS,
        )
        self.assertEqual(
            list(projection["issues_summary"].keys()),
            ISSUES_SUMMARY_KEYS,
        )
        self.assertEqual(
            list(projection["rule_evaluation_summary"].keys()),
            RULE_EVALUATION_SUMMARY_KEYS,
        )
        self.assertEqual(
            list(projection["deterministic_identity_summary"].keys()),
            DETERMINISTIC_IDENTITY_KEYS,
        )
        self.assertEqual(
            list(projection["semantic_interpretation_summary"].keys()),
            SEMANTIC_INTERPRETATION_KEYS,
        )

        self.assertIn(projection["status_summary"]["status"], EXPECTED_STATUS_VALUES)
        self.assertEqual(
            projection["rule_evaluation_summary"]["phase_order"],
            EXPECTED_PHASE_ORDER,
        )
        for rule in projection["rule_evaluation_summary"]["ordered_rules"]:
            self.assertEqual(list(rule.keys()), RULE_SUMMARY_KEYS)
        self.assertIn(
            projection["semantic_interpretation_summary"]["scope_resolution"],
            EXPECTED_SCOPE_RESOLUTION_VALUES,
        )

    def assert_consumption_surface_split(self, projection: dict[str, Any]) -> None:
        display_surface = {key: projection[key] for key in DISPLAY_SUMMARY_FIELDS}
        traceability_surface = {key: projection[key] for key in TRACEABILITY_FIELDS}

        self.assertEqual(set(display_surface.keys()), DISPLAY_SUMMARY_FIELDS)
        self.assertEqual(set(traceability_surface.keys()), TRACEABILITY_FIELDS)

        self.assertNotIn("deterministic_identity_summary", display_surface)
        self.assertNotIn("issues_summary", traceability_surface)
        self.assertNotIn("status_summary", traceability_surface)

    def assert_non_goal_boundaries(self, projection: dict[str, Any]) -> None:
        all_keys: set[str] = set()
        collect_object_keys(projection, all_keys)
        self.assertTrue(FORBIDDEN_FIELDS.isdisjoint(all_keys))


if __name__ == "__main__":
    unittest.main()
