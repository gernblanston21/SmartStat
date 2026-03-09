#!/usr/bin/env python3
"""WP-19 Target-04 viewer projection summary contract hardening harness."""

from __future__ import annotations

import copy
import hashlib
import json
import unittest
from pathlib import Path
from typing import Any

PROJECTION_CONTRACT = "wp19.viewer_projection.v1"
EXPECTED_PHASE_ORDER = ["STRUCTURAL", "SEMANTIC", "DETERMINISM", "BOUNDARY"]

FORBIDDEN_VIEW_MODEL_FIELDS = {
    "runtime_result",
    "apply_result",
    "bridge_result",
    "engine_result",
    "trio_command",
    "execution_trace",
    "runtime_side_effects",
}

PROJECTION_KEY_ORDER = [
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
INPUT_IDENTITY_KEY_ORDER = ["artifact_path", "input_fingerprint_sha256"]
STATUS_SUMMARY_KEY_ORDER = ["status", "error_count", "warning_count"]
ISSUES_SUMMARY_KEY_ORDER = ["errors", "warnings"]
RULE_EVALUATION_SUMMARY_KEY_ORDER = ["phase_order", "ordered_rules"]
RULE_SUMMARY_KEY_ORDER = ["category", "rule_id", "outcome"]
DETERMINISTIC_IDENTITY_KEY_ORDER = [
    "normalized_plan_hash",
    "replay_identity",
    "validator_run_identity",
]
SEMANTIC_INTERPRETATION_KEY_ORDER = [
    "scope_resolution",
    "effective_scope",
    "evidence_source",
]


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


def get_row_by_suffix(payload: dict[str, Any], suffix: str) -> dict[str, Any]:
    for row in payload["results"]:
        if str(row["input_artifact"]).endswith(suffix):
            return row
    raise AssertionError(f"Unable to find artifact row ending with: {suffix}")


def build_projection(result_row: dict[str, Any]) -> dict[str, Any]:
    validation_result = result_row["validation_result"]
    errors = copy.deepcopy(validation_result["errors"])
    warnings = copy.deepcopy(validation_result["warnings"])
    ordered_rules = [
        {
            "category": rule_eval["category"],
            "rule_id": rule_eval["rule_id"],
            "outcome": rule_eval["outcome"],
        }
        for rule_eval in validation_result["rule_evaluations"]
    ]
    return {
        "projection_contract": PROJECTION_CONTRACT,
        "projection_kind": "read_only_view_model",
        "input_artifact": result_row["input_artifact"],
        "input_identity": {
            "artifact_path": result_row["input_identity"]["artifact_path"],
            "input_fingerprint_sha256": result_row["input_identity"][
                "input_fingerprint_sha256"
            ],
        },
        "status_summary": {
            "status": validation_result["status"],
            "error_count": len(errors),
            "warning_count": len(warnings),
        },
        "issues_summary": {
            "errors": errors,
            "warnings": warnings,
        },
        "rule_evaluation_summary": {
            "phase_order": EXPECTED_PHASE_ORDER,
            "ordered_rules": ordered_rules,
        },
        "deterministic_identity_summary": {
            "normalized_plan_hash": validation_result["normalized_plan_hash"],
            "replay_identity": validation_result["replay_identity"],
            "validator_run_identity": validation_result["validator_run_identity"],
        },
        "semantic_interpretation_summary": {
            "scope_resolution": validation_result["semantic_interpretation"][
                "scope_resolution"
            ],
            "effective_scope": validation_result["semantic_interpretation"][
                "effective_scope"
            ],
            "evidence_source": validation_result["semantic_interpretation"][
                "evidence_source"
            ],
        },
    }


class ViewerProjectionContractTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.repo_root = find_repo_root(Path(__file__).resolve().parent)
        cls.good_results_path = (
            cls.repo_root
            / "tests"
            / "wp-18"
            / "artifacts"
            / "wp18_validator_runs"
            / "target06"
            / "good_results.json"
        )
        cls.bad_results_path = (
            cls.repo_root
            / "tests"
            / "wp-18"
            / "artifacts"
            / "wp18_validator_runs"
            / "target06"
            / "bad_results.json"
        )
        cls.canonical_scope_results_path = (
            cls.repo_root
            / "tests"
            / "wp-18"
            / "artifacts"
            / "wp18_validator_runs"
            / "target06"
            / "canonical_scope_results.json"
        )
        cls.fixture_dir = cls.repo_root / "tests" / "wp-19" / "target-03" / "fixtures"
        cls.case_definitions = [
            {
                "case_id": "projection_pass_case",
                "source_path": cls.good_results_path,
                "artifact_suffix": "good_schema_compatible_capture.json",
                "fixture_name": "projection_pass_case.json",
            },
            {
                "case_id": "projection_refuse_case",
                "source_path": cls.bad_results_path,
                "artifact_suffix": "bad_semantic_formatter_not_terminal_adjacent.json",
                "fixture_name": "projection_refuse_case.json",
            },
            {
                "case_id": "projection_stats_implicit_default_scope_case",
                "source_path": cls.canonical_scope_results_path,
                "artifact_suffix": "stats_player_hits.json",
                "fixture_name": "projection_stats_implicit_default_scope_case.json",
            },
            {
                "case_id": "projection_stats_explicit_scope_case",
                "source_path": cls.canonical_scope_results_path,
                "artifact_suffix": "stats_player_season_month_april_hits.json",
                "fixture_name": "projection_stats_explicit_scope_case.json",
            },
        ]

    def test_projection_contract_fixtures_are_stable(self) -> None:
        source_paths = {
            self.good_results_path,
            self.bad_results_path,
            self.canonical_scope_results_path,
        }
        hashes_before = {path: file_sha256(path) for path in source_paths}

        for case in self.case_definitions:
            payload = load_json(case["source_path"])
            row = get_row_by_suffix(payload, case["artifact_suffix"])
            row_snapshot = copy.deepcopy(row)

            projection_a = build_projection(row)
            projection_b = build_projection(row)

            self.assertEqual(projection_a, projection_b)
            self.assertEqual(
                json.dumps(projection_a, separators=(",", ":"), ensure_ascii=False),
                json.dumps(projection_b, separators=(",", ":"), ensure_ascii=False),
            )
            self.assertEqual(row, row_snapshot)
            self.assert_projection_shape(projection_a)
            self.assert_summary_sections(projection_a, row)
            self.assert_order_preserved(projection_a, row)
            self.assert_projection_has_no_forbidden_fields(projection_a)

            fixture_path = self.fixture_dir / case["fixture_name"]
            expected_projection = load_json(fixture_path)
            self.assertEqual(projection_a, expected_projection)

        hashes_after = {path: file_sha256(path) for path in source_paths}
        self.assertEqual(hashes_before, hashes_after)

    def assert_projection_shape(self, projection: dict[str, Any]) -> None:
        self.assertEqual(list(projection.keys()), PROJECTION_KEY_ORDER)
        self.assertEqual(projection["projection_contract"], PROJECTION_CONTRACT)
        self.assertEqual(projection["projection_kind"], "read_only_view_model")

        self.assertEqual(
            list(projection["input_identity"].keys()), INPUT_IDENTITY_KEY_ORDER
        )
        self.assertEqual(
            list(projection["status_summary"].keys()), STATUS_SUMMARY_KEY_ORDER
        )
        self.assertEqual(
            list(projection["issues_summary"].keys()), ISSUES_SUMMARY_KEY_ORDER
        )
        self.assertEqual(
            list(projection["deterministic_identity_summary"].keys()),
            DETERMINISTIC_IDENTITY_KEY_ORDER,
        )
        self.assertEqual(
            list(projection["semantic_interpretation_summary"].keys()),
            SEMANTIC_INTERPRETATION_KEY_ORDER,
        )

        rule_summary = projection["rule_evaluation_summary"]
        self.assertEqual(
            list(rule_summary.keys()), RULE_EVALUATION_SUMMARY_KEY_ORDER
        )
        self.assertEqual(rule_summary["phase_order"], EXPECTED_PHASE_ORDER)
        for rule in rule_summary["ordered_rules"]:
            self.assertEqual(list(rule.keys()), RULE_SUMMARY_KEY_ORDER)

    def assert_summary_sections(
        self, projection: dict[str, Any], source_row: dict[str, Any]
    ) -> None:
        validation_result = source_row["validation_result"]
        issues_summary = projection["issues_summary"]
        status_summary = projection["status_summary"]
        deterministic_identity_summary = projection["deterministic_identity_summary"]
        semantic_summary = projection["semantic_interpretation_summary"]

        self.assertEqual(status_summary["status"], validation_result["status"])
        self.assertEqual(status_summary["error_count"], len(issues_summary["errors"]))
        self.assertEqual(status_summary["warning_count"], len(issues_summary["warnings"]))
        self.assertEqual(issues_summary["errors"], validation_result["errors"])
        self.assertEqual(issues_summary["warnings"], validation_result["warnings"])

        self.assertEqual(
            deterministic_identity_summary["normalized_plan_hash"],
            validation_result["normalized_plan_hash"],
        )
        self.assertEqual(
            deterministic_identity_summary["replay_identity"],
            validation_result["replay_identity"],
        )
        self.assertEqual(
            deterministic_identity_summary["validator_run_identity"],
            validation_result["validator_run_identity"],
        )

        self.assertEqual(
            semantic_summary["scope_resolution"],
            validation_result["semantic_interpretation"]["scope_resolution"],
        )
        self.assertEqual(
            semantic_summary["effective_scope"],
            validation_result["semantic_interpretation"]["effective_scope"],
        )
        self.assertEqual(
            semantic_summary["evidence_source"],
            validation_result["semantic_interpretation"]["evidence_source"],
        )

    def assert_order_preserved(
        self, projection: dict[str, Any], source_row: dict[str, Any]
    ) -> None:
        src_rule_evals = source_row["validation_result"]["rule_evaluations"]
        projected_rules = projection["rule_evaluation_summary"]["ordered_rules"]
        src_triplets = [
            {
                "category": rule_eval["category"],
                "rule_id": rule_eval["rule_id"],
                "outcome": rule_eval["outcome"],
            }
            for rule_eval in src_rule_evals
        ]
        self.assertEqual(projected_rules, src_triplets)

        src_phase_indexes = [
            EXPECTED_PHASE_ORDER.index(rule_eval["category"]) for rule_eval in src_rule_evals
        ]
        projected_phase_indexes = [
            EXPECTED_PHASE_ORDER.index(rule["category"]) for rule in projected_rules
        ]
        self.assertEqual(src_phase_indexes, sorted(src_phase_indexes))
        self.assertEqual(projected_phase_indexes, sorted(projected_phase_indexes))

    def assert_projection_has_no_forbidden_fields(self, projection: dict[str, Any]) -> None:
        all_keys: set[str] = set()
        collect_object_keys(projection, all_keys)
        self.assertTrue(FORBIDDEN_VIEW_MODEL_FIELDS.isdisjoint(all_keys))


if __name__ == "__main__":
    unittest.main()
