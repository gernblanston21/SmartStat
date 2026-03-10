#!/usr/bin/env python3
"""WP-19 Target-06 projection-to-view-model adapter contract harness."""

from __future__ import annotations

import copy
import hashlib
import json
import unittest
from pathlib import Path
from typing import Any

ADAPTER_CONTRACT = "wp19.projection_to_view_model_adapter.v1"
SOURCE_PROJECTION_CONTRACT = "wp19.viewer_projection.v1"
SOURCE_PROJECTION_KIND = "read_only_view_model"

ADAPTER_TOP_LEVEL_ORDER = [
    "adapter_contract",
    "source_projection_contract",
    "source_projection_kind",
    "view_model",
]

VIEW_MODEL_TOP_LEVEL_ORDER = [
    "status_view",
    "issues_view",
    "rules_view",
    "semantic_view",
    "trace_view",
]

STATUS_VIEW_KEYS = ["status", "error_count", "warning_count"]
ISSUES_VIEW_KEYS = ["errors", "warnings"]
RULES_VIEW_KEYS = ["phase_order", "ordered_rules"]
SEMANTIC_VIEW_KEYS = ["scope_resolution", "effective_scope", "evidence_source"]
TRACE_VIEW_KEYS = [
    "projection_contract",
    "projection_kind",
    "input_artifact",
    "input_identity",
    "deterministic_identity_summary",
]

RULE_SUMMARY_KEYS = ["category", "rule_id", "outcome"]
EXPECTED_PHASE_ORDER = ["STRUCTURAL", "SEMANTIC", "DETERMINISM", "BOUNDARY"]
EXPECTED_STATUS_VALUES = {"PASS", "REFUSE"}

FORBIDDEN_FIELDS = {
    "runtime_result",
    "apply_result",
    "bridge_result",
    "engine_result",
    "trio_command",
    "execution_trace",
    "runtime_side_effects",
}

MAPPING_RULES = [
    ("status_summary.status", "view_model.status_view.status"),
    ("status_summary.error_count", "view_model.status_view.error_count"),
    ("status_summary.warning_count", "view_model.status_view.warning_count"),
    ("issues_summary.errors", "view_model.issues_view.errors"),
    ("issues_summary.warnings", "view_model.issues_view.warnings"),
    ("rule_evaluation_summary.phase_order", "view_model.rules_view.phase_order"),
    ("rule_evaluation_summary.ordered_rules", "view_model.rules_view.ordered_rules"),
    (
        "semantic_interpretation_summary.scope_resolution",
        "view_model.semantic_view.scope_resolution",
    ),
    (
        "semantic_interpretation_summary.effective_scope",
        "view_model.semantic_view.effective_scope",
    ),
    (
        "semantic_interpretation_summary.evidence_source",
        "view_model.semantic_view.evidence_source",
    ),
    ("projection_contract", "view_model.trace_view.projection_contract"),
    ("projection_kind", "view_model.trace_view.projection_kind"),
    ("input_artifact", "view_model.trace_view.input_artifact"),
    ("input_identity", "view_model.trace_view.input_identity"),
    (
        "deterministic_identity_summary",
        "view_model.trace_view.deterministic_identity_summary",
    ),
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


def resolve_path(data: dict[str, Any], dotted_path: str) -> Any:
    current: Any = data
    for token in dotted_path.split("."):
        current = current[token]
    return current


def adapt_projection_to_view_model(projection: dict[str, Any]) -> dict[str, Any]:
    return {
        "adapter_contract": ADAPTER_CONTRACT,
        "source_projection_contract": projection["projection_contract"],
        "source_projection_kind": projection["projection_kind"],
        "view_model": {
            "status_view": {
                "status": projection["status_summary"]["status"],
                "error_count": projection["status_summary"]["error_count"],
                "warning_count": projection["status_summary"]["warning_count"],
            },
            "issues_view": {
                "errors": copy.deepcopy(projection["issues_summary"]["errors"]),
                "warnings": copy.deepcopy(projection["issues_summary"]["warnings"]),
            },
            "rules_view": {
                "phase_order": copy.deepcopy(
                    projection["rule_evaluation_summary"]["phase_order"]
                ),
                "ordered_rules": copy.deepcopy(
                    projection["rule_evaluation_summary"]["ordered_rules"]
                ),
            },
            "semantic_view": {
                "scope_resolution": projection["semantic_interpretation_summary"][
                    "scope_resolution"
                ],
                "effective_scope": projection["semantic_interpretation_summary"][
                    "effective_scope"
                ],
                "evidence_source": projection["semantic_interpretation_summary"][
                    "evidence_source"
                ],
            },
            "trace_view": {
                "projection_contract": projection["projection_contract"],
                "projection_kind": projection["projection_kind"],
                "input_artifact": projection["input_artifact"],
                "input_identity": copy.deepcopy(projection["input_identity"]),
                "deterministic_identity_summary": copy.deepcopy(
                    projection["deterministic_identity_summary"]
                ),
            },
        },
    }


class ProjectionToViewModelAdapterContractTest(unittest.TestCase):
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

    def test_projection_to_view_model_adapter_contract(self) -> None:
        hashes_before = {path: file_sha256(path) for path in self.fixture_paths}

        for fixture_path in self.fixture_paths:
            projection = load_json(fixture_path)
            projection_snapshot = copy.deepcopy(projection)

            adapter_a = adapt_projection_to_view_model(projection)
            adapter_b = adapt_projection_to_view_model(projection)

            self.assertEqual(projection, projection_snapshot)
            self.assertEqual(adapter_a, adapter_b)
            self.assertEqual(
                json.dumps(adapter_a, separators=(",", ":"), ensure_ascii=False),
                json.dumps(adapter_b, separators=(",", ":"), ensure_ascii=False),
            )

            self.assert_adapter_shape(adapter_a)
            self.assert_mapping_rules(projection, adapter_a)
            self.assert_non_goal_boundaries(adapter_a)

        hashes_after = {path: file_sha256(path) for path in self.fixture_paths}
        self.assertEqual(hashes_before, hashes_after)

    def assert_adapter_shape(self, adapter_output: dict[str, Any]) -> None:
        self.assertEqual(list(adapter_output.keys()), ADAPTER_TOP_LEVEL_ORDER)
        self.assertEqual(adapter_output["adapter_contract"], ADAPTER_CONTRACT)
        self.assertEqual(
            adapter_output["source_projection_contract"], SOURCE_PROJECTION_CONTRACT
        )
        self.assertEqual(adapter_output["source_projection_kind"], SOURCE_PROJECTION_KIND)

        view_model = adapter_output["view_model"]
        self.assertEqual(list(view_model.keys()), VIEW_MODEL_TOP_LEVEL_ORDER)
        self.assertEqual(list(view_model["status_view"].keys()), STATUS_VIEW_KEYS)
        self.assertEqual(list(view_model["issues_view"].keys()), ISSUES_VIEW_KEYS)
        self.assertEqual(list(view_model["rules_view"].keys()), RULES_VIEW_KEYS)
        self.assertEqual(list(view_model["semantic_view"].keys()), SEMANTIC_VIEW_KEYS)
        self.assertEqual(list(view_model["trace_view"].keys()), TRACE_VIEW_KEYS)

        self.assertIn(view_model["status_view"]["status"], EXPECTED_STATUS_VALUES)
        self.assertEqual(view_model["rules_view"]["phase_order"], EXPECTED_PHASE_ORDER)
        for rule in view_model["rules_view"]["ordered_rules"]:
            self.assertEqual(list(rule.keys()), RULE_SUMMARY_KEYS)

    def assert_mapping_rules(
        self, projection: dict[str, Any], adapter_output: dict[str, Any]
    ) -> None:
        for source_path, target_path in MAPPING_RULES:
            self.assertEqual(
                resolve_path(projection, source_path),
                resolve_path(adapter_output, target_path),
                msg=f"Mapping mismatch: {source_path} -> {target_path}",
            )

    def assert_non_goal_boundaries(self, adapter_output: dict[str, Any]) -> None:
        all_keys: set[str] = set()
        collect_object_keys(adapter_output, all_keys)
        self.assertTrue(FORBIDDEN_FIELDS.isdisjoint(all_keys))


if __name__ == "__main__":
    unittest.main()
