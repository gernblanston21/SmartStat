#!/usr/bin/env python3
"""WP-19 Target-02 read-only intake contract harness."""

from __future__ import annotations

import copy
import ast
import hashlib
import json
import unittest
from pathlib import Path
from typing import Any

EXPECTED_PHASE_ORDER = ["STRUCTURAL", "SEMANTIC", "DETERMINISM", "BOUNDARY"]

EXPECTED_RULE_IDS_BY_CATEGORY = {
    "STRUCTURAL": [
        "STRUCT_SLOT_ORDER_CONTIGUOUS_ASC",
        "STRUCT_CAPTURED_PLAN_TERMINAL_REQUIRED",
        "STRUCT_ILLEGAL_SLOT_COMBINATION",
        "STRUCT_NO_NON_FORMATTER_AFTER_TERMINAL",
        "STRUCT_FORMATTER_COUNT_MAX_ONE",
    ],
    "SEMANTIC": [
        "SEM_REQUIRED_DEPENDENCIES_PRESENT",
        "SEM_FORMATTER_TERMINAL_ADJACENT_AND_TYPE_COMPATIBLE",
        "SEM_ENTITY_FAMILY_OPERATOR_TERMINAL_COMPATIBLE",
        "SEM_AMBIGUOUS_OR_INCOMPATIBLE_COMBINATION_REFUSED",
    ],
    "DETERMINISM": [
        "DET_RULE_EVALUATION_ORDER_STABLE",
        "DET_ERROR_WARNING_ORDER_STABLE",
        "DET_OUTPUT_NORMALIZATION_STABLE",
        "DET_AMBIGUOUS_INTERPRETATION_REFUSED",
        "DET_REPLAY_IDENTITY_STABLE",
    ],
    "BOUNDARY": [
        "BOUND_VALIDATION_RUNTIME_INDEPENDENT",
        "BOUND_NO_TRIO_OR_ENGINE_APPLY_CALLS",
        "BOUND_CAPTURED_PLAN_READ_ONLY",
        "BOUND_NO_RUNTIME_SIDE_EFFECT_INFERENCE",
    ],
}

ISSUE_KEYS = {"code", "message", "rule_id", "slot_order"}
SEMANTIC_INTERPRETATION_KEYS = {"scope_resolution", "effective_scope", "evidence_source"}
VALIDATION_RESULT_KEYS = {
    "errors",
    "normalized_plan_hash",
    "replay_identity",
    "rule_evaluations",
    "semantic_interpretation",
    "status",
    "validator_run_identity",
    "warnings",
}
RULE_EVAL_KEYS = {"category", "detail", "outcome", "rule_id"}
ROW_KEYS = {"input_artifact", "input_identity", "validation_result"}
INPUT_IDENTITY_KEYS = {"artifact_path", "input_fingerprint_sha256"}
TOP_LEVEL_KEYS = {"artifact_count", "results", "schema_path", "validator_contract"}

FORBIDDEN_VIEW_MODEL_FIELDS = {
    "runtime_result",
    "apply_result",
    "bridge_result",
    "engine_result",
    "trio_command",
    "execution_trace",
    "runtime_side_effects",
}

FORBIDDEN_IMPORT_MODULES = {
    "subprocess",
    "socket",
    "win32com",
    "win32com.client",
}
FORBIDDEN_CALLS = {
    "subprocess.run",
    "subprocess.popen",
    "socket.socket",
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


def project_view_model(result_row: dict[str, Any]) -> dict[str, Any]:
    """Project a minimal read-only viewer contract shape from WP-18 output."""
    validation_result = result_row["validation_result"]
    return {
        "input_identity": copy.deepcopy(result_row["input_identity"]),
        "status": validation_result["status"],
        "errors": copy.deepcopy(validation_result["errors"]),
        "warnings": copy.deepcopy(validation_result["warnings"]),
        "rule_evaluations": copy.deepcopy(validation_result["rule_evaluations"]),
        "deterministic_identities": {
            "normalized_plan_hash": validation_result["normalized_plan_hash"],
            "replay_identity": validation_result["replay_identity"],
            "validator_run_identity": validation_result["validator_run_identity"],
        },
        "semantic_interpretation": copy.deepcopy(validation_result["semantic_interpretation"]),
    }


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


class ReadOnlyIntakeContractTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.repo_root = find_repo_root(Path(__file__).resolve().parent)
        cls.wp17_captured_plan = (
            cls.repo_root / "tests" / "wp-17" / "fixtures" / "good" / "good_captured_plan_basic.json"
        )
        cls.wp18_good_results = (
            cls.repo_root
            / "tests"
            / "wp-18"
            / "artifacts"
            / "wp18_validator_runs"
            / "target06"
            / "good_results.json"
        )
        cls.wp18_bad_results = (
            cls.repo_root
            / "tests"
            / "wp-18"
            / "artifacts"
            / "wp18_validator_runs"
            / "target06"
            / "bad_results.json"
        )

    def test_wp17_captured_plan_is_read_only_intake(self) -> None:
        before_hash = file_sha256(self.wp17_captured_plan)
        payload = load_json(self.wp17_captured_plan)

        self.assertEqual(payload["contract_version"], "wp17.plan_capture.v1")
        self.assertEqual(payload["artifact_type"], "captured_plan")
        self.assertEqual(payload["mode"], "read_only_contract")
        self.assertIsInstance(payload["slot_sequence"], list)
        self.assertGreater(len(payload["slot_sequence"]), 0)
        self.assertIn("terminal", payload)
        self.assertIn("deferred_boundaries", payload)

        snapshot = copy.deepcopy(payload)
        intake_summary = {
            "query_text": payload["source"]["query_text"],
            "slot_count": len(payload["slot_sequence"]),
            "terminal_token": payload["terminal"]["token"],
        }
        self.assertEqual(intake_summary["slot_count"], len(payload["slot_sequence"]))
        self.assertEqual(payload, snapshot)
        self.assertEqual(before_hash, file_sha256(self.wp17_captured_plan))

    def test_wp18_validation_payload_shape_has_required_viewer_inputs(self) -> None:
        for path in [self.wp18_good_results, self.wp18_bad_results]:
            before_hash = file_sha256(path)
            payload = load_json(path)

            self.assertEqual(set(payload.keys()), TOP_LEVEL_KEYS)
            self.assertEqual(payload["artifact_count"], len(payload["results"]))

            for row in payload["results"]:
                self.assertEqual(set(row.keys()), ROW_KEYS)
                self.assertEqual(set(row["input_identity"].keys()), INPUT_IDENTITY_KEYS)

                result = row["validation_result"]
                self.assertEqual(set(result.keys()), VALIDATION_RESULT_KEYS)
                self.assertEqual(
                    set(result["semantic_interpretation"].keys()), SEMANTIC_INTERPRETATION_KEYS
                )

                for issue in result["errors"] + result["warnings"]:
                    self.assertEqual(set(issue.keys()), ISSUE_KEYS)

                for rule_eval in result["rule_evaluations"]:
                    self.assertEqual(set(rule_eval.keys()), RULE_EVAL_KEYS)

            self.assertEqual(before_hash, file_sha256(path))

    def test_ordering_surfaces_are_preserved_for_projection(self) -> None:
        good_payload = load_json(self.wp18_good_results)
        bad_payload = load_json(self.wp18_bad_results)
        rows_to_check = [
            good_payload["results"][0],
            get_row_by_suffix(
                bad_payload, "bad_semantic_formatter_not_terminal_adjacent.json"
            ),
        ]

        for row in rows_to_check:
            original_rule_evals = row["validation_result"]["rule_evaluations"]
            projected = project_view_model(row)
            projected_rule_evals = projected["rule_evaluations"]

            self.assertEqual(
                [rule_eval["rule_id"] for rule_eval in projected_rule_evals],
                [rule_eval["rule_id"] for rule_eval in original_rule_evals],
            )
            self.assertEqual(
                [rule_eval["category"] for rule_eval in projected_rule_evals],
                [rule_eval["category"] for rule_eval in original_rule_evals],
            )

            phase_indexes = [
                EXPECTED_PHASE_ORDER.index(str(rule_eval["category"]))
                for rule_eval in projected_rule_evals
            ]
            self.assertEqual(phase_indexes, sorted(phase_indexes))

            for category in EXPECTED_PHASE_ORDER:
                category_rule_ids = [
                    rule_eval["rule_id"]
                    for rule_eval in projected_rule_evals
                    if rule_eval["category"] == category
                ]
                self.assertEqual(category_rule_ids, EXPECTED_RULE_IDS_BY_CATEGORY[category])

            self.assertEqual(projected["errors"], row["validation_result"]["errors"])
            self.assertEqual(projected["warnings"], row["validation_result"]["warnings"])

    def test_boundary_constraints_no_mutation_no_runtime_surfaces(self) -> None:
        good_before = file_sha256(self.wp18_good_results)
        bad_before = file_sha256(self.wp18_bad_results)

        good_payload = load_json(self.wp18_good_results)
        bad_payload = load_json(self.wp18_bad_results)

        for row in [good_payload["results"][0], bad_payload["results"][0]]:
            projected = project_view_model(row)
            all_keys: set[str] = set()
            collect_object_keys(projected, all_keys)
            self.assertTrue(FORBIDDEN_VIEW_MODEL_FIELDS.isdisjoint(all_keys))

            boundary_rule_ids = [
                rule_eval["rule_id"]
                for rule_eval in projected["rule_evaluations"]
                if rule_eval["category"] == "BOUNDARY"
            ]
            self.assertEqual(
                boundary_rule_ids, EXPECTED_RULE_IDS_BY_CATEGORY["BOUNDARY"]
            )

        harness_dir = self.repo_root / "tests" / "wp-19" / "harness"
        for py_file in harness_dir.glob("*.py"):
            source = py_file.read_text(encoding="utf-8")
            tree = ast.parse(source)

            imported_modules: set[str] = set()
            called_functions: set[str] = set()

            for node in ast.walk(tree):
                if isinstance(node, ast.Import):
                    for alias in node.names:
                        imported_modules.add(alias.name.lower())
                elif isinstance(node, ast.ImportFrom):
                    if node.module is not None:
                        imported_modules.add(node.module.lower())
                elif isinstance(node, ast.Call):
                    called_name = self._extract_called_name(node.func)
                    if called_name:
                        called_functions.add(called_name.lower())

            for module_name in FORBIDDEN_IMPORT_MODULES:
                self.assertNotIn(
                    module_name,
                    imported_modules,
                    msg=f"Forbidden runtime module import '{module_name}' in {py_file}",
                )
            for call_name in FORBIDDEN_CALLS:
                self.assertNotIn(
                    call_name,
                    called_functions,
                    msg=f"Forbidden runtime call '{call_name}' in {py_file}",
                )

        self.assertEqual(good_before, file_sha256(self.wp18_good_results))
        self.assertEqual(bad_before, file_sha256(self.wp18_bad_results))

    @staticmethod
    def _extract_called_name(node: ast.AST) -> str:
        if isinstance(node, ast.Name):
            return node.id
        if isinstance(node, ast.Attribute):
            parent = ReadOnlyIntakeContractTest._extract_called_name(node.value)
            return f"{parent}.{node.attr}" if parent else node.attr
        return ""


if __name__ == "__main__":
    unittest.main()
