#!/usr/bin/env python3
"""WP-18 Target-05 determinism+boundary replay harness."""

from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from typing import Any

STRUCTURAL_RULE_ORDER = [
    "STRUCT_SLOT_ORDER_CONTIGUOUS_ASC",
    "STRUCT_CAPTURED_PLAN_TERMINAL_REQUIRED",
    "STRUCT_ILLEGAL_SLOT_COMBINATION",
    "STRUCT_NO_NON_FORMATTER_AFTER_TERMINAL",
    "STRUCT_FORMATTER_COUNT_MAX_ONE",
]

SEMANTIC_RULE_ORDER = [
    "SEM_REQUIRED_DEPENDENCIES_PRESENT",
    "SEM_FORMATTER_TERMINAL_ADJACENT_AND_TYPE_COMPATIBLE",
    "SEM_ENTITY_FAMILY_OPERATOR_TERMINAL_COMPATIBLE",
    "SEM_AMBIGUOUS_OR_INCOMPATIBLE_COMBINATION_REFUSED",
]

DETERMINISM_RULE_ORDER = [
    "DET_RULE_EVALUATION_ORDER_STABLE",
    "DET_ERROR_WARNING_ORDER_STABLE",
    "DET_OUTPUT_NORMALIZATION_STABLE",
    "DET_AMBIGUOUS_INTERPRETATION_REFUSED",
    "DET_REPLAY_IDENTITY_STABLE",
]

BOUNDARY_RULE_ORDER = [
    "BOUND_VALIDATION_RUNTIME_INDEPENDENT",
    "BOUND_NO_TRIO_OR_ENGINE_APPLY_CALLS",
    "BOUND_CAPTURED_PLAN_READ_ONLY",
    "BOUND_NO_RUNTIME_SIDE_EFFECT_INFERENCE",
]

EXPECTED_RULE_ORDER = (
    STRUCTURAL_RULE_ORDER
    + SEMANTIC_RULE_ORDER
    + DETERMINISM_RULE_ORDER
    + BOUNDARY_RULE_ORDER
)


def find_repo_root(start: Path) -> Path:
    current = start.resolve()
    for candidate in [current, *current.parents]:
        if (candidate / "AGENTS.md").exists():
            return candidate
    raise FileNotFoundError(f"Unable to locate repo root from: {start}")


class DeterminismRuleLayerTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.repo_root = find_repo_root(Path(__file__).resolve().parent)
        cls.runner = cls.repo_root / "tests" / "wp-18" / "validator" / "validator_runner.py"
        cls.good_fixture = (
            cls.repo_root
            / "tests"
            / "wp-18"
            / "fixtures"
            / "good"
            / "good_semantic_terminal_adjacent_formatter_compatible.json"
        )
        cls.refused_fixture = (
            cls.repo_root
            / "tests"
            / "wp-18"
            / "fixtures"
            / "bad"
            / "bad_semantic_ambiguous_combination_refused.json"
        )

    def run_validator(self, fixture: Path, output_path: Path) -> tuple[int, dict[str, Any]]:
        cmd = [
            sys.executable,
            str(self.runner),
            "--input",
            str(fixture),
            "--output",
            str(output_path),
        ]
        completed = subprocess.run(
            cmd,
            cwd=self.repo_root,
            capture_output=True,
            text=True,
            check=False,
        )
        with output_path.open("r", encoding="utf-8") as handle:
            payload = json.load(handle)
        return completed.returncode, payload

    def first_result(self, payload: dict[str, Any]) -> dict[str, Any]:
        self.assertEqual(payload.get("artifact_count"), 1)
        return payload["results"][0]

    def rule_ids(self, result_row: dict[str, Any]) -> list[str]:
        return [
            str(item["rule_id"])
            for item in result_row["validation_result"]["rule_evaluations"]
        ]

    def find_rule_eval(self, result_row: dict[str, Any], rule_id: str) -> dict[str, Any]:
        for item in result_row["validation_result"]["rule_evaluations"]:
            if str(item["rule_id"]) == rule_id:
                return item
        raise AssertionError(f"Missing rule evaluation: {rule_id}")

    def test_replay_identity_stable_on_valid_input(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            out_a = Path(temp_dir) / "good_a.json"
            out_b = Path(temp_dir) / "good_b.json"
            rc_a, payload_a = self.run_validator(self.good_fixture, out_a)
            rc_b, payload_b = self.run_validator(self.good_fixture, out_b)

        self.assertEqual(rc_a, 0)
        self.assertEqual(rc_b, 0)
        self.assertEqual(payload_a, payload_b)

        row = self.first_result(payload_a)
        self.assertEqual(row["validation_result"]["status"], "PASS")
        self.assertEqual(self.rule_ids(row), EXPECTED_RULE_ORDER)

        det_eval = self.find_rule_eval(row, "DET_REPLAY_IDENTITY_STABLE")
        self.assertEqual(det_eval["outcome"], "PASS")
        self.assertIn("replay_identity=", det_eval["detail"])

    def test_replay_identity_stable_on_refused_input(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            out_a = Path(temp_dir) / "refused_a.json"
            out_b = Path(temp_dir) / "refused_b.json"
            rc_a, payload_a = self.run_validator(self.refused_fixture, out_a)
            rc_b, payload_b = self.run_validator(self.refused_fixture, out_b)

        self.assertEqual(rc_a, 2)
        self.assertEqual(rc_b, 2)
        self.assertEqual(payload_a, payload_b)

        row = self.first_result(payload_a)
        self.assertEqual(row["validation_result"]["status"], "REFUSE")
        self.assertEqual(self.rule_ids(row), EXPECTED_RULE_ORDER)

        det_ambiguity = self.find_rule_eval(row, "DET_AMBIGUOUS_INTERPRETATION_REFUSED")
        self.assertEqual(det_ambiguity["outcome"], "PASS")

        det_replay = self.find_rule_eval(row, "DET_REPLAY_IDENTITY_STABLE")
        self.assertEqual(det_replay["outcome"], "PASS")
        self.assertIn("replay_identity=", det_replay["detail"])

    def test_rule_order_emission_stable_across_pass_and_refuse(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            out_good = Path(temp_dir) / "good.json"
            out_bad = Path(temp_dir) / "bad.json"
            _, payload_good = self.run_validator(self.good_fixture, out_good)
            _, payload_bad = self.run_validator(self.refused_fixture, out_bad)

        row_good = self.first_result(payload_good)
        row_bad = self.first_result(payload_bad)
        self.assertEqual(self.rule_ids(row_good), EXPECTED_RULE_ORDER)
        self.assertEqual(self.rule_ids(row_bad), EXPECTED_RULE_ORDER)


if __name__ == "__main__":
    unittest.main()
