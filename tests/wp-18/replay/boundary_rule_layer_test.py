#!/usr/bin/env python3
"""WP-18 Target-05 boundary-rule harness."""

from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from typing import Any

BOUNDARY_RULE_IDS = [
    "BOUND_VALIDATION_RUNTIME_INDEPENDENT",
    "BOUND_NO_TRIO_OR_ENGINE_APPLY_CALLS",
    "BOUND_CAPTURED_PLAN_READ_ONLY",
    "BOUND_NO_RUNTIME_SIDE_EFFECT_INFERENCE",
]


def find_repo_root(start: Path) -> Path:
    current = start.resolve()
    for candidate in [current, *current.parents]:
        if (candidate / "AGENTS.md").exists():
            return candidate
    raise FileNotFoundError(f"Unable to locate repo root from: {start}")


class BoundaryRuleLayerTest(unittest.TestCase):
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
            / "good_semantic_dependency_chain.json"
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

    def find_rule_eval(self, row: dict[str, Any], rule_id: str) -> dict[str, Any]:
        for item in row["validation_result"]["rule_evaluations"]:
            if str(item["rule_id"]) == rule_id:
                return item
        raise AssertionError(f"Missing rule evaluation: {rule_id}")

    def test_input_artifact_unchanged_after_validation(self) -> None:
        before = self.good_fixture.read_bytes()
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = Path(temp_dir) / "good.json"
            rc, payload = self.run_validator(self.good_fixture, output_path)
        after = self.good_fixture.read_bytes()

        self.assertEqual(rc, 0)
        self.assertEqual(before, after, "Validator mutated the captured-plan fixture.")

        row = self.first_result(payload)
        eval_read_only = self.find_rule_eval(row, "BOUND_CAPTURED_PLAN_READ_ONLY")
        self.assertEqual(eval_read_only["outcome"], "PASS")

    def test_output_contract_remains_architecture_only(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = Path(temp_dir) / "good.json"
            rc, payload = self.run_validator(self.good_fixture, output_path)

        self.assertEqual(rc, 0)

        self.assertEqual(
            set(payload.keys()),
            {"artifact_count", "results", "schema_path", "validator_contract"},
        )

        row = self.first_result(payload)
        result = row["validation_result"]
        self.assertEqual(
            set(result.keys()),
            {"status", "errors", "warnings", "normalized_plan_hash", "rule_evaluations"},
        )

        for issue in result["errors"] + result["warnings"]:
            self.assertEqual(set(issue.keys()), {"code", "message", "rule_id", "slot_order"})

        for evaluation in result["rule_evaluations"]:
            self.assertEqual(set(evaluation.keys()), {"rule_id", "category", "outcome", "detail"})

        for rule_id in BOUNDARY_RULE_IDS:
            boundary_eval = self.find_rule_eval(row, rule_id)
            self.assertEqual(boundary_eval["category"], "BOUNDARY")
            self.assertEqual(boundary_eval["outcome"], "PASS")

    def test_boundary_rules_pass_on_refused_semantic_plan(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = Path(temp_dir) / "refused.json"
            rc, payload = self.run_validator(self.refused_fixture, output_path)

        self.assertEqual(rc, 2)
        row = self.first_result(payload)
        self.assertEqual(row["validation_result"]["status"], "REFUSE")

        for rule_id in BOUNDARY_RULE_IDS:
            boundary_eval = self.find_rule_eval(row, rule_id)
            self.assertEqual(boundary_eval["outcome"], "PASS")


if __name__ == "__main__":
    unittest.main()
