#!/usr/bin/env python3
"""WP-18 Target-06 result-model hardening harness."""

from __future__ import annotations

import hashlib
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from typing import Any

TOP_LEVEL_KEYS = ["artifact_count", "results", "schema_path", "validator_contract"]
RESULT_ROW_KEYS = ["input_artifact", "input_identity", "validation_result"]
INPUT_IDENTITY_KEYS = ["artifact_path", "input_fingerprint_sha256"]
VALIDATION_RESULT_KEYS = [
    "errors",
    "normalized_plan_hash",
    "replay_identity",
    "rule_evaluations",
    "semantic_interpretation",
    "status",
    "validator_run_identity",
    "warnings",
]
ISSUE_KEYS = ["code", "message", "rule_id", "slot_order"]
RULE_EVALUATION_KEYS = ["category", "detail", "outcome", "rule_id"]
SEMANTIC_INTERPRETATION_KEYS = ["effective_scope", "evidence_source", "scope_resolution"]

ALL_RULE_CATEGORIES = {"STRUCTURAL", "SEMANTIC", "DETERMINISM", "BOUNDARY"}

SOURCE_DICTIONARY_BY_SLOT_CLASS = {
    "family": "query_skeletons",
    "entity": "entity_dictionary",
    "scope": "filter_grammar_dictionary",
    "filter": "filter_grammar_dictionary",
    "terminal_measure": "measure_dictionary",
}


def find_repo_root(start: Path) -> Path:
    current = start.resolve()
    for candidate in [current, *current.parents]:
        if (candidate / "AGENTS.md").exists():
            return candidate
    raise FileNotFoundError(f"Unable to locate repo root from: {start}")


def make_slot(order: int, slot_class: str, token: str, parameters: list[str] | None = None) -> dict[str, Any]:
    return {
        "order": order,
        "slot_class": slot_class,
        "token": token,
        "token_kind": "canonical",
        "parameters": [] if parameters is None else parameters,
        "source_dictionary": SOURCE_DICTIONARY_BY_SLOT_CLASS[slot_class],
        "candidate_status": "resolved",
    }


def build_captured_plan(
    query_text: str,
    normalized_query_text: str,
    slots: list[dict[str, Any]],
) -> dict[str, Any]:
    fingerprint = hashlib.sha256(query_text.encode("utf-8")).hexdigest().upper()
    terminal_slot = next(slot for slot in slots if slot["slot_class"] == "terminal_measure")
    return {
        "contract_version": "wp17.plan_capture.v1",
        "artifact_type": "captured_plan",
        "mode": "read_only_contract",
        "source": {
            "query_text": query_text,
            "normalized_query_text": normalized_query_text,
            "semantic_record_id": "measure:mlb:hits",
            "semantic_record_type": "measure",
            "league": "mlb",
        },
        "determinism": {
            "ordering_contract": "slot_sequence.order_asc.v1",
            "canonicalization_version": "wp17.canonical_json.v1",
            "input_fingerprint_sha256": fingerprint,
        },
        "slot_sequence": slots,
        "terminal": {
            "slot_class": terminal_slot["slot_class"],
            "token": terminal_slot["token"],
        },
        "refusal": None,
        "deferred_boundaries": [
            "runtime apply behavior",
            "planner execution",
            "runtime bridge behavior",
        ],
    }


class ResultModelHardeningTest(unittest.TestCase):
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

    def assert_hardened_shape(self, payload: dict[str, Any]) -> None:
        self.assertEqual(list(payload.keys()), sorted(TOP_LEVEL_KEYS))
        self.assertEqual(set(payload.keys()), set(TOP_LEVEL_KEYS))
        self.assertEqual(payload["artifact_count"], len(payload["results"]))

        for row in payload["results"]:
            self.assertEqual(list(row.keys()), sorted(RESULT_ROW_KEYS))
            self.assertEqual(set(row.keys()), set(RESULT_ROW_KEYS))

            input_identity = row["input_identity"]
            self.assertEqual(list(input_identity.keys()), sorted(INPUT_IDENTITY_KEYS))
            self.assertEqual(set(input_identity.keys()), set(INPUT_IDENTITY_KEYS))

            result = row["validation_result"]
            self.assertEqual(list(result.keys()), sorted(VALIDATION_RESULT_KEYS))
            self.assertEqual(set(result.keys()), set(VALIDATION_RESULT_KEYS))

            interpretation = result["semantic_interpretation"]
            self.assertEqual(list(interpretation.keys()), sorted(SEMANTIC_INTERPRETATION_KEYS))
            self.assertEqual(set(interpretation.keys()), set(SEMANTIC_INTERPRETATION_KEYS))

            for issue in result["errors"] + result["warnings"]:
                self.assertEqual(list(issue.keys()), sorted(ISSUE_KEYS))
                self.assertEqual(set(issue.keys()), set(ISSUE_KEYS))

            categories = set()
            for rule_eval in result["rule_evaluations"]:
                self.assertEqual(list(rule_eval.keys()), sorted(RULE_EVALUATION_KEYS))
                self.assertEqual(set(rule_eval.keys()), set(RULE_EVALUATION_KEYS))
                categories.add(str(rule_eval["category"]))
            self.assertEqual(categories, ALL_RULE_CATEGORIES)

    def test_result_model_shape_on_pass_and_refuse(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            pass_output = Path(temp_dir) / "pass.json"
            refuse_output = Path(temp_dir) / "refuse.json"
            rc_pass, payload_pass = self.run_validator(self.good_fixture, pass_output)
            rc_refuse, payload_refuse = self.run_validator(self.refused_fixture, refuse_output)

        self.assertEqual(rc_pass, 0)
        self.assertEqual(rc_refuse, 2)
        self.assert_hardened_shape(payload_pass)
        self.assert_hardened_shape(payload_refuse)

    def test_canonical_scope_example_interpretation(self) -> None:
        cases = [
            (
                "stats_player_hits",
                "{{stats.player.hits}}",
                "stats.player.hits",
                [
                    make_slot(1, "family", "stats"),
                    make_slot(2, "entity", "player"),
                    make_slot(3, "terminal_measure", "hits"),
                ],
                {
                    "scope_resolution": "implicit_default",
                    "effective_scope": "career",
                    "evidence_source": "operator_grounded_default",
                },
            ),
            (
                "stats_player_career_hits",
                "{{stats.player.career.hits}}",
                "stats.player.career.hits",
                [
                    make_slot(1, "family", "stats"),
                    make_slot(2, "entity", "player"),
                    make_slot(3, "scope", "career"),
                    make_slot(4, "terminal_measure", "hits"),
                ],
                {
                    "scope_resolution": "explicit",
                    "effective_scope": "career",
                    "evidence_source": "artifact_explicit",
                },
            ),
            (
                "stats_player_month_april_hits",
                "{{stats.player.month(april).hits}}",
                "stats.player.month(april).hits",
                [
                    make_slot(1, "family", "stats"),
                    make_slot(2, "entity", "player"),
                    make_slot(3, "filter", "month", ["april"]),
                    make_slot(4, "terminal_measure", "hits"),
                ],
                {
                    "scope_resolution": "implicit_default",
                    "effective_scope": "career",
                    "evidence_source": "operator_grounded_default",
                },
            ),
            (
                "stats_player_career_month_april_hits",
                "{{stats.player.career.month(april).hits}}",
                "stats.player.career.month(april).hits",
                [
                    make_slot(1, "family", "stats"),
                    make_slot(2, "entity", "player"),
                    make_slot(3, "scope", "career"),
                    make_slot(4, "filter", "month", ["april"]),
                    make_slot(5, "terminal_measure", "hits"),
                ],
                {
                    "scope_resolution": "explicit",
                    "effective_scope": "career",
                    "evidence_source": "artifact_explicit",
                },
            ),
            (
                "stats_player_season_month_april_hits",
                "{{stats.player.season.month(april).hits}}",
                "stats.player.season.month(april).hits",
                [
                    make_slot(1, "family", "stats"),
                    make_slot(2, "entity", "player"),
                    make_slot(3, "scope", "season"),
                    make_slot(4, "filter", "month", ["april"]),
                    make_slot(5, "terminal_measure", "hits"),
                ],
                {
                    "scope_resolution": "explicit",
                    "effective_scope": "season",
                    "evidence_source": "artifact_explicit",
                },
            ),
        ]

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            for name, query_text, normalized_query_text, slots, expected_interpretation in cases:
                fixture_path = temp_root / f"{name}.json"
                output_path = temp_root / f"{name}.out.json"
                payload = build_captured_plan(query_text, normalized_query_text, slots)
                fixture_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")

                rc, run_payload = self.run_validator(fixture_path, output_path)
                self.assertEqual(rc, 0, msg=f"Canonical scope case failed: {name}")
                self.assertEqual(run_payload["artifact_count"], 1)

                row = run_payload["results"][0]
                result = row["validation_result"]
                self.assertEqual(result["status"], "PASS")
                self.assertEqual(result["semantic_interpretation"], expected_interpretation)


if __name__ == "__main__":
    unittest.main()
