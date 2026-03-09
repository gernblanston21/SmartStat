#!/usr/bin/env python3
"""Deterministic replay harness stub for WP-18 Target-01."""

from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


def find_repo_root(start: Path) -> Path:
    current = start.resolve()
    for candidate in [current, *current.parents]:
        if (candidate / "AGENTS.md").exists():
            return candidate
    raise FileNotFoundError(f"Unable to locate repo root from: {start}")


class DeterministicReplayTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.repo_root = find_repo_root(Path(__file__).resolve().parent)
        cls.runner = cls.repo_root / "tests" / "wp-18" / "validator" / "validator_runner.py"
        cls.fixture = (
            cls.repo_root
            / "tests"
            / "wp-18"
            / "fixtures"
            / "good"
            / "good_schema_compatible_capture.json"
        )

    def run_validator(self, output_path: Path) -> dict:
        cmd = [
            sys.executable,
            str(self.runner),
            "--input",
            str(self.fixture),
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

        self.assertEqual(
            completed.returncode,
            0,
            msg=(
                "validator_runner.py did not return PASS for replay fixture.\n"
                f"stdout:\n{completed.stdout}\n"
                f"stderr:\n{completed.stderr}"
            ),
        )

        with output_path.open("r", encoding="utf-8") as handle:
            return json.load(handle)

    def test_deterministic_replay(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            run_a_path = Path(temp_dir) / "run_a.json"
            run_b_path = Path(temp_dir) / "run_b.json"

            run_a = self.run_validator(run_a_path)
            run_b = self.run_validator(run_b_path)

        self.assertEqual(run_a, run_b, "Full validator outputs differ between replay runs.")

        hash_a = run_a["results"][0]["validation_result"]["normalized_plan_hash"]
        hash_b = run_b["results"][0]["validation_result"]["normalized_plan_hash"]
        self.assertEqual(hash_a, hash_b, "normalized_plan_hash differs between replay runs.")


if __name__ == "__main__":
    unittest.main()

