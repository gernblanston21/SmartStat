#!/usr/bin/env python3
"""
WP-18 Target-01 validator scaffolding runner.

This runner intentionally performs only:
- captured-plan JSON loading
- schema compatibility checks against docs/onair/plan-capture.schema.json
- deterministic validation_result stub emission

It intentionally does not perform WP-18 validation rule evaluation yet.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from pathlib import Path
from typing import Any

VALIDATOR_CONTRACT = "wp18.validator_runner.stub.v1"
HASH_PLACEHOLDER_CONTRACT = "wp18.normalized_plan_hash.placeholder.v1"
SCHEMA_RULE_ID = "SCHEMA_COMPATIBILITY"


def find_repo_root(start: Path) -> Path:
    current = start.resolve()
    for candidate in [current, *current.parents]:
        if (candidate / "AGENTS.md").exists():
            return candidate
    raise FileNotFoundError(f"Unable to locate repo root from: {start}")


def canonical_json(value: Any) -> str:
    return json.dumps(value, ensure_ascii=True, separators=(",", ":"), sort_keys=True)


def make_issue(code: str, message: str) -> dict[str, Any]:
    return {
        "code": code,
        "message": message,
        "rule_id": SCHEMA_RULE_ID,
        "slot_order": None,
    }


def build_normalized_plan_hash(payload: Any, raw_text: str) -> str:
    if payload is None:
        canonical_payload = raw_text
    else:
        canonical_payload = canonical_json(payload)

    digest_input = (
        HASH_PLACEHOLDER_CONTRACT + "\n" + canonical_payload
    ).encode("utf-8")
    return hashlib.sha256(digest_input).hexdigest()


def json_path(parts: list[Any]) -> str:
    path = "$"
    for part in parts:
        if isinstance(part, int):
            path += f"[{part}]"
        else:
            path += f".{part}"
    return path


def validate_schema_with_jsonschema(payload: Any, schema: dict[str, Any]) -> list[dict[str, Any]] | None:
    try:
        import jsonschema  # type: ignore
    except Exception:
        return None

    validator = jsonschema.Draft202012Validator(schema)
    issues: list[dict[str, Any]] = []
    for error in validator.iter_errors(payload):
        message = f"{json_path(list(error.absolute_path))}: {error.message}"
        issues.append(make_issue("SCHEMA_VALIDATION_ERROR", message))

    issues.sort(key=lambda item: (item["code"], item["message"]))
    return issues


def validate_schema_fallback(payload: Any, schema: dict[str, Any]) -> list[dict[str, Any]]:
    issues: list[dict[str, Any]] = []

    if not isinstance(payload, dict):
        issues.append(make_issue("INVALID_TOP_LEVEL_TYPE", "Top-level JSON value must be an object."))
        return issues

    required = schema.get("required", [])
    for key in sorted(required):
        if key not in payload:
            issues.append(make_issue("MISSING_REQUIRED_KEY", f"Missing required key '{key}' at $."))

    if schema.get("additionalProperties") is False:
        known_keys = set(schema.get("properties", {}).keys())
        for key in sorted(payload.keys()):
            if key not in known_keys:
                issues.append(make_issue("UNKNOWN_TOP_LEVEL_KEY", f"Unknown key '{key}' at $."))

    properties = schema.get("properties", {})
    for key in sorted(properties.keys()):
        if key not in payload:
            continue

        prop_schema = properties[key]
        value = payload[key]

        if "const" in prop_schema and value != prop_schema["const"]:
            issues.append(
                make_issue(
                    "CONST_MISMATCH",
                    f"Key '{key}' must equal '{prop_schema['const']}'.",
                )
            )

        enum_values = prop_schema.get("enum")
        if enum_values is not None and value not in enum_values:
            issues.append(
                make_issue(
                    "ENUM_MISMATCH",
                    f"Key '{key}' value '{value}' is not in allowed enum.",
                )
            )

    issues.sort(key=lambda item: (item["code"], item["message"]))
    return issues


def collect_artifact_paths(input_path: Path) -> list[Path]:
    if input_path.is_file():
        if input_path.suffix.lower() != ".json":
            raise ValueError(f"Input file must be JSON: {input_path}")
        return [input_path.resolve()]

    if input_path.is_dir():
        files = sorted(
            (path.resolve() for path in input_path.rglob("*.json") if path.is_file()),
            key=lambda p: p.as_posix().lower(),
        )
        if not files:
            raise ValueError(f"No JSON artifacts found under directory: {input_path}")
        return files

    raise FileNotFoundError(f"Input path not found: {input_path}")


def rel_path(path: Path, repo_root: Path) -> str:
    try:
        return path.resolve().relative_to(repo_root.resolve()).as_posix()
    except ValueError:
        return path.resolve().as_posix()


def process_artifact(
    artifact_path: Path, schema: dict[str, Any], repo_root: Path
) -> dict[str, Any]:
    raw_text = artifact_path.read_text(encoding="utf-8")
    payload: Any | None = None
    errors: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []

    try:
        payload = json.loads(raw_text)
    except json.JSONDecodeError as exc:
        errors.append(
            make_issue(
                "JSON_PARSE_ERROR",
                f"Invalid JSON at {artifact_path.name}:{exc.lineno}:{exc.colno}.",
            )
        )

    if not errors:
        schema_issues = validate_schema_with_jsonschema(payload, schema)
        if schema_issues is None:
            warnings.append(
                make_issue(
                    "SCHEMA_VALIDATION_FALLBACK",
                    "jsonschema package unavailable; using minimal schema compatibility checks.",
                )
            )
            schema_issues = validate_schema_fallback(payload, schema)

        errors.extend(schema_issues)

    errors.sort(key=lambda item: (item["code"], item["message"]))
    warnings.sort(key=lambda item: (item["code"], item["message"]))

    validation_result = {
        "status": "PASS" if not errors else "REFUSE",
        "errors": errors,
        "warnings": warnings,
        "normalized_plan_hash": build_normalized_plan_hash(payload, raw_text),
        "rule_evaluations": [],
    }

    return {
        "input_artifact": rel_path(artifact_path, repo_root),
        "validation_result": validation_result,
    }


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="WP-18 validator scaffolding runner.")
    parser.add_argument(
        "--input",
        required=True,
        help="Input captured-plan JSON file or directory containing JSON files.",
    )
    parser.add_argument(
        "--schema",
        default="",
        help="Optional schema path. Defaults to docs/onair/plan-capture.schema.json.",
    )
    parser.add_argument(
        "--output",
        default="",
        help="Optional output JSON file path. If omitted, output is printed to stdout.",
    )
    return parser.parse_args(argv)


def main(argv: list[str]) -> int:
    args = parse_args(argv)

    script_dir = Path(__file__).resolve().parent
    repo_root = find_repo_root(script_dir)

    input_path = Path(args.input).resolve()
    schema_path = (
        Path(args.schema).resolve()
        if args.schema
        else (repo_root / "docs" / "onair" / "plan-capture.schema.json").resolve()
    )

    if not schema_path.exists():
        raise FileNotFoundError(f"Schema file not found: {schema_path}")

    schema = json.loads(schema_path.read_text(encoding="utf-8"))
    artifact_paths = collect_artifact_paths(input_path)

    results = [process_artifact(path, schema, repo_root) for path in artifact_paths]
    results.sort(key=lambda item: item["input_artifact"].lower())

    output_payload = {
        "artifact_count": len(results),
        "results": results,
        "schema_path": rel_path(schema_path, repo_root),
        "validator_contract": VALIDATOR_CONTRACT,
    }
    output_text = canonical_json(output_payload) + "\n"

    if args.output:
        output_path = Path(args.output).resolve()
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(output_text, encoding="utf-8")
    else:
        sys.stdout.write(output_text)

    all_pass = all(
        row["validation_result"]["status"] == "PASS" for row in results
    )
    return 0 if all_pass else 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))

