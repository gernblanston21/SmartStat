#!/usr/bin/env python3
"""
WP-18 Target-02 validator runner.

This runner performs:
- captured-plan JSON loading
- schema compatibility checks against docs/onair/plan-capture.schema.json
- structural rule evaluation (only)
- deterministic validation_result emission

This runner intentionally does not implement semantic rules.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from pathlib import Path
from typing import Any

VALIDATOR_CONTRACT = "wp18.validator_runner.structural.v1"
HASH_PLACEHOLDER_CONTRACT = "wp18.normalized_plan_hash.placeholder.v1"

SCHEMA_RULE_ID = "SCHEMA_COMPATIBILITY"

RULE_SLOT_ORDER_CONTIGUOUS_ASC = "STRUCT_SLOT_ORDER_CONTIGUOUS_ASC"
RULE_CAPTURED_PLAN_TERMINAL_REQUIRED = "STRUCT_CAPTURED_PLAN_TERMINAL_REQUIRED"
RULE_ILLEGAL_SLOT_COMBINATION = "STRUCT_ILLEGAL_SLOT_COMBINATION"
RULE_NO_NON_FORMATTER_AFTER_TERMINAL = "STRUCT_NO_NON_FORMATTER_AFTER_TERMINAL"
RULE_FORMATTER_COUNT_MAX_ONE = "STRUCT_FORMATTER_COUNT_MAX_ONE"

STRUCTURAL_RULE_ORDER = [
    RULE_SLOT_ORDER_CONTIGUOUS_ASC,
    RULE_CAPTURED_PLAN_TERMINAL_REQUIRED,
    RULE_ILLEGAL_SLOT_COMBINATION,
    RULE_NO_NON_FORMATTER_AFTER_TERMINAL,
    RULE_FORMATTER_COUNT_MAX_ONE,
]

TERMINAL_SLOT_CLASSES = {"terminal_measure", "terminal_attribute"}


def find_repo_root(start: Path) -> Path:
    current = start.resolve()
    for candidate in [current, *current.parents]:
        if (candidate / "AGENTS.md").exists():
            return candidate
    raise FileNotFoundError(f"Unable to locate repo root from: {start}")


def canonical_json(value: Any) -> str:
    return json.dumps(value, ensure_ascii=True, separators=(",", ":"), sort_keys=True)


def make_issue(
    code: str,
    message: str,
    rule_id: str = SCHEMA_RULE_ID,
    slot_order: int | None = None,
) -> dict[str, Any]:
    return {
        "code": code,
        "message": message,
        "rule_id": rule_id,
        "slot_order": slot_order,
    }


def make_rule_evaluation(rule_id: str, outcome: str, detail: str) -> dict[str, str]:
    return {
        "rule_id": rule_id,
        "category": "STRUCTURAL",
        "outcome": outcome,
        "detail": detail,
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


def get_slots(payload: Any) -> list[dict[str, Any]]:
    if not isinstance(payload, dict):
        return []
    raw_slots = payload.get("slot_sequence")
    if not isinstance(raw_slots, list):
        return []
    return [slot for slot in raw_slots if isinstance(slot, dict)]


def terminal_orders(slots: list[dict[str, Any]]) -> list[int]:
    orders: list[int] = []
    for slot in slots:
        slot_class = slot.get("slot_class")
        order = slot.get("order")
        if slot_class in TERMINAL_SLOT_CLASSES and isinstance(order, int):
            orders.append(order)
    return sorted(orders)


def skipped_structural_evaluations() -> list[dict[str, str]]:
    return [
        make_rule_evaluation(
            rule_id=rule_id,
            outcome="WARN",
            detail="Skipped because schema compatibility failed.",
        )
        for rule_id in STRUCTURAL_RULE_ORDER
    ]


def evaluate_structural_rules(payload: Any) -> tuple[list[dict[str, str]], list[dict[str, Any]]]:
    evaluations: list[dict[str, str]] = []
    errors: list[dict[str, Any]] = []

    if not isinstance(payload, dict):
        return (skipped_structural_evaluations(), errors)

    slots = get_slots(payload)
    slot_orders = [slot.get("order") for slot in slots if isinstance(slot.get("order"), int)]
    artifact_type = payload.get("artifact_type")
    terminal_definition = payload.get("terminal")

    # 1) Slot ordering contiguous and ascending, no gaps.
    expected_orders = list(range(1, len(slot_orders) + 1))
    if slot_orders == expected_orders:
        evaluations.append(
            make_rule_evaluation(
                RULE_SLOT_ORDER_CONTIGUOUS_ASC,
                "PASS",
                f"slot_sequence.order values are contiguous and ascending: {slot_orders}.",
            )
        )
    else:
        evaluations.append(
            make_rule_evaluation(
                RULE_SLOT_ORDER_CONTIGUOUS_ASC,
                "REFUSE",
                f"slot_sequence.order values are not contiguous ascending. expected={expected_orders} actual={slot_orders}.",
            )
        )
        errors.append(
            make_issue(
                "STRUCT_SLOT_ORDER_NON_CONTIGUOUS",
                "slot_sequence.order must be contiguous and ascending starting at 1.",
                rule_id=RULE_SLOT_ORDER_CONTIGUOUS_ASC,
                slot_order=None,
            )
        )

    # 2) captured_plan artifacts must include terminal definition.
    slot_terminal_orders = terminal_orders(slots)
    has_terminal_definition = isinstance(terminal_definition, dict)
    has_terminal_slot = len(slot_terminal_orders) >= 1
    if artifact_type == "captured_plan":
        if has_terminal_definition and has_terminal_slot:
            evaluations.append(
                make_rule_evaluation(
                    RULE_CAPTURED_PLAN_TERMINAL_REQUIRED,
                    "PASS",
                    "captured_plan has terminal definition and terminal slot.",
                )
            )
        else:
            evaluations.append(
                make_rule_evaluation(
                    RULE_CAPTURED_PLAN_TERMINAL_REQUIRED,
                    "REFUSE",
                    "captured_plan requires both terminal definition and terminal slot.",
                )
            )
            errors.append(
                make_issue(
                    "STRUCT_MISSING_TERMINAL_DEFINITION",
                    "captured_plan must include a terminal definition and terminal slot.",
                    rule_id=RULE_CAPTURED_PLAN_TERMINAL_REQUIRED,
                    slot_order=None,
                )
            )
    else:
        evaluations.append(
            make_rule_evaluation(
                RULE_CAPTURED_PLAN_TERMINAL_REQUIRED,
                "PASS",
                f"Rule applies to captured_plan only; artifact_type='{artifact_type}'.",
            )
        )

    # 3) Illegal slot combinations.
    slot_classes = [str(slot.get("slot_class")) for slot in slots]
    class_set = set(slot_classes)
    illegal_reasons: list[str] = []
    if "family" in class_set and "operator" in class_set:
        illegal_reasons.append("family and operator cannot both be present")
    if "terminal_measure" in class_set and "terminal_attribute" in class_set:
        illegal_reasons.append("terminal_measure and terminal_attribute cannot both be present")

    if illegal_reasons:
        detail = "; ".join(sorted(illegal_reasons))
        evaluations.append(
            make_rule_evaluation(
                RULE_ILLEGAL_SLOT_COMBINATION,
                "REFUSE",
                f"Illegal slot combination detected: {detail}.",
            )
        )
        errors.append(
            make_issue(
                "STRUCT_ILLEGAL_SLOT_COMBINATION",
                f"Illegal slot combination detected: {detail}.",
                rule_id=RULE_ILLEGAL_SLOT_COMBINATION,
                slot_order=None,
            )
        )
    else:
        evaluations.append(
            make_rule_evaluation(
                RULE_ILLEGAL_SLOT_COMBINATION,
                "PASS",
                "No illegal slot combinations detected.",
            )
        )

    # 4) No non-formatter slots after terminal slot.
    if slot_terminal_orders:
        first_terminal = slot_terminal_orders[0]
        offending_slots: list[tuple[int, str]] = []
        for slot in slots:
            order = slot.get("order")
            slot_class = str(slot.get("slot_class"))
            if isinstance(order, int) and order > first_terminal and slot_class != "formatter":
                offending_slots.append((order, slot_class))

        offending_slots.sort(key=lambda item: (item[0], item[1]))
        if offending_slots:
            first_offender_order, first_offender_class = offending_slots[0]
            evaluations.append(
                make_rule_evaluation(
                    RULE_NO_NON_FORMATTER_AFTER_TERMINAL,
                    "REFUSE",
                    f"Non-formatter slot after terminal: order={first_offender_order}, slot_class={first_offender_class}.",
                )
            )
            errors.append(
                make_issue(
                    "STRUCT_NON_FORMATTER_AFTER_TERMINAL",
                    "Non-formatter slot appears after terminal slot.",
                    rule_id=RULE_NO_NON_FORMATTER_AFTER_TERMINAL,
                    slot_order=first_offender_order,
                )
            )
        else:
            evaluations.append(
                make_rule_evaluation(
                    RULE_NO_NON_FORMATTER_AFTER_TERMINAL,
                    "PASS",
                    "No non-formatter slots appear after terminal slot.",
                )
            )
    else:
        evaluations.append(
            make_rule_evaluation(
                RULE_NO_NON_FORMATTER_AFTER_TERMINAL,
                "WARN",
                "Skipped because no terminal slot is available.",
            )
        )

    # 5) Formatter count at most one.
    formatter_orders = sorted(
        [
            slot.get("order")
            for slot in slots
            if slot.get("slot_class") == "formatter" and isinstance(slot.get("order"), int)
        ]
    )
    if len(formatter_orders) <= 1:
        evaluations.append(
            make_rule_evaluation(
                RULE_FORMATTER_COUNT_MAX_ONE,
                "PASS",
                f"Formatter count is within limit: {len(formatter_orders)}.",
            )
        )
    else:
        second_formatter_order = formatter_orders[1]
        evaluations.append(
            make_rule_evaluation(
                RULE_FORMATTER_COUNT_MAX_ONE,
                "REFUSE",
                f"Formatter count exceeds limit: {len(formatter_orders)}.",
            )
        )
        errors.append(
            make_issue(
                "STRUCT_MULTIPLE_FORMATTERS",
                "At most one formatter slot is allowed.",
                rule_id=RULE_FORMATTER_COUNT_MAX_ONE,
                slot_order=second_formatter_order,
            )
        )

    return (evaluations, errors)


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="WP-18 validator runner (structural rules only).")
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

    if errors:
        rule_evaluations = skipped_structural_evaluations()
    else:
        structural_evals, structural_errors = evaluate_structural_rules(payload)
        rule_evaluations = structural_evals
        errors.extend(structural_errors)

    errors.sort(
        key=lambda item: (
            str(item.get("code", "")),
            str(item.get("rule_id", "")),
            str(item.get("message", "")),
            str(item.get("slot_order", "")),
        )
    )
    warnings.sort(
        key=lambda item: (
            str(item.get("code", "")),
            str(item.get("rule_id", "")),
            str(item.get("message", "")),
            str(item.get("slot_order", "")),
        )
    )

    validation_result = {
        "status": "PASS" if not errors else "REFUSE",
        "errors": errors,
        "warnings": warnings,
        "normalized_plan_hash": build_normalized_plan_hash(payload, raw_text),
        "rule_evaluations": rule_evaluations,
    }

    return {
        "input_artifact": rel_path(artifact_path, repo_root),
        "validation_result": validation_result,
    }


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

    all_pass = all(row["validation_result"]["status"] == "PASS" for row in results)
    return 0 if all_pass else 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))

