#!/usr/bin/env python3
"""
WP-18 Target-05 validator runner.

This runner performs:
- captured-plan JSON loading
- schema compatibility checks against docs/onair/plan-capture.schema.json
- structural rule evaluation
- semantic rule evaluation
- determinism rule evaluation
- boundary rule evaluation
- deterministic validation_result emission

This runner is validation-only, deterministic, read-only, and runtime-independent.
"""

from __future__ import annotations

import argparse
import ast
import hashlib
import json
import sys
from pathlib import Path
from typing import Any

VALIDATOR_CONTRACT = "wp18.validator_runner.boundary.v1"
HASH_PLACEHOLDER_CONTRACT = "wp18.normalized_plan_hash.placeholder.v1"
REPLAY_IDENTITY_CONTRACT = "wp18.replay_identity.v1"

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

RULE_REQUIRED_DEPENDENCIES_PRESENT = "SEM_REQUIRED_DEPENDENCIES_PRESENT"
RULE_FORMATTER_TERMINAL_ADJACENT_AND_TYPE_COMPATIBLE = (
    "SEM_FORMATTER_TERMINAL_ADJACENT_AND_TYPE_COMPATIBLE"
)
RULE_ENTITY_FAMILY_OPERATOR_TERMINAL_COMPATIBLE = "SEM_ENTITY_FAMILY_OPERATOR_TERMINAL_COMPATIBLE"
RULE_AMBIGUOUS_OR_INCOMPATIBLE_COMBINATION_REFUSED = (
    "SEM_AMBIGUOUS_OR_INCOMPATIBLE_COMBINATION_REFUSED"
)

SEMANTIC_RULE_ORDER = [
    RULE_REQUIRED_DEPENDENCIES_PRESENT,
    RULE_FORMATTER_TERMINAL_ADJACENT_AND_TYPE_COMPATIBLE,
    RULE_ENTITY_FAMILY_OPERATOR_TERMINAL_COMPATIBLE,
    RULE_AMBIGUOUS_OR_INCOMPATIBLE_COMBINATION_REFUSED,
]

RULE_DET_RULE_EVALUATION_ORDER_STABLE = "DET_RULE_EVALUATION_ORDER_STABLE"
RULE_DET_ERROR_WARNING_ORDER_STABLE = "DET_ERROR_WARNING_ORDER_STABLE"
RULE_DET_OUTPUT_NORMALIZATION_STABLE = "DET_OUTPUT_NORMALIZATION_STABLE"
RULE_DET_AMBIGUOUS_INTERPRETATION_REFUSED = "DET_AMBIGUOUS_INTERPRETATION_REFUSED"
RULE_DET_REPLAY_IDENTITY_STABLE = "DET_REPLAY_IDENTITY_STABLE"

DETERMINISM_RULE_ORDER = [
    RULE_DET_RULE_EVALUATION_ORDER_STABLE,
    RULE_DET_ERROR_WARNING_ORDER_STABLE,
    RULE_DET_OUTPUT_NORMALIZATION_STABLE,
    RULE_DET_AMBIGUOUS_INTERPRETATION_REFUSED,
    RULE_DET_REPLAY_IDENTITY_STABLE,
]

PRE_DETERMINISM_RULE_ORDER = STRUCTURAL_RULE_ORDER + SEMANTIC_RULE_ORDER

RULE_BOUND_VALIDATION_RUNTIME_INDEPENDENT = "BOUND_VALIDATION_RUNTIME_INDEPENDENT"
RULE_BOUND_NO_TRIO_OR_ENGINE_APPLY_CALLS = "BOUND_NO_TRIO_OR_ENGINE_APPLY_CALLS"
RULE_BOUND_CAPTURED_PLAN_READ_ONLY = "BOUND_CAPTURED_PLAN_READ_ONLY"
RULE_BOUND_NO_RUNTIME_SIDE_EFFECT_INFERENCE = "BOUND_NO_RUNTIME_SIDE_EFFECT_INFERENCE"

BOUNDARY_RULE_ORDER = [
    RULE_BOUND_VALIDATION_RUNTIME_INDEPENDENT,
    RULE_BOUND_NO_TRIO_OR_ENGINE_APPLY_CALLS,
    RULE_BOUND_CAPTURED_PLAN_READ_ONLY,
    RULE_BOUND_NO_RUNTIME_SIDE_EFFECT_INFERENCE,
]

TERMINAL_SLOT_CLASSES = {"terminal_measure", "terminal_attribute"}

DEPENDENCY_REQUIREMENTS: dict[str, tuple[str, ...]] = {
    "entity": ("family", "operator"),
    "scope": ("entity",),
    "filter": ("entity", "scope", "filter"),
    "terminal_measure": ("entity", "scope", "filter"),
    "terminal_attribute": ("entity", "scope", "filter"),
    "formatter": ("terminal_measure", "terminal_attribute"),
}

KNOWN_FAMILY_BASELINES = {"info", "stats"}
KNOWN_OPERATOR_BASES = {
    "calendar",
    "conditional",
    "custom",
    "game_high",
    "games_with",
    "leader",
    "math",
    "previous",
    "rank",
    "streak",
}

DATE_ONLY_FORMATTERS = {
    "day_long",
    "day_short",
    "flex_long",
    "flex_short",
    "long_noyear",
    "long_year",
    "short_noyear",
    "short_year",
}

NUMERIC_FORMATTERS = {
    "*n",
    "+n",
    "-n",
    "/n",
    "ordinal",
}

TEXT_FORMATTERS = {
    "file_path",
    "lowercase",
    "smallcaps",
    "title",
    "uppercase",
}

DATE_LIKE_ATTRIBUTE_EXACT = {
    "birthdate",
    "day_of_week",
    "month",
    "pro_debut",
    "rookie_year",
    "season",
    "year",
}

DATE_LIKE_ATTRIBUTE_HINTS = (
    "date",
    "day",
    "month",
    "season",
    "year",
)

SCHEMA_SKIP_DETAIL = "Skipped because schema compatibility failed."
STRUCTURAL_SKIP_DETAIL = "Skipped because structural rules failed."
FORBIDDEN_RUNTIME_REFERENCE_MARKERS = [
    "smartstat_v4.0.0_beta.vbs",
    "smartstat_templateconfig.ini",
    "smartstat_mappings.ini",
    "smartstat_staticoverrides.ini",
    "smartstatvalidator.exe",
]
FORBIDDEN_TRIO_OR_ENGINE_CALL_NAME_MARKERS = [
    "triocmd",
    "engine_apply",
    "applyplan",
]
FORBIDDEN_TRIO_OR_ENGINE_LITERAL_MARKERS = [
    "page:get_property",
    "page:set_property",
    "script:run",
]


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


def make_rule_evaluation(
    rule_id: str,
    outcome: str,
    detail: str,
    category: str = "STRUCTURAL",
) -> dict[str, str]:
    return {
        "rule_id": rule_id,
        "category": category,
        "outcome": outcome,
        "detail": detail,
    }


def issue_sort_key(item: dict[str, Any]) -> tuple[str, str, str, str]:
    return (
        str(item.get("code", "")),
        str(item.get("rule_id", "")),
        str(item.get("message", "")),
        str(item.get("slot_order", "")),
    )


def sort_issues(issues: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return sorted(issues, key=issue_sort_key)


def only_expected_keys(rows: list[dict[str, Any]], expected_keys: set[str]) -> bool:
    for row in rows:
        if set(row.keys()) != expected_keys:
            return False
    return True


def dotted_name(node: ast.AST) -> str:
    if isinstance(node, ast.Name):
        return node.id
    if isinstance(node, ast.Attribute):
        parent = dotted_name(node.value)
        if parent:
            return f"{parent}.{node.attr}"
        return node.attr
    return ""


def string_literals(node: ast.AST) -> list[str]:
    literals: list[str] = []
    for child in ast.walk(node):
        if isinstance(child, ast.Constant) and isinstance(child.value, str):
            literals.append(child.value)
    return literals


def collect_call_sites(source_text: str) -> list[tuple[str, list[str]]]:
    tree = ast.parse(source_text)
    call_sites: list[tuple[str, list[str]]] = []

    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        call_name = dotted_name(node.func).lower()
        arg_literals: list[str] = []
        for arg in node.args:
            arg_literals.extend(string_literals(arg))
        for kwarg in node.keywords:
            if kwarg.value is not None:
                arg_literals.extend(string_literals(kwarg.value))
        call_sites.append((call_name, [text.lower() for text in arg_literals]))

    return call_sites


def build_normalized_plan_hash(payload: Any, raw_text: str) -> str:
    if payload is None:
        canonical_payload = raw_text
    else:
        canonical_payload = canonical_json(payload)

    digest_input = (HASH_PLACEHOLDER_CONTRACT + "\n" + canonical_payload).encode("utf-8")
    return hashlib.sha256(digest_input).hexdigest()


def make_replay_identity(input_artifact: str, normalized_plan_hash: str) -> str:
    digest_input = (
        REPLAY_IDENTITY_CONTRACT
        + "\n"
        + input_artifact.lower()
        + "\n"
        + normalized_plan_hash
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


def slot_order(slot: dict[str, Any]) -> int | None:
    order = slot.get("order")
    return order if isinstance(order, int) else None


def slot_class(slot: dict[str, Any]) -> str:
    return str(slot.get("slot_class", "")).strip()


def slot_token(slot: dict[str, Any]) -> str:
    return str(slot.get("token", "")).strip()


def ordered_slots(slots: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return sorted(
        slots,
        key=lambda item: (
            slot_order(item) if slot_order(item) is not None else 10**9,
            slot_class(item),
            slot_token(item),
        ),
    )


def terminal_orders(slots: list[dict[str, Any]]) -> list[int]:
    orders: list[int] = []
    for slot in slots:
        this_class = slot_class(slot)
        this_order = slot_order(slot)
        if this_class in TERMINAL_SLOT_CLASSES and this_order is not None:
            orders.append(this_order)
    return sorted(orders)


def first_slot_of_class(slots: list[dict[str, Any]], class_name: str) -> dict[str, Any] | None:
    matches = [slot for slot in slots if slot_class(slot) == class_name]
    if not matches:
        return None
    return ordered_slots(matches)[0]


def first_terminal_slot(slots: list[dict[str, Any]]) -> dict[str, Any] | None:
    terminals = [slot for slot in slots if slot_class(slot) in TERMINAL_SLOT_CLASSES]
    if not terminals:
        return None
    return ordered_slots(terminals)[0]


def token_base(token: str) -> str:
    normalized = str(token).strip().lower()
    if normalized.startswith("|"):
        normalized = normalized[1:].strip()
    if "(" in normalized:
        normalized = normalized.split("(", 1)[0].strip()
    return normalized


def operator_base_from_slot(slot: dict[str, Any] | None) -> str:
    if slot is None:
        return ""
    return token_base(slot_token(slot))


def entity_context_from_slot(slot: dict[str, Any] | None) -> str:
    if slot is None:
        return ""

    base = token_base(slot_token(slot))
    team_aliases = {"away", "home", "them", "us"}
    if base in team_aliases:
        return "team"
    return base


def formatter_category(token: str) -> str:
    normalized = token_base(token)

    if normalized in DATE_ONLY_FORMATTERS:
        return "date"

    if normalized in TEXT_FORMATTERS:
        return "text"

    if normalized in NUMERIC_FORMATTERS:
        return "numeric"

    if normalized.startswith("numeric_"):
        return "numeric"

    if normalized and normalized[0] in {"+", "-", "*", "/"}:
        return "numeric"

    return "unknown"


def is_date_like_attribute(attribute_token: str) -> bool:
    normalized = token_base(attribute_token)
    if normalized in DATE_LIKE_ATTRIBUTE_EXACT:
        return True
    return any(hint in normalized for hint in DATE_LIKE_ATTRIBUTE_HINTS)


def load_attribute_entity_compatibility(repo_root: Path) -> dict[str, set[str]]:
    path = repo_root / "docs" / "onair" / "attribute-dictionary.md"
    if not path.exists():
        return {}

    mapping: dict[str, set[str]] = {}
    for line in path.read_text(encoding="utf-8").splitlines():
        stripped = line.strip()
        if not stripped.startswith("|"):
            continue

        cells = [cell.strip() for cell in stripped.split("|")[1:-1]]
        if len(cells) < 4:
            continue

        attribute_name = cells[0].lower()
        entity_type = cells[1].lower()

        if (
            attribute_name in {"", "attribute_name"}
            or attribute_name.startswith("---")
            or entity_type.startswith("---")
        ):
            continue

        entities = [piece.strip() for piece in entity_type.split("/") if piece.strip()]
        if not entities:
            continue

        if attribute_name not in mapping:
            mapping[attribute_name] = set()

        for entity in entities:
            mapping[attribute_name].add(entity)

    return mapping


def skipped_structural_evaluations(detail: str = SCHEMA_SKIP_DETAIL) -> list[dict[str, str]]:
    return [
        make_rule_evaluation(
            rule_id=rule_id,
            outcome="WARN",
            detail=detail,
            category="STRUCTURAL",
        )
        for rule_id in STRUCTURAL_RULE_ORDER
    ]


def skipped_semantic_evaluations(detail: str) -> list[dict[str, str]]:
    return [
        make_rule_evaluation(
            rule_id=rule_id,
            outcome="WARN",
            detail=detail,
            category="SEMANTIC",
        )
        for rule_id in SEMANTIC_RULE_ORDER
    ]


def skipped_determinism_evaluations(detail: str) -> list[dict[str, str]]:
    return [
        make_rule_evaluation(
            rule_id=rule_id,
            outcome="WARN",
            detail=detail,
            category="DETERMINISM",
        )
        for rule_id in DETERMINISM_RULE_ORDER
    ]


def skipped_boundary_evaluations(detail: str) -> list[dict[str, str]]:
    return [
        make_rule_evaluation(
            rule_id=rule_id,
            outcome="WARN",
            detail=detail,
            category="BOUNDARY",
        )
        for rule_id in BOUNDARY_RULE_ORDER
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
    slot_classes = [slot_class(slot) for slot in slots]
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
            this_order = slot_order(slot)
            this_class = slot_class(slot)
            if this_order is not None and this_order > first_terminal and this_class != "formatter":
                offending_slots.append((this_order, this_class))

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
            slot_order(slot)
            for slot in slots
            if slot_class(slot) == "formatter" and slot_order(slot) is not None
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

def evaluate_sem_required_dependencies_present(
    slots: list[dict[str, Any]],
) -> tuple[dict[str, str], list[dict[str, Any]]]:
    seen: set[str] = set()
    violations: list[tuple[int | None, str, tuple[str, ...]]] = []

    for slot in ordered_slots(slots):
        this_class = slot_class(slot)
        required = DEPENDENCY_REQUIREMENTS.get(this_class)
        if required is not None:
            if not any(requirement in seen for requirement in required):
                violations.append((slot_order(slot), this_class, required))
        seen.add(this_class)

    if violations:
        violations.sort(
            key=lambda item: (
                item[0] if item[0] is not None else 10**9,
                item[1],
                ",".join(item[2]),
            )
        )
        first_order, first_class, first_required = violations[0]
        required_text = ", ".join(first_required)
        detail = (
            f"Missing required dependency for slot_class='{first_class}' at order={first_order}; "
            f"required_any_of=[{required_text}]."
        )
        evaluation = make_rule_evaluation(
            RULE_REQUIRED_DEPENDENCIES_PRESENT,
            "REFUSE",
            detail,
            category="SEMANTIC",
        )
        errors = [
            make_issue(
                "SEM_MISSING_REQUIRED_DEPENDENCY",
                detail,
                rule_id=RULE_REQUIRED_DEPENDENCIES_PRESENT,
                slot_order=first_order,
            )
        ]
        return (evaluation, errors)

    evaluation = make_rule_evaluation(
        RULE_REQUIRED_DEPENDENCIES_PRESENT,
        "PASS",
        "All dependent slot classes have required predecessors.",
        category="SEMANTIC",
    )
    return (evaluation, [])


def evaluate_sem_formatter_terminal_adjacent_and_type_compatible(
    slots: list[dict[str, Any]],
) -> tuple[dict[str, str], list[dict[str, Any]]]:
    formatter = first_slot_of_class(slots, "formatter")
    terminal = first_terminal_slot(slots)

    if formatter is None:
        evaluation = make_rule_evaluation(
            RULE_FORMATTER_TERMINAL_ADJACENT_AND_TYPE_COMPATIBLE,
            "PASS",
            "No formatter slot present.",
            category="SEMANTIC",
        )
        return (evaluation, [])

    formatter_order = slot_order(formatter)
    formatter_token = slot_token(formatter)

    if terminal is None:
        detail = "Formatter present but no terminal slot available for adjacency/type checks."
        evaluation = make_rule_evaluation(
            RULE_FORMATTER_TERMINAL_ADJACENT_AND_TYPE_COMPATIBLE,
            "REFUSE",
            detail,
            category="SEMANTIC",
        )
        errors = [
            make_issue(
                "SEM_FORMATTER_NOT_TERMINAL_ADJACENT",
                detail,
                rule_id=RULE_FORMATTER_TERMINAL_ADJACENT_AND_TYPE_COMPATIBLE,
                slot_order=formatter_order,
            )
        ]
        return (evaluation, errors)

    terminal_order = slot_order(terminal)
    terminal_class = slot_class(terminal)
    terminal_token = slot_token(terminal)

    if formatter_order is None or terminal_order is None or formatter_order != terminal_order + 1:
        detail = (
            f"Formatter at order={formatter_order} is not terminal-adjacent; "
            f"terminal_order={terminal_order}."
        )
        evaluation = make_rule_evaluation(
            RULE_FORMATTER_TERMINAL_ADJACENT_AND_TYPE_COMPATIBLE,
            "REFUSE",
            detail,
            category="SEMANTIC",
        )
        errors = [
            make_issue(
                "SEM_FORMATTER_NOT_TERMINAL_ADJACENT",
                detail,
                rule_id=RULE_FORMATTER_TERMINAL_ADJACENT_AND_TYPE_COMPATIBLE,
                slot_order=formatter_order,
            )
        ]
        return (evaluation, errors)

    formatter_kind = formatter_category(formatter_token)
    formatter_base = token_base(formatter_token)

    incompatible_reason = ""
    if terminal_class == "terminal_measure" and formatter_kind == "date":
        incompatible_reason = (
            f"Formatter '{formatter_base}' is date-only and incompatible with terminal_measure '{token_base(terminal_token)}'."
        )
    elif terminal_class == "terminal_attribute":
        if formatter_kind == "numeric":
            incompatible_reason = (
                f"Formatter '{formatter_base}' is numeric and incompatible with terminal_attribute '{token_base(terminal_token)}'."
            )
        elif formatter_kind == "date":
            terminal_attribute = token_base(terminal_token)
            if formatter_base in {"day_long", "day_short"}:
                if terminal_attribute != "day_of_week":
                    incompatible_reason = (
                        f"Formatter '{formatter_base}' requires terminal_attribute 'day_of_week'; actual='{terminal_attribute}'."
                    )
            elif not is_date_like_attribute(terminal_attribute):
                incompatible_reason = (
                    f"Formatter '{formatter_base}' requires date-like terminal_attribute; actual='{terminal_attribute}'."
                )

    if incompatible_reason:
        evaluation = make_rule_evaluation(
            RULE_FORMATTER_TERMINAL_ADJACENT_AND_TYPE_COMPATIBLE,
            "REFUSE",
            incompatible_reason,
            category="SEMANTIC",
        )
        errors = [
            make_issue(
                "SEM_FORMATTER_TYPE_INCOMPATIBLE",
                incompatible_reason,
                rule_id=RULE_FORMATTER_TERMINAL_ADJACENT_AND_TYPE_COMPATIBLE,
                slot_order=formatter_order,
            )
        ]
        return (evaluation, errors)

    detail = (
        f"Formatter is terminal-adjacent and type-compatible: formatter='{formatter_base}', "
        f"terminal_slot_class='{terminal_class}'."
    )
    if formatter_kind == "unknown":
        detail = (
            f"Formatter is terminal-adjacent; formatter token '{formatter_base}' has unknown type and is deferred "
            "to ambiguity rule."
        )

    evaluation = make_rule_evaluation(
        RULE_FORMATTER_TERMINAL_ADJACENT_AND_TYPE_COMPATIBLE,
        "PASS",
        detail,
        category="SEMANTIC",
    )
    return (evaluation, [])


def evaluate_sem_entity_family_operator_terminal_compatible(
    slots: list[dict[str, Any]],
    attribute_entity_compatibility: dict[str, set[str]],
) -> tuple[dict[str, str], list[dict[str, Any]]]:
    family = first_slot_of_class(slots, "family")
    operator = first_slot_of_class(slots, "operator")
    entity = first_slot_of_class(slots, "entity")
    terminal = first_terminal_slot(slots)

    issues: list[dict[str, Any]] = []

    if terminal is not None:
        terminal_cls = slot_class(terminal)
        terminal_order = slot_order(terminal)
        terminal_base = token_base(slot_token(terminal))

        family_base = token_base(slot_token(family)) if family is not None else ""
        if family_base == "stats" and terminal_cls != "terminal_measure":
            issues.append(
                make_issue(
                    "SEM_FAMILY_OPERATOR_CONTEXT_INCOMPATIBLE",
                    "Family 'stats' requires terminal_measure.",
                    rule_id=RULE_ENTITY_FAMILY_OPERATOR_TERMINAL_COMPATIBLE,
                    slot_order=terminal_order,
                )
            )
        if family_base == "info" and terminal_cls != "terminal_attribute":
            issues.append(
                make_issue(
                    "SEM_FAMILY_OPERATOR_CONTEXT_INCOMPATIBLE",
                    "Family 'info' requires terminal_attribute.",
                    rule_id=RULE_ENTITY_FAMILY_OPERATOR_TERMINAL_COMPATIBLE,
                    slot_order=terminal_order,
                )
            )

        operator_base = operator_base_from_slot(operator)
        if operator_base == "rank" and terminal_cls != "terminal_measure":
            issues.append(
                make_issue(
                    "SEM_FAMILY_OPERATOR_CONTEXT_INCOMPATIBLE",
                    "Operator 'rank' requires terminal_measure context.",
                    rule_id=RULE_ENTITY_FAMILY_OPERATOR_TERMINAL_COMPATIBLE,
                    slot_order=terminal_order,
                )
            )

        if operator_base == "previous":
            has_filter_before_terminal = False
            for slot in slots:
                if (
                    slot_class(slot) == "filter"
                    and slot_order(slot) is not None
                    and terminal_order is not None
                    and slot_order(slot) < terminal_order
                ):
                    has_filter_before_terminal = True
                    break

            if not has_filter_before_terminal:
                issues.append(
                    make_issue(
                        "SEM_FAMILY_OPERATOR_CONTEXT_INCOMPATIBLE",
                        "Operator 'previous' requires filter context before terminal output.",
                        rule_id=RULE_ENTITY_FAMILY_OPERATOR_TERMINAL_COMPATIBLE,
                        slot_order=terminal_order,
                    )
                )

        entity_context = entity_context_from_slot(entity)
        if entity_context:
            if terminal_cls == "terminal_measure" and entity_context == "time":
                issues.append(
                    make_issue(
                        "SEM_ENTITY_TERMINAL_INCOMPATIBLE",
                        "Entity 'time' is incompatible with terminal_measure; use info/attribute context.",
                        rule_id=RULE_ENTITY_FAMILY_OPERATOR_TERMINAL_COMPATIBLE,
                        slot_order=terminal_order,
                    )
                )
            elif terminal_cls == "terminal_attribute":
                allowed_entities = attribute_entity_compatibility.get(terminal_base)
                if allowed_entities and entity_context not in allowed_entities:
                    allowed_text = ", ".join(sorted(allowed_entities))
                    issues.append(
                        make_issue(
                            "SEM_ENTITY_TERMINAL_INCOMPATIBLE",
                            (
                                f"Entity '{entity_context}' is incompatible with terminal_attribute "
                                f"'{terminal_base}'. allowed=[{allowed_text}]."
                            ),
                            rule_id=RULE_ENTITY_FAMILY_OPERATOR_TERMINAL_COMPATIBLE,
                            slot_order=terminal_order,
                        )
                    )

    if issues:
        issues.sort(
            key=lambda item: (
                str(item.get("code", "")),
                str(item.get("message", "")),
                str(item.get("slot_order", "")),
            )
        )
        first = issues[0]
        evaluation = make_rule_evaluation(
            RULE_ENTITY_FAMILY_OPERATOR_TERMINAL_COMPATIBLE,
            "REFUSE",
            str(first["message"]),
            category="SEMANTIC",
        )
        return (evaluation, [first])

    evaluation = make_rule_evaluation(
        RULE_ENTITY_FAMILY_OPERATOR_TERMINAL_COMPATIBLE,
        "PASS",
        "Entity/family/operator context is compatible with terminal kind.",
        category="SEMANTIC",
    )
    return (evaluation, [])

def evaluate_sem_ambiguous_or_incompatible_combination_refused(
    payload: Any,
    slots: list[dict[str, Any]],
    attribute_entity_compatibility: dict[str, set[str]],
) -> tuple[dict[str, str], list[dict[str, Any]]]:
    reasons: list[tuple[int | None, str]] = []

    family_slots = [slot for slot in slots if slot_class(slot) == "family"]
    operator_slots = [slot for slot in slots if slot_class(slot) == "operator"]
    entity_slots = [slot for slot in slots if slot_class(slot) == "entity"]
    formatter_slot = first_slot_of_class(slots, "formatter")
    terminal_slot = first_terminal_slot(slots)

    if len(family_slots) > 1:
        reason_slot = ordered_slots(family_slots)[1]
        reasons.append((slot_order(reason_slot), "Multiple family slots create ambiguous semantic context."))

    if len(operator_slots) > 1:
        reason_slot = ordered_slots(operator_slots)[1]
        reasons.append((slot_order(reason_slot), "Multiple operator slots create ambiguous semantic context."))

    if len(entity_slots) > 1:
        reason_slot = ordered_slots(entity_slots)[1]
        reasons.append((slot_order(reason_slot), "Multiple entity slots create ambiguous semantic context."))

    family = first_slot_of_class(slots, "family")
    operator = first_slot_of_class(slots, "operator")

    family_base = token_base(slot_token(family)) if family is not None else ""
    if family_base and family_base not in KNOWN_FAMILY_BASELINES and operator is None:
        reasons.append(
            (
                slot_order(family),
                f"Family '{family_base}' has no declared baseline semantic policy in this validator.",
            )
        )

    operator_base = operator_base_from_slot(operator)
    if operator_base and operator_base not in KNOWN_OPERATOR_BASES:
        reasons.append(
            (
                slot_order(operator),
                f"Operator '{operator_base}' has no declared semantic policy in this validator.",
            )
        )

    if formatter_slot is not None and formatter_category(slot_token(formatter_slot)) == "unknown":
        formatter_base = token_base(slot_token(formatter_slot))
        reasons.append(
            (
                slot_order(formatter_slot),
                f"Formatter '{formatter_base}' has unknown semantic type compatibility.",
            )
        )

    entity = first_slot_of_class(slots, "entity")
    if terminal_slot is not None and entity is not None and slot_class(terminal_slot) == "terminal_attribute":
        attribute_base = token_base(slot_token(terminal_slot))
        if attribute_base not in attribute_entity_compatibility:
            reasons.append(
                (
                    slot_order(terminal_slot),
                    f"Terminal attribute '{attribute_base}' has no entity-compatibility mapping evidence.",
                )
            )

    terminal_definition = payload.get("terminal") if isinstance(payload, dict) else None
    if isinstance(terminal_definition, dict) and terminal_slot is not None:
        definition_slot_class = str(terminal_definition.get("slot_class", "")).strip()
        definition_token = token_base(str(terminal_definition.get("token", "")))
        slot_slot_class = slot_class(terminal_slot)
        slot_token_base = token_base(slot_token(terminal_slot))
        if definition_slot_class != slot_slot_class or definition_token != slot_token_base:
            reasons.append(
                (
                    slot_order(terminal_slot),
                    (
                        "Terminal definition does not match resolved terminal slot "
                        f"(definition={definition_slot_class}:{definition_token}, "
                        f"slot={slot_slot_class}:{slot_token_base})."
                    ),
                )
            )

    if reasons:
        reasons.sort(key=lambda item: (item[0] if item[0] is not None else 10**9, item[1]))
        refusal_order, refusal_message = reasons[0]
        evaluation = make_rule_evaluation(
            RULE_AMBIGUOUS_OR_INCOMPATIBLE_COMBINATION_REFUSED,
            "REFUSE",
            refusal_message,
            category="SEMANTIC",
        )
        errors = [
            make_issue(
                "SEM_AMBIGUOUS_SEMANTIC_COMBINATION",
                refusal_message,
                rule_id=RULE_AMBIGUOUS_OR_INCOMPATIBLE_COMBINATION_REFUSED,
                slot_order=refusal_order,
            )
        ]
        return (evaluation, errors)

    evaluation = make_rule_evaluation(
        RULE_AMBIGUOUS_OR_INCOMPATIBLE_COMBINATION_REFUSED,
        "PASS",
        "No ambiguous semantic combinations detected.",
        category="SEMANTIC",
    )
    return (evaluation, [])


def evaluate_semantic_rules(
    payload: Any,
    attribute_entity_compatibility: dict[str, set[str]],
) -> tuple[list[dict[str, str]], list[dict[str, Any]]]:
    slots = get_slots(payload)

    evaluations: list[dict[str, str]] = []
    errors: list[dict[str, Any]] = []

    dependency_eval, dependency_errors = evaluate_sem_required_dependencies_present(slots)
    evaluations.append(dependency_eval)
    errors.extend(dependency_errors)

    formatter_eval, formatter_errors = evaluate_sem_formatter_terminal_adjacent_and_type_compatible(slots)
    evaluations.append(formatter_eval)
    errors.extend(formatter_errors)

    compatibility_eval, compatibility_errors = evaluate_sem_entity_family_operator_terminal_compatible(
        slots,
        attribute_entity_compatibility,
    )
    evaluations.append(compatibility_eval)
    errors.extend(compatibility_errors)

    ambiguity_eval, ambiguity_errors = evaluate_sem_ambiguous_or_incompatible_combination_refused(
        payload,
        slots,
        attribute_entity_compatibility,
    )
    evaluations.append(ambiguity_eval)
    errors.extend(ambiguity_errors)

    return (evaluations, errors)


def evaluate_determinism_rules(
    input_artifact: str,
    payload: Any,
    raw_text: str,
    normalized_plan_hash: str,
    errors: list[dict[str, Any]],
    warnings: list[dict[str, Any]],
    rule_evaluations_before_determinism: list[dict[str, str]],
) -> tuple[list[dict[str, str]], list[dict[str, Any]], list[dict[str, Any]]]:
    evaluations: list[dict[str, str]] = []
    det_errors: list[dict[str, Any]] = []
    det_warnings: list[dict[str, Any]] = []

    # 1) Deterministic rule evaluation ordering check.
    expected_rule_order = PRE_DETERMINISM_RULE_ORDER
    actual_rule_order = [str(row.get("rule_id", "")) for row in rule_evaluations_before_determinism]
    if actual_rule_order == expected_rule_order:
        evaluations.append(
            make_rule_evaluation(
                rule_id=RULE_DET_RULE_EVALUATION_ORDER_STABLE,
                outcome="PASS",
                detail="Rule evaluation order is stable for schema/structural/semantic phases.",
                category="DETERMINISM",
            )
        )
    else:
        detail = (
            "Rule evaluation order is unstable. "
            f"expected={expected_rule_order} actual={actual_rule_order}."
        )
        evaluations.append(
            make_rule_evaluation(
                rule_id=RULE_DET_RULE_EVALUATION_ORDER_STABLE,
                outcome="REFUSE",
                detail=detail,
                category="DETERMINISM",
            )
        )
        det_errors.append(
            make_issue(
                code="DET_RULE_ORDER_UNSTABLE",
                message=detail,
                rule_id=RULE_DET_RULE_EVALUATION_ORDER_STABLE,
                slot_order=None,
            )
        )

    # 2) Deterministic error/warning ordering check.
    order_issues: list[str] = []
    if errors != sort_issues(errors):
        detail = "Error ordering is not deterministic."
        det_errors.append(
            make_issue(
                code="DET_ERROR_ORDER_UNSTABLE",
                message=detail,
                rule_id=RULE_DET_ERROR_WARNING_ORDER_STABLE,
                slot_order=None,
            )
        )
        order_issues.append(detail)
    if warnings != sort_issues(warnings):
        detail = "Warning ordering is not deterministic."
        det_errors.append(
            make_issue(
                code="DET_WARNING_ORDER_UNSTABLE",
                message=detail,
                rule_id=RULE_DET_ERROR_WARNING_ORDER_STABLE,
                slot_order=None,
            )
        )
        order_issues.append(detail)

    if order_issues:
        evaluations.append(
            make_rule_evaluation(
                rule_id=RULE_DET_ERROR_WARNING_ORDER_STABLE,
                outcome="REFUSE",
                detail="; ".join(order_issues),
                category="DETERMINISM",
            )
        )
    else:
        evaluations.append(
            make_rule_evaluation(
                rule_id=RULE_DET_ERROR_WARNING_ORDER_STABLE,
                outcome="PASS",
                detail="Error and warning ordering is stable.",
                category="DETERMINISM",
            )
        )

    # 3) Stable output normalization check.
    recomputed_hash = build_normalized_plan_hash(payload, raw_text)
    if recomputed_hash == normalized_plan_hash:
        evaluations.append(
            make_rule_evaluation(
                rule_id=RULE_DET_OUTPUT_NORMALIZATION_STABLE,
                outcome="PASS",
                detail="normalized_plan_hash is stable under canonical normalization.",
                category="DETERMINISM",
            )
        )
    else:
        detail = (
            "normalized_plan_hash is unstable under canonical normalization. "
            f"expected={normalized_plan_hash} actual={recomputed_hash}."
        )
        evaluations.append(
            make_rule_evaluation(
                rule_id=RULE_DET_OUTPUT_NORMALIZATION_STABLE,
                outcome="REFUSE",
                detail=detail,
                category="DETERMINISM",
            )
        )
        det_errors.append(
            make_issue(
                code="DET_OUTPUT_NORMALIZATION_UNSTABLE",
                message=detail,
                rule_id=RULE_DET_OUTPUT_NORMALIZATION_STABLE,
                slot_order=None,
            )
        )

    # 4) Ambiguous interpretation fail-closed check.
    ambiguity_errors = [
        issue for issue in errors if str(issue.get("code", "")) == "SEM_AMBIGUOUS_SEMANTIC_COMBINATION"
    ]
    semantic_ambiguity_eval = next(
        (
            row
            for row in rule_evaluations_before_determinism
            if str(row.get("rule_id", "")) == RULE_AMBIGUOUS_OR_INCOMPATIBLE_COMBINATION_REFUSED
        ),
        None,
    )

    if ambiguity_errors:
        if semantic_ambiguity_eval is not None and str(semantic_ambiguity_eval.get("outcome", "")) == "REFUSE":
            evaluations.append(
                make_rule_evaluation(
                    rule_id=RULE_DET_AMBIGUOUS_INTERPRETATION_REFUSED,
                    outcome="PASS",
                    detail="Ambiguous interpretation was refused by semantic layer.",
                    category="DETERMINISM",
                )
            )
        else:
            detail = "Ambiguous semantic combination was not fail-closed by semantic layer."
            evaluations.append(
                make_rule_evaluation(
                    rule_id=RULE_DET_AMBIGUOUS_INTERPRETATION_REFUSED,
                    outcome="REFUSE",
                    detail=detail,
                    category="DETERMINISM",
                )
            )
            det_errors.append(
                make_issue(
                    code="DET_AMBIGUOUS_INTERPRETATION_REFUSED",
                    message=detail,
                    rule_id=RULE_DET_AMBIGUOUS_INTERPRETATION_REFUSED,
                    slot_order=None,
                )
            )
    else:
        evaluations.append(
            make_rule_evaluation(
                rule_id=RULE_DET_AMBIGUOUS_INTERPRETATION_REFUSED,
                outcome="PASS",
                detail="No ambiguous semantic interpretation detected in this artifact.",
                category="DETERMINISM",
            )
        )

    # 5) Deterministic replay identity representation check.
    replay_identity = make_replay_identity(input_artifact, normalized_plan_hash)
    replay_identity_recomputed = make_replay_identity(input_artifact, normalized_plan_hash)
    if replay_identity == replay_identity_recomputed:
        evaluations.append(
            make_rule_evaluation(
                rule_id=RULE_DET_REPLAY_IDENTITY_STABLE,
                outcome="PASS",
                detail=f"replay_identity={replay_identity}",
                category="DETERMINISM",
            )
        )
    else:
        detail = (
            "Replay identity computation is unstable. "
            f"expected={replay_identity} actual={replay_identity_recomputed}."
        )
        evaluations.append(
            make_rule_evaluation(
                rule_id=RULE_DET_REPLAY_IDENTITY_STABLE,
                outcome="REFUSE",
                detail=detail,
                category="DETERMINISM",
            )
        )
        det_errors.append(
            make_issue(
                code="DET_REPLAY_IDENTITY_UNSTABLE",
                message=detail,
                rule_id=RULE_DET_REPLAY_IDENTITY_STABLE,
                slot_order=None,
            )
        )

    return (evaluations, det_errors, det_warnings)


def evaluate_boundary_rules(
    artifact_path: Path,
    raw_text: str,
    errors: list[dict[str, Any]],
    warnings: list[dict[str, Any]],
    rule_evaluations_before_boundary: list[dict[str, str]],
) -> tuple[list[dict[str, str]], list[dict[str, Any]], list[dict[str, Any]]]:
    evaluations: list[dict[str, str]] = []
    boundary_errors: list[dict[str, Any]] = []
    boundary_warnings: list[dict[str, Any]] = []

    validator_source_text = Path(__file__).read_text(encoding="utf-8")
    source_analysis_error = ""
    call_sites: list[tuple[str, list[str]]] = []

    try:
        call_sites = collect_call_sites(validator_source_text)
    except SyntaxError as exc:
        source_analysis_error = (
            "Unable to parse validator source for boundary analysis. "
            f"{exc.__class__.__name__}: {exc}."
        )

    # 1) Validation runtime-independence guard.
    runtime_reference_hits: list[str] = []
    if not source_analysis_error:
        for call_name, arg_literals in call_sites:
            for literal in arg_literals:
                for marker in FORBIDDEN_RUNTIME_REFERENCE_MARKERS:
                    if marker in literal:
                        runtime_reference_hits.append(f"{call_name}:{marker}")
    runtime_reference_hits = sorted(set(runtime_reference_hits))

    if source_analysis_error:
        detail = source_analysis_error
        evaluations.append(
            make_rule_evaluation(
                rule_id=RULE_BOUND_VALIDATION_RUNTIME_INDEPENDENT,
                outcome="REFUSE",
                detail=detail,
                category="BOUNDARY",
            )
        )
        boundary_errors.append(
            make_issue(
                code="BOUND_RUNTIME_BEHAVIOR_DETECTED",
                message=detail,
                rule_id=RULE_BOUND_VALIDATION_RUNTIME_INDEPENDENT,
                slot_order=None,
            )
        )
    elif runtime_reference_hits:
        detail = (
            "Runtime artifact references detected in validator call sites: "
            f"{runtime_reference_hits}."
        )
        evaluations.append(
            make_rule_evaluation(
                rule_id=RULE_BOUND_VALIDATION_RUNTIME_INDEPENDENT,
                outcome="REFUSE",
                detail=detail,
                category="BOUNDARY",
            )
        )
        boundary_errors.append(
            make_issue(
                code="BOUND_RUNTIME_BEHAVIOR_DETECTED",
                message=detail,
                rule_id=RULE_BOUND_VALIDATION_RUNTIME_INDEPENDENT,
                slot_order=None,
            )
        )
    else:
        evaluations.append(
            make_rule_evaluation(
                rule_id=RULE_BOUND_VALIDATION_RUNTIME_INDEPENDENT,
                outcome="PASS",
                detail="Validator call sites contain no SmartStat runtime artifact references.",
                category="BOUNDARY",
            )
        )

    # 2) No Trio/apply call surface guard.
    trio_or_engine_hits: list[str] = []
    if not source_analysis_error:
        for call_name, arg_literals in call_sites:
            for marker in FORBIDDEN_TRIO_OR_ENGINE_CALL_NAME_MARKERS:
                if marker in call_name:
                    trio_or_engine_hits.append(f"call:{call_name}")
            for literal in arg_literals:
                for marker in FORBIDDEN_TRIO_OR_ENGINE_LITERAL_MARKERS:
                    if marker in literal:
                        trio_or_engine_hits.append(f"literal:{marker}")
    trio_or_engine_hits = sorted(set(trio_or_engine_hits))

    if source_analysis_error:
        detail = source_analysis_error
        evaluations.append(
            make_rule_evaluation(
                rule_id=RULE_BOUND_NO_TRIO_OR_ENGINE_APPLY_CALLS,
                outcome="REFUSE",
                detail=detail,
                category="BOUNDARY",
            )
        )
        boundary_errors.append(
            make_issue(
                code="BOUND_TRIO_OR_ENGINE_CALL_DETECTED",
                message=detail,
                rule_id=RULE_BOUND_NO_TRIO_OR_ENGINE_APPLY_CALLS,
                slot_order=None,
            )
        )
    elif trio_or_engine_hits:
        detail = (
            "Trio/apply call surfaces detected in validator call sites: "
            f"{trio_or_engine_hits}."
        )
        evaluations.append(
            make_rule_evaluation(
                rule_id=RULE_BOUND_NO_TRIO_OR_ENGINE_APPLY_CALLS,
                outcome="REFUSE",
                detail=detail,
                category="BOUNDARY",
            )
        )
        boundary_errors.append(
            make_issue(
                code="BOUND_TRIO_OR_ENGINE_CALL_DETECTED",
                message=detail,
                rule_id=RULE_BOUND_NO_TRIO_OR_ENGINE_APPLY_CALLS,
                slot_order=None,
            )
        )
    else:
        evaluations.append(
            make_rule_evaluation(
                rule_id=RULE_BOUND_NO_TRIO_OR_ENGINE_APPLY_CALLS,
                outcome="PASS",
                detail="Validator call sites contain no Trio integration or engine/apply surfaces.",
                category="BOUNDARY",
            )
        )

    # 3) Captured-plan read-only guarantee.
    raw_text_after = artifact_path.read_text(encoding="utf-8")
    if raw_text_after != raw_text:
        detail = (
            "Input artifact changed during validation; runner must remain read-only."
        )
        evaluations.append(
            make_rule_evaluation(
                rule_id=RULE_BOUND_CAPTURED_PLAN_READ_ONLY,
                outcome="REFUSE",
                detail=detail,
                category="BOUNDARY",
            )
        )
        boundary_errors.append(
            make_issue(
                code="BOUND_CAPTURED_PLAN_MUTATION_DETECTED",
                message=detail,
                rule_id=RULE_BOUND_CAPTURED_PLAN_READ_ONLY,
                slot_order=None,
            )
        )
    else:
        evaluations.append(
            make_rule_evaluation(
                rule_id=RULE_BOUND_CAPTURED_PLAN_READ_ONLY,
                outcome="PASS",
                detail="Input artifact remained byte-identical before/after validation.",
                category="BOUNDARY",
            )
        )

    # 4) No runtime side-effect inference in validator output contract.
    valid_error_shape = only_expected_keys(errors, {"code", "message", "rule_id", "slot_order"})
    valid_warning_shape = only_expected_keys(warnings, {"code", "message", "rule_id", "slot_order"})
    valid_rule_eval_shape = only_expected_keys(
        rule_evaluations_before_boundary,
        {"rule_id", "category", "outcome", "detail"},
    )
    valid_categories = all(
        str(item.get("category", "")) in {"STRUCTURAL", "SEMANTIC", "DETERMINISM"}
        for item in rule_evaluations_before_boundary
    )

    if valid_error_shape and valid_warning_shape and valid_rule_eval_shape and valid_categories:
        evaluations.append(
            make_rule_evaluation(
                rule_id=RULE_BOUND_NO_RUNTIME_SIDE_EFFECT_INFERENCE,
                outcome="PASS",
                detail="Output contract remains validation-only with no runtime side-effect fields inferred.",
                category="BOUNDARY",
            )
        )
    else:
        detail = (
            "Output contract shape/category drift detected; runtime side-effect inference risk."
        )
        evaluations.append(
            make_rule_evaluation(
                rule_id=RULE_BOUND_NO_RUNTIME_SIDE_EFFECT_INFERENCE,
                outcome="REFUSE",
                detail=detail,
                category="BOUNDARY",
            )
        )
        boundary_errors.append(
            make_issue(
                code="BOUND_RUNTIME_SIDE_EFFECT_INFERENCE_DETECTED",
                message=detail,
                rule_id=RULE_BOUND_NO_RUNTIME_SIDE_EFFECT_INFERENCE,
                slot_order=None,
            )
        )

    return (evaluations, boundary_errors, boundary_warnings)


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="WP-18 validator runner (schema + structural + semantic + determinism + boundary rules)."
    )
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
    artifact_path: Path,
    schema: dict[str, Any],
    repo_root: Path,
    attribute_entity_compatibility: dict[str, set[str]],
) -> dict[str, Any]:
    input_artifact = rel_path(artifact_path, repo_root)
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

    normalized_plan_hash = build_normalized_plan_hash(payload, raw_text)

    if errors:
        structural_evals = skipped_structural_evaluations(SCHEMA_SKIP_DETAIL)
        semantic_evals = skipped_semantic_evaluations(SCHEMA_SKIP_DETAIL)
        determinism_evals = skipped_determinism_evaluations(SCHEMA_SKIP_DETAIL)
        boundary_evals = skipped_boundary_evaluations(SCHEMA_SKIP_DETAIL)
        rule_evaluations = structural_evals + semantic_evals + determinism_evals + boundary_evals
    else:
        structural_evals, structural_errors = evaluate_structural_rules(payload)
        errors.extend(structural_errors)

        if structural_errors:
            semantic_evals = skipped_semantic_evaluations(STRUCTURAL_SKIP_DETAIL)
            determinism_evals = skipped_determinism_evaluations(STRUCTURAL_SKIP_DETAIL)
            boundary_evals = skipped_boundary_evaluations(STRUCTURAL_SKIP_DETAIL)
            rule_evaluations = structural_evals + semantic_evals + determinism_evals + boundary_evals
        else:
            semantic_evals, semantic_errors = evaluate_semantic_rules(
                payload,
                attribute_entity_compatibility,
            )
            errors.extend(semantic_errors)
            pre_determinism_rule_evals = structural_evals + semantic_evals

            # Determinism assertions evaluate only after schema + structural + semantic phases.
            errors = sort_issues(errors)
            warnings = sort_issues(warnings)
            determinism_evals, determinism_errors, determinism_warnings = evaluate_determinism_rules(
                input_artifact=input_artifact,
                payload=payload,
                raw_text=raw_text,
                normalized_plan_hash=normalized_plan_hash,
                errors=errors,
                warnings=warnings,
                rule_evaluations_before_determinism=pre_determinism_rule_evals,
            )
            errors.extend(determinism_errors)
            warnings.extend(determinism_warnings)
            pre_boundary_rule_evals = pre_determinism_rule_evals + determinism_evals

            # Boundary assertions evaluate only after schema + structural + semantic + determinism phases.
            errors = sort_issues(errors)
            warnings = sort_issues(warnings)
            boundary_evals, boundary_errors, boundary_warnings = evaluate_boundary_rules(
                artifact_path=artifact_path,
                raw_text=raw_text,
                errors=errors,
                warnings=warnings,
                rule_evaluations_before_boundary=pre_boundary_rule_evals,
            )
            errors.extend(boundary_errors)
            warnings.extend(boundary_warnings)
            rule_evaluations = pre_boundary_rule_evals + boundary_evals

    errors = sort_issues(errors)
    warnings = sort_issues(warnings)

    validation_result = {
        "status": "PASS" if not errors else "REFUSE",
        "errors": errors,
        "warnings": warnings,
        "normalized_plan_hash": normalized_plan_hash,
        "rule_evaluations": rule_evaluations,
    }

    return {
        "input_artifact": input_artifact,
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
    attribute_entity_compatibility = load_attribute_entity_compatibility(repo_root)
    artifact_paths = collect_artifact_paths(input_path)

    results = [
        process_artifact(path, schema, repo_root, attribute_entity_compatibility)
        for path in artifact_paths
    ]
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
