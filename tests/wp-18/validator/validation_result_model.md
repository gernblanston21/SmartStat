# WP-18 Validation Result Model (Target-05 Boundary Layer)

Reference contract:
- `docs/onair/plan-validation-contract.md`

This document captures the Target-01 output model emitted by:
- `tests/wp-18/validator/validator_runner.py`

Target-05 scope:
- schema compatibility checks
- structural rule evaluation
- semantic rule evaluation
- determinism rule evaluation
- boundary rule evaluation
- deterministic output ordering

## Output Shape

```yaml
validation_result:
  status: PASS | REFUSE
  errors: []
  warnings: []
  normalized_plan_hash: string
  rule_evaluations:
    - rule_id: string
      category: STRUCTURAL | SEMANTIC | DETERMINISM | BOUNDARY
      outcome: PASS | REFUSE | WARN
      detail: string
```

Notes:
- Structural rule IDs are emitted in stable declared order:
  - `STRUCT_SLOT_ORDER_CONTIGUOUS_ASC`
  - `STRUCT_CAPTURED_PLAN_TERMINAL_REQUIRED`
  - `STRUCT_ILLEGAL_SLOT_COMBINATION`
  - `STRUCT_NO_NON_FORMATTER_AFTER_TERMINAL`
  - `STRUCT_FORMATTER_COUNT_MAX_ONE`
- Semantic rule IDs are emitted after structural rules in stable declared order:
  - `SEM_REQUIRED_DEPENDENCIES_PRESENT`
  - `SEM_FORMATTER_TERMINAL_ADJACENT_AND_TYPE_COMPATIBLE`
  - `SEM_ENTITY_FAMILY_OPERATOR_TERMINAL_COMPATIBLE`
  - `SEM_AMBIGUOUS_OR_INCOMPATIBLE_COMBINATION_REFUSED`
- Determinism rule IDs are emitted after semantic rules in stable declared order:
  - `DET_RULE_EVALUATION_ORDER_STABLE`
  - `DET_ERROR_WARNING_ORDER_STABLE`
  - `DET_OUTPUT_NORMALIZATION_STABLE`
  - `DET_AMBIGUOUS_INTERPRETATION_REFUSED`
  - `DET_REPLAY_IDENTITY_STABLE`
- Boundary rule IDs are emitted after determinism rules in stable declared order:
  - `BOUND_VALIDATION_RUNTIME_INDEPENDENT`
  - `BOUND_NO_TRIO_OR_ENGINE_APPLY_CALLS`
  - `BOUND_CAPTURED_PLAN_READ_ONLY`
  - `BOUND_NO_RUNTIME_SIDE_EFFECT_INFERENCE`
- `status` is driven by schema + structural + semantic + determinism + boundary failures.
- If schema compatibility fails, structural + semantic + determinism + boundary evaluations are emitted as deterministic `WARN`.
- If structural rules fail, semantic + determinism + boundary evaluations are emitted as deterministic `WARN`.
- If semantic rules fail, determinism + boundary rules still evaluate where meaningful.
- If determinism rules fail, boundary rules still evaluate where meaningful.
- `normalized_plan_hash` is a deterministic placeholder hash for replay stability.
