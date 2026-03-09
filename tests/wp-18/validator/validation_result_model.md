# WP-18 Validation Result Model (Target-01 Scaffolding)

Reference contract:
- `docs/onair/plan-validation-contract.md`

This document captures the Target-01 output model emitted by:
- `tests/wp-18/validator/validator_runner.py`

Target-02 scope:
- schema compatibility checks
- structural rule evaluation only
- deterministic output ordering
- no semantic rule implementation yet

## Output Shape

```yaml
validation_result:
  status: PASS | REFUSE
  errors: []
  warnings: []
  normalized_plan_hash: string
  rule_evaluations:
    - rule_id: string
      category: STRUCTURAL
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
- `status` is driven by schema + structural failures.
- `normalized_plan_hash` is a deterministic placeholder hash for replay stability.
- Future WP-18 targets may add semantic/boundary determinism policy checks without changing runtime behavior.
