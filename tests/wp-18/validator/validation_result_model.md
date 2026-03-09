# WP-18 Validation Result Model (Target-01 Scaffolding)

Reference contract:
- `docs/onair/plan-validation-contract.md`

This document captures the Target-01 output model emitted by:
- `tests/wp-18/validator/validator_runner.py`

Target-01 scope:
- schema compatibility checks only
- deterministic output ordering
- no validation rule implementation yet

## Output Shape

```yaml
validation_result:
  status: PASS | REFUSE
  errors: []
  warnings: []
  normalized_plan_hash: string
  rule_evaluations: []
```

Notes:
- `rule_evaluations` is intentionally empty in Target-01.
- `status` is currently driven by JSON parse + schema compatibility only.
- `normalized_plan_hash` is a deterministic placeholder hash for replay stability.
- Future WP-18 targets will populate rule results without changing runtime behavior.

