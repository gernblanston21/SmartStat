# OnAir Plan Validation Contract (WP-18 Kickoff Gate)

Contract version: `wp18.plan_validation.v1`  
Status: `eligible to start; documentation/governance only; validator implementation not started`  
Last updated: `2026-03-09`

## 1. Purpose

WP-18 defines the read-only Plan Validation Contract Layer between:

- WP-17 plan capture artifacts
- future layers (WP-19 viewer, WP-20 runtime bridge, both deferred)

Validation answers one question:

`Is this captured plan valid according to SmartStat architecture rules?`

This layer does not execute runtime behavior and does not apply stats to graphics.

## 2. Inputs and Contract Surface

Validation input is a WP-17 artifact that already conforms to:

- `docs/onair/plan-capture-contract.md`
- `docs/onair/plan-capture.schema.json`

Validation is read-only over that captured artifact and must not mutate it.

## 3. What Counts as a Validated Plan

A captured plan is validated only when all mandatory rules pass:

- structural integrity of slot chain
- semantic compatibility of slot relationships
- deterministic interpretation (no ambiguous branch accepted)
- boundary compliance (validation-only behavior)

If any mandatory rule fails, output status is `REFUSE` (fail closed).

## 4. Deterministic Evaluation Procedure

Validation must execute a fixed, deterministic rule order:

1. Input compatibility checks (WP-17 contract/version presence)
2. Structural rule set
3. Semantic rule set
4. Determinism rule set
5. Boundary rule set
6. Status reduction (`PASS` or `REFUSE`)
7. Stable `normalized_plan_hash` generation

Determinism requirements:

- same input artifact + same validation contract version -> same status, errors, warnings, and hash
- rule evaluations emitted in stable, predeclared order
- no environment-dependent interpretation

## 5. Validation Rule Taxonomy

### 5.1 Structural Rules

- Slot ordering is contiguous and architecturally legal.
- Terminal slot is present for `captured_plan`.
- Illegal slot combinations are refused.
- No non-formatter slots may appear after terminal.
- Formatter count is at most one.

### 5.2 Semantic Rules

- Slot compatibility matches architecture model.
- Required dependencies are present before dependent slots.
- Formatter placement is terminal-adjacent and type-compatible.
- Entity/family/operator context and terminal kind are compatible.

### 5.3 Determinism Rules

- Stable ordering is enforced for all rule evaluations.
- Deterministic replay produces identical validation output.
- Ambiguous interpretation is refused; no best-effort branch selection.
- Canonical hash input normalization is versioned and stable.

### 5.4 Boundary Rules

- Validation must not execute runtime behavior.
- Validation cannot mutate captured plans.
- Validation cannot call SmartStat engine apply paths.
- Validation cannot infer runtime side effects from architecture-only artifacts.

## 6. Refusal Conditions (Fail Closed)

`status=REFUSE` is mandatory when any of these occur:

- incompatible or unsupported capture contract version
- malformed slot order or missing terminal requirements
- illegal slot compatibility/dependency chain
- formatter placement/type incompatibility
- ambiguous interpretation requiring non-deterministic branch choice
- attempted boundary violation (runtime/apply behavior, mutation, integration side effect)

## 7. Validation Result Model (Normative Shape)

Validation output contract:

```yaml
validation_result:
  status: PASS | REFUSE
  errors:
    - code: string
      message: string
      rule_id: string
      slot_order: integer|null
  warnings:
    - code: string
      message: string
      rule_id: string
      slot_order: integer|null
  normalized_plan_hash: string   # sha256, stable under canonicalization contract
  rule_evaluations:
    - rule_id: string
      category: STRUCTURAL | SEMANTIC | DETERMINISM | BOUNDARY
      outcome: PASS | REFUSE | WARN
      detail: string
```

Status reduction rule:

- `PASS` only if no refusing rule is triggered.
- `REFUSE` if one or more refusing rules are triggered.

## 8. WP-18 Harness Expectations (Not Implemented in Kickoff Gate)

Planned harness root:

- `tests/wp-18/`

Expected elements (future implementation work, not in this pass):

- validator runner
- good/bad plan fixtures
- deterministic replay checks
- refusal-case tests

Evidence expectations (future WP-18 implementation):

- run summary with pass/refuse counts
- deterministic replay hash comparison output
- explicit refusal diagnostics for failing fixtures

## 9. Explicit WP-18 Boundaries

Allowed:

- validating captured plans
- deterministic rule evaluation
- validator tooling
- schema compatibility checks

Not allowed:

- runtime execution
- Trio integration
- SmartStat engine calls
- applying stats to graphics
- modifying captured plan artifacts

## 10. Sequencing and Deferred Work

WP-18 kickoff gate marks the validation layer as eligible to start implementation work.

Still deferred:

- WP-19 plan viewer
- WP-20 runtime bridge (requires explicit approval before implementation)
