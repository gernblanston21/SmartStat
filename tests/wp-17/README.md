# WP-17 Plan Capture Contract Harness

This directory contains docs/tests/tooling-only validation artifacts for WP-17 plan capture contract enforcement.

## Scope

- Validate versioned captured-plan JSON artifacts (`wp17.plan_capture.v1`)
- Enforce fail-closed capture/refusal boundaries
- Prove deterministic canonical hash stability via replay

## Runner

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File tests/wp-17/run_wp17.ps1 -RunLabel <label>
```

## Validator

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File tests/wp-17/contract-validators/validate_plan_capture_contract.ps1 -Path <artifact.json> -RunLabel <label> -CaseLabel <case>
```

## Contract Inputs

- `docs/onair/plan-capture-contract.md`
- `docs/onair/plan-capture.schema.json`

## Artifacts

- `tests/wp-17/contract-validators/artifacts/<runLabel>/run_summary.json`
- `tests/wp-17/contract-validators/artifacts/<runLabel>/run_summary.txt`
