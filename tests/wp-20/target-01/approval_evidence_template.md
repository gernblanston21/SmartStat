# WP-20 Approval Evidence Template (Target-01)

Use this template in the future WP-20 implementation lane to submit the
required approval package before code changes begin.

## 1) Lane Charter

- Approved scope:
- Explicit non-goals:
- Risk-class statement:
- Approval references:

## 2) Branch Isolation

- Proposed branch/lane:
- Isolation justification:
- Merge ownership:
- Rollback ownership:

## 3) Regression/Evidence Plan

- Baseline evidence sources:
- Determinism comparison method:
- Fail-closed checks:
- Boundary checks:

## 4) Rollback Plan

- Rollback trigger conditions:
- Rollback procedure:
- Validation after rollback:

## 5) Abort Criteria

- Immediate halt triggers:
- Escalation owner:
- Decision log location:

## 6) Prototype Acceptance Gates

- Gate A Scope: PASS/FAIL evidence path
- Gate B Determinism: PASS/FAIL evidence path
- Gate C Fail-Closed: PASS/FAIL evidence path
- Gate D Boundary: PASS/FAIL evidence path
- Gate E Rollback: PASS/FAIL evidence path
- Gate F Governance: PASS/FAIL evidence path

## 7) Protected Surface Confirmation

- `SmartStat_v4.0.0_beta.vbs`: changed/unchanged
- `SmartStat_TemplateConfig.ini`: changed/unchanged
- `SmartStat_Mappings.ini`: changed/unchanged
- `SmartStat_StaticOverrides.ini`: changed/unchanged
- `docs/onair/plan-capture.schema.json`: changed/unchanged
- `docs/onair/plan-validation-contract.md`: changed/unchanged
