# Runtime Slice-1 Live Validation Intake Checklist

Scope: `WP20_RUNTIME_SLICE_01_READONLY_INGRESS`  
Pass type: operator-executed live evidence capture.

## Preconditions

- [ ] Confirm `SmartStat_v4.0.0_beta.vbs` is untouched.
- [ ] Confirm runtime successor line is `SmartStat_v4.1.0.vbs`.
- [ ] Confirm this pass is validation-only (no runtime edits).
- [ ] Confirm live operator trigger method is approved for this environment.

## Gate-OFF Parity Capture

- [ ] Capture baseline snapshot file:
  - `tests/_scratch/runtime-slice-01-readonly-ingress/live_evidence/gateoff_baseline/live_gateoff_baseline_snapshot.txt`
- [ ] Capture successor snapshot file:
  - `tests/_scratch/runtime-slice-01-readonly-ingress/live_evidence/gateoff_successor/live_gateoff_successor_snapshot.txt`
- [ ] Confirm both snapshots were captured from equivalent page/template state.

## Gate-ON Determinism Capture

- [ ] Capture run1 JSON:
  - `tests/_scratch/runtime-slice-01-readonly-ingress/live_evidence/gateon_run1/live_gateon_ingress_run1.json`
- [ ] Capture run2 JSON:
  - `tests/_scratch/runtime-slice-01-readonly-ingress/live_evidence/gateon_run2/live_gateon_ingress_run2.json`
- [ ] Capture run1 log (recommended):
  - `tests/_scratch/runtime-slice-01-readonly-ingress/live_evidence/gateon_run1/live_gateon_ingress_run1.log`
- [ ] Capture run2 log (recommended):
  - `tests/_scratch/runtime-slice-01-readonly-ingress/live_evidence/gateon_run2/live_gateon_ingress_run2.log`
- [ ] Confirm run1 and run2 used identical operator inputs.

## Operator Notes

- [ ] Fill:
  - `tests/_scratch/runtime-slice-01-readonly-ingress/live_evidence/operator_notes/operator_observations.md`
- [ ] Record any anomalies and exact step/time observed.

## Deterministic Review Output

- [ ] Run:
  - `powershell -ExecutionPolicy Bypass -File tests/_scratch/runtime-slice-01-readonly-ingress/tools/live_validation_compare.ps1`
- [ ] Confirm summary created:
  - `tests/_scratch/runtime-slice-01-readonly-ingress/live_evidence/live_validation_summary.md`
- [ ] Fill:
  - `tests/_scratch/runtime-slice-01-readonly-ingress/live_validation_review_template.md`
