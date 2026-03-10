# Runtime Slice-1 Live Evidence Drop

Scope: `WP20_RUNTIME_SLICE_01_READONLY_INGRESS` live-validation support only.

This folder is reserved for operator-captured live Viz Trio evidence.  
Codex does not claim live execution proof from this repository alone.

## Required Drop Files

1. Gate-OFF baseline snapshot
   - `gateoff_baseline/live_gateoff_baseline_snapshot.txt`
2. Gate-OFF successor snapshot
   - `gateoff_successor/live_gateoff_successor_snapshot.txt`
3. Gate-ON run1
   - `gateon_run1/live_gateon_ingress_run1.json`
   - `gateon_run1/live_gateon_ingress_run1.log` (recommended)
4. Gate-ON run2
   - `gateon_run2/live_gateon_ingress_run2.json`
   - `gateon_run2/live_gateon_ingress_run2.log` (recommended)
5. Operator notes
   - `operator_notes/operator_observations.md`

## Intake + Comparison Workflow

1. Place live evidence files in the folders above.
2. Complete:
   - `../live_validation_intake_checklist.md`
3. Run comparison helper:

```powershell
powershell -ExecutionPolicy Bypass -File tests/_scratch/runtime-slice-01-readonly-ingress/tools/live_validation_compare.ps1
```

4. Review generated summary:
   - `live_validation_summary.md`
5. Fill:
   - `../live_validation_review_template.md`

## Notes

- If required files are missing, comparison returns `INCOMPLETE` (fail-closed).
- Gate-OFF parity compares normalized snapshot text hashes.
- Gate-ON determinism compares run1/run2 JSON hashes.
