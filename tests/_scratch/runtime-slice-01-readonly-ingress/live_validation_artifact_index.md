# Runtime Slice-1 Live Validation Artifact Index

| File Path | Artifact Purpose | Relevance to Validation |
|---|---|---|
| `tests/_scratch/runtime-slice-01-readonly-ingress/live_evidence/live_validation_summary.md` | Computed final parity/determinism status summary | Primary pass/fail reference for closeout decision |
| `tests/_scratch/runtime-slice-01-readonly-ingress/live_evidence/gateoff_baseline/live_gateoff_baseline_snapshot.txt` | Operator-captured gate-OFF baseline state snapshot | Gate-OFF parity source A |
| `tests/_scratch/runtime-slice-01-readonly-ingress/live_evidence/gateoff_successor/live_gateoff_successor_snapshot.txt` | Operator-captured gate-OFF successor state snapshot | Gate-OFF parity source B |
| `tests/_scratch/runtime-slice-01-readonly-ingress/live_evidence/gateon_run1/live_gateon_ingress_run1.json` | Gate-ON ingress outcome run1 | Gate-ON determinism source A |
| `tests/_scratch/runtime-slice-01-readonly-ingress/live_evidence/gateon_run2/live_gateon_ingress_run2.json` | Gate-ON ingress outcome run2 | Gate-ON determinism source B |
| `tests/_scratch/runtime-slice-01-readonly-ingress/live_evidence/gateon_run1/live_gateon_ingress_run1.log` | Gate-ON operator diag/log run1 | Supplemental runtime trace context |
| `tests/_scratch/runtime-slice-01-readonly-ingress/live_evidence/gateon_run2/live_gateon_ingress_run2.log` | Gate-ON operator diag/log run2 | Supplemental runtime trace context |
| `tests/_scratch/runtime-slice-01-readonly-ingress/live_evidence/operator_notes/operator_observations.md` | Operator-entered observations/anomalies | Live execution provenance and anomaly context |
| `tests/_scratch/runtime-slice-01-readonly-ingress/live_validation_intake_checklist.md` | Required evidence intake checklist | Completeness controls for operator capture |
| `tests/_scratch/runtime-slice-01-readonly-ingress/live_validation_review_template.md` | Structured reviewer template | Standardized review capture format |
| `tests/_scratch/runtime-slice-01-readonly-ingress/live_trio_spotcheck_runbook.md` | Operator-controlled live procedure | Provenance and method reference |
| `tests/_scratch/runtime-slice-01-readonly-ingress/validation_matrix.md` | Prior fixture-based validation matrix | Supports boundary/determinism continuity |
| `tests/_scratch/runtime-slice-01-readonly-ingress/tools/live_validation_compare.ps1` | Gate-OFF/Gate-ON comparison and summary generator | Deterministic computation logic for closeout summary |
