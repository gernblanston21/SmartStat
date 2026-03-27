# WP-10 Phase-4 Regression Pack

This pack validates regression determinism for WP-10 Phase-4 without changing production code or INI mappings.

Use HARNESS unless your environment requires OFF.

## Scenarios

- SUCCESS: single-winner deterministic success case.
- AMBIGUITY: deterministic ambiguous tie fail-closed case.

## Command Set

```powershell
cscript //nologo tests/wp-10/phase-4/tmp_phase4_runner.vbs SmartStat_v4.0.0_beta.vbs SUCCESS HARNESS_STRICT > tests/wp-10/phase-4/tmp_phase4_strict_run1_success.txt
cscript //nologo tests/wp-10/phase-4/tmp_phase4_runner.vbs SmartStat_v4.0.0_beta.vbs AMBIGUITY HARNESS_STRICT > tests/wp-10/phase-4/tmp_phase4_strict_run1_ambiguity.txt

cscript //nologo tests/wp-10/phase-4/tmp_phase4_runner.vbs SmartStat_v4.0.0_beta.vbs SUCCESS HARNESS_STRICT > tests/wp-10/phase-4/tmp_phase4_strict_run2_success.txt
cscript //nologo tests/wp-10/phase-4/tmp_phase4_runner.vbs SmartStat_v4.0.0_beta.vbs AMBIGUITY HARNESS_STRICT > tests/wp-10/phase-4/tmp_phase4_strict_run2_ambiguity.txt

cscript //nologo tests/wp-10/phase-4/tmp_phase4_runner.vbs SmartStat_v4.0.0_beta.vbs SUCCESS HARNESS > tests/wp-10/phase-4/tmp_phase4_nonstrict_run1_success.txt
cscript //nologo tests/wp-10/phase-4/tmp_phase4_runner.vbs SmartStat_v4.0.0_beta.vbs AMBIGUITY HARNESS > tests/wp-10/phase-4/tmp_phase4_nonstrict_run1_ambiguity.txt

cscript //nologo tests/wp-10/phase-4/tmp_phase4_runner.vbs SmartStat_v4.0.0_beta.vbs SUCCESS HARNESS > tests/wp-10/phase-4/tmp_phase4_nonstrict_run2_success.txt
cscript //nologo tests/wp-10/phase-4/tmp_phase4_runner.vbs SmartStat_v4.0.0_beta.vbs AMBIGUITY HARNESS > tests/wp-10/phase-4/tmp_phase4_nonstrict_run2_ambiguity.txt

powershell -NoProfile -ExecutionPolicy Bypass -File tests/wp-10/phase-4/tmp_phase4_hash_report.ps1
```

## Validation Mapping

- STRICT run #1 vs STRICT run #2: repeat-run determinism in HARNESS_STRICT mode.
- NONSTRICT run #1 vs NONSTRICT run #2: repeat-run determinism in non-strict mode.
- STRICT run #1 SUCCESS vs NONSTRICT run #1 SUCCESS: strict/non-strict success parity.
- AMBIGUITY outputs: deterministic fail-closed ambiguity summary across repeat runs.

## Pass/Fail Criteria

Pass if all are true in tmp_phase4_hash_report.txt:

- STRICT_1_vs_STRICT_2_HASH_EQUAL=True
- NONSTRICT_1_vs_NONSTRICT_2_HASH_EQUAL=True
- STRICT_1_vs_NONSTRICT_1_SUCCESS_HASH_EQUAL=True
- STRICT_AMBIG_1_vs_2_HASH_EQUAL=True
- NONSTRICT_AMBIG_1_vs_2_HASH_EQUAL=True
- STRICT_AMBIG_SUMMARY_1_vs_2_EQUAL=True
- NONSTRICT_AMBIG_SUMMARY_1_vs_2_EQUAL=True
- PHASE4_PASS=True

All generated files under /tests/wp-10/phase-4/ are archived regression evidence.
