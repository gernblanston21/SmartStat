---
name: smartstat-determinism-audit
description: Use this skill to prove SmartStat determinism across runs. Run harness or compare OperatorDiag logs, compute stable hashes (timestamp-normalized), and fail closed if output differs.
---

# SmartStat Determinism Audit

## When to use
- "Is this change deterministic?"
- "Compare two runs"
- "Show determinism evidence"
- WP-10/WP determinism tasks and RC validation

## Repo truth sources
- Harness packs: `tests/wp-10/**`
- Logs: `DiagLogs/**`

## Workflow
1) Produce two comparable artifacts (two log files OR two harness outputs).
2) Normalize timestamps/volatile tokens.
3) Compute stable hashes.
4) Diff normalized outputs.
5) Fail closed if any mismatch exists.

## Scripts
- `scripts/determinism_smoke.ps1` (run twice and compare)
- `scripts/stable_hash_diag.ps1` (normalize + hash)
- `scripts/diff_text.ps1` (line diff helper)
