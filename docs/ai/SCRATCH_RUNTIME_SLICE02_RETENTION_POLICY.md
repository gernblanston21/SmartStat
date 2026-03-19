# _scratch Runtime Slice-02 Retention Policy

Scope:
- Applies only to `tests/_scratch/runtime-slice-02-readonly-plan-bridge/`.
- Repo hygiene only; no runtime/viewer behavior impact.

Expected Local Artifacts:
- `fixtures/`: temporary positive/negative projection fixtures used during pass validation.
- `runs/`: generated run outputs used for determinism/fail-closed checks.

Tracking Rule:
- `_scratch` artifacts remain untracked by default.
- Do not commit `_scratch` artifacts unless a pass explicitly authorizes a tracked exception.

Minimum Evidence To Keep In Tracked Docs:
- `SESSION.md` must record, per approved pass/slice:
  - objective and boundary scope
  - deterministic validation summary (including run IDs and hash parity where used)
  - fail-closed validation summary (including negative-case outcomes/error codes)
  - regression/boundary summary
- `ROADMAP.md` should include only factual state markers needed for lane/slice status.

Local Retention Window:
- Keep artifacts for the current pass/slice and immediately previous pass/slice during active iteration.
- Keep older artifacts only as needed for local troubleshooting.
- Default cleanup threshold for routine hygiene: older than 30 days and older than the two most recent pass numbers in filename prefixes (`passNN_`).

Manual Prune Triggers:
- Run a prune review after each approved closeout.
- Run a prune review when `_scratch` file count in this subtree exceeds 50 files.
- Run a prune review before preparing commits, to reduce triage noise.

Safety Rules For Pruning:
- Review candidates in dry-run mode first.
- Never prune artifacts for the currently active/just-approved pass without explicit confirmation.
- If evidence is still needed, retain locally and document outcomes in `SESSION.md` instead of committing artifacts.
