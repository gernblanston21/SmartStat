# SmartStat v4 Release Governance

## Current Line
- Active development branch: `v4_Dev`
- Frozen baseline: `v4.0.0_beta`
- Target release candidate: `v4.0.0_RC1`

This document defines version boundary rules and qualification criteria for the v4.0.0 release line.

---

# 1. Version Philosophy

The v4.0.x line is a stability line, not a feature expansion line.

- `v4.0.0_beta` — Structural freeze completed.
- `v4.0.0_RC1` — Production readiness validation layer.
- `v4.1.0+` — Feature or architectural evolution.

Any change that modifies runtime semantics automatically disqualifies the change from the v4.0.x line and requires version advancement to v4.1.0 or higher.

---

# 2. v4.0.0_RC1 Qualification Scope

## 2.1 Allowed Changes (RC1-Safe)

The following changes are permitted in RC1:

### A) Logging & Diagnostics
- Clarify log wording without changing logic.
- Normalize log formatting for consistency.
- Reduce verbosity if logging remains available under `DIAG_MODE=True`.
- Add logging-only transparency for troubleshooting.

### B) Defensive Guards (No Behavior Change)
- Add null checks (`Is Nothing`), type checks, or safe early returns to prevent runtime faults.
- Add idempotence guards to prevent duplicate artifact emission.
- Fix implementation defects where intended behavior already exists but execution is structurally incorrect.
- Add `Err.Clear` or controlled error handling without altering control flow.

### C) Harness Improvements (Non-Behavioral)
- Improve artifact formatting, naming, or metadata clarity.
- Improve STRICT mode messaging (without changing gating rules).
- Add additional logging-only harness validation checks.

### D) Documentation & Governance
- Update SESSION.md.
- Update CHANGELOG.
- Add or refine operator runbooks.
- Add known limitations documentation.

### E) Repository Hygiene
- Comment clarity.
- Non-functional formatting.
- Documentation reorganization.
- .gitignore or non-runtime configuration cleanup.

---

# 3. Explicitly Forbidden in RC1

The following are NOT permitted in RC1:

## 3.1 Runtime Semantics Changes
- Resolver scoring math modifications.
- Threshold or distance logic changes.
- Candidate ordering changes in selection path.
- Tie-break rule modifications (must remain FIRST_SEEN).
- Ambiguity gating changes.
- Output_map inference logic changes.
- ApplyPlan / transaction behavior changes.
- Phase ordering changes.
- New early-exit behaviors that alter control flow.

## 3.2 Schema & Contract Changes
- INI key additions or changes.
- Section semantic modifications.
- Naming convention changes.
- Breaking changes affecting SmartStatTrayApp or external tooling.

## 3.3 Refactors
- Function renaming or extraction.
- Reordering structural blocks.
- Code cleanup that changes diff footprint substantially.
- INI reordering.

If a change cannot be categorized as logging-only, guard-only, harness-only, documentation-only, or hygiene-only, it is not RC1-safe.

---

# 4. RC1 Acceptance Criteria

All of the following must pass before tagging `v4.0.0_RC1`.

## 4.1 Harness Validation

- Run `HARNESS_STRICT` across 5–10 representative templates.
- Validate:
  - STRICT blocks commit when `diffCount > 0`.
  - STRICT allows commit when `diffCount = 0`.
  - No duplicate artifact emissions.
  - Snapshot and grouped diff alignment confirmed.
- No unexpected runtime errors.

## 4.2 Behavioral Lock Verification

- No new `Phase_Fail` codes introduced.
- No changes to ambiguity behavior.
- No changes to resolver tie-break behavior.
- No changes to transaction validation semantics.
- No change in `OUTMAP.EMPTY` enforcement.

## 4.3 Execution Integrity

- Script loads via:
  cscript //nologo SmartStat_v4.0.0_beta.vbs
  without host-level runtime errors.
- DIAG logs show no unexpected new abort patterns.
- Ambiguity detail remains capped at 5 entries with truncation marker.
- Resolver logging confirms tie_rule=FIRST_SEEN.

---

# 5. Version Boundary Rule

If any proposed change affects:

- Scoring math
- Selection ordering
- Ambiguity gating
- ApplyPlan logic
- Output_map semantics
- Config schema

Then the change must target:

v4.1.0

and must not be merged into the v4.0.x line.

---

# 6. Branching Policy

- RC1 work should occur on a dedicated branch (e.g., `v4_RC`).
- Only RC1-qualified changes may be merged.
- Diff footprint must remain minimal and reviewable.
- All changes must be categorized as:
  - LOG
  - GUARD
  - HARNESS
  - DOC
  - HYGIENE

Any commit outside those categories requires version escalation.

---

# 7. Release Tagging Rule

`v4.0.0_RC1` may be tagged only when:

- All acceptance criteria pass.
- STRICT harness artifacts are archived.
- No open stability defects exist.
- No unresolved ambiguity leakage is present.
- No nondeterministic resolver behavior has been introduced.

---

# 8. Stability Commitment

The v4.0.x line guarantees:

- Fail-closed ambiguity enforcement.
- FIRST_SEEN deterministic tie resolution.
- Strict transaction validation.
- Harness-verifiable commit safety.
- No silent early exits.
- No unbounded ambiguity dumps.
- No hidden override behavior.
- No uncontrolled diff commits.

Any violation of these guarantees requires version advancement.

---

End of governance document.
