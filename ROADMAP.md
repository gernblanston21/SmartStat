# ROADMAP - SmartStat v4 Lifecycle

## Current State
v4.0.0_beta - Frozen and stable
Transitioning to v4.0.0_RC1

---

# RC1 Phase (Immediate)

## Objectives
- Validate STRICT harness across representative templates
- Confirm no behavioral drift since beta freeze
- Verify determinism logging (no math changes)
- Confirm transaction gating integrity

## Definition of Done
- No runtime errors
- No ambiguity leakage
- No unintended diff commits
- RC1_CHECKLIST.md fully executed
- Tag v4.0.0_RC1 created

---

# Post-RC Track (v4.1.0)

## Guardrails (Post-RC)
- RC1 behavior is the baseline; changes must be intentional, scoped, and validated.
- No naming convention changes.
- No INI key reordering.
- Fail-closed semantics preserved (ambiguity + validation gates).
- Each WP must include: scope, DoD, regression plan, and validation artifacts.

## WP-10 (v4.1.0): Explicit deterministic key sorting (controlled change)
### Scope
- Introduce explicit deterministic ordering where unordered iteration can affect output stability
- Keep functional outputs identical except for deterministic ordering (no math/logic changes)

### Definition of Done
- Deterministic ordering verified across representative templates
- No new ambiguity leakage
- Harness runs produce stable outputs across repeated executions
- Changelog entry + validation record added

### Validation
- STRICT harness regression on representative templates (repeat runs)
- Diff review confirms only ordering changes

## WP-11 (v4.1.0): Harness regression pack framework
### Scope
- Define a repeatable “regression pack” set of templates / scenarios
- Standardize how artifacts are stored and compared

### Definition of Done
- Regression pack documented and runnable
- Artifact locations standardized
- Clear pass/fail criteria captured

## WP-12 (v4.1.0): Enhanced learn system validation
### Scope
- Verify learn file writes are correct, stable, and governed
- Ensure learn updates are validated (format + intent) before acceptance

### Definition of Done
- Learn write validation rules implemented
- Bad/partial writes are blocked or quarantined (fail-closed)
- Validation artifacts recorded

## WP-13 (v4.1.0): Resolver performance optimization
### Scope
- Optimize resolver hot paths without changing resolution outcomes
- Preserve determinism and logging semantics

### Definition of Done
- Performance improvement measured
- No behavioral diffs in STRICT harness regression pack

## WP-14 (v4.1.0): TrayApp alignment preparation
### Scope
- Define and stabilize the contract between SmartStat core + TrayApp
- Ensure mappings/config expectations are explicit and version-safe

### Definition of Done
- Contract documented
- Compatibility checks identified
- No breaking changes without explicit versioning

---

# Long-Term (v4.2+ / v5.0)

- Plan schema formalization
- Internal model extraction for GUI control
- Structured test harness automation
- Config validation engine

---

Release discipline enforced starting 2026-02-28.
