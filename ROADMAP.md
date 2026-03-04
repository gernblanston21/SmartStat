# ROADMAP - SmartStat v4 Lifecycle

## Current State
v4.0.0_beta - Frozen and stable
Transitioning to v4.0.0_RC1
v4.1.0_RC1 stabilization (Path A) is active.

## Determinism Doctrine (v4+)

SmartStat guarantees (within a given version + config):
- Identical input state + config ? identical output
- Identical STRICT harness run (same version/config) ? identical artifacts
- No implicit precedence via iteration order in resolver logic
- No filesystem-order-dependent behavior

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

## WP-10 (v4.1.0): Determinism Surface Stabilization (Explicit Ordering)

### Objective
Eliminate nondeterministic iteration surfaces that affect:
- Visible output stability
- Log emission stability
- Harness artifact reproducibility
- Resolver decision consistency
- INI traversal consistency

This is a controlled architectural stabilization pass.
No math, resolution logic, or precedence changes are permitted.

---

### Determinism Surface Classes (Mapped)

The following unordered iteration classes are in scope:

1. Dictionary key iteration (runtime collections)
2. Multi-source merge order (config overlays)
3. First-match resolver scans
4. Output_map emission ordering
5. Log/diagnostic dump ordering
6. INI section traversal
7. INI key traversal (if iterative)
8. Filesystem enumeration (if used for config load)
9. List parsing rehydration into unordered containers
10. Candidate set construction during heuristic scans

Each surface must be classified as:
- Behavior-affecting
- Presentation-only
- Non-impacting

HIGH-risk surfaces (behavioral) must be stabilized before emission-level sorting.

---

### Explicit Non-Scope

WP-10 must NOT:
- Change resolution precedence rules
- Alter ambiguity detection behavior
- Modify fail-closed gating
- Reorder INI files
- Refactor resolver algorithm design
- Introduce silent precedence changes
- Mask latent ambiguity bugs via sorting

---

### Implementation Phases

Phase 1 – Surface Audit (no code changes)
- Confirm actual presence of each determinism surface class
- Identify behavioral vs presentation-only cases
- Document implicit precedence dependencies

Phase 2 – Emission Stabilization
- Stabilize output_map ordering
- Stabilize ambiguity dump ordering
- Stabilize harness artifact serialization

Phase 3 – Behavioral Surface Hardening
- Stabilize dictionary iteration used in resolver decisions
- Stabilize overlay merge ordering explicitly
- Stabilize candidate evaluation order

### WP-10 Phase 3 — Behavioral Surface Hardening (Targets)

Completed:
- Target #1 — Commit ordering determinism (ApplyPlan.Keys sorted before non-atomic commit loop).
- Target #2 — Resolver tie determinism (two-pass tie detection; ties fail closed).
- Target #3 — TRANSFORMS_REGEX determinism (sorted load/apply; strict conflict fail-closed).
- Target #4 — AmbiguityContext lifecycle determinism + strict invariants (AMBIGUOUS_CONTEXT_INVALID; stable ambiguity emissions).
- Target #5 — TryCanonLookupFlexible: deterministic first-match behavior (normalize-collision handling).
- Target #6 — LoadIniSectionDictNormalized / LoadIni: deterministic normalized-key collision handling.
- Target #7 — SuggestQualifierMapping / ResolveFilterFragments: deterministic first-hit scanning (containment + fuzzy fallback).
- Target #8 — ResolveQualifierSmart candidate pool ordering determinism.
- Target #9 — ResolveCategorySmart alias/canonical merge determinism.
- Target #10 — Heuristic scanner input normalization

- Target #11 — Residual normalized lookup surfaces (if discovered) (Completed)

Phase-3 complete: All documented HIGH nondeterministic behavioral surfaces stabilized with ordering-only fixes. No ambiguity, scoring, resolver, or logging drift introduced.

Phase 4 – Regression Verification
- STRICT harness repeat-run validation
- Confirm identical resolution outcomes
- Confirm diffs are ordering-only
- Document validation artifact record

WP-10 CLOSED — Determinism surface stabilization and regression verification complete.

---

### Definition of Done (Expanded)

- All HIGH-risk determinism surfaces stabilized
- All MEDIUM-risk surfaces stabilized for artifact consistency
- STRICT harness repeat-run produces identical artifacts
- No new ambiguity leakage
- No resolution outcome changes
- Changelog + validation record committed

---

### Versioning

This is a minor version bump (v4.1.0) because:
- Output ordering will change
- Diff behavior changes
- Determinism guarantees are strengthened

### WP-10 Retrospective — Determinism Stabilization

WP-10 hardened SmartStat’s resolution engine to guarantee deterministic behavior across runs.
All previously identified HIGH-risk nondeterministic surfaces (dictionary iteration order, candidate pool construction, alias/canonical merge order, normalize-first lookup helpers, and heuristic scanner ingress) were stabilized without altering resolver math, scoring rules, or ambiguity policy.

The work focused exclusively on deterministic ordering and fail-closed ambiguity preservation so that identical input state and configuration now produce identical outputs every time.

Phase-4 regression verification validated repeat-run determinism in both STRICT and runtime harness modes, with archived evidence under `/tests/wp-10/phase-4/`.

This milestone establishes SmartStat’s first fully verified deterministic core and provides a stable foundation for future resolver enhancements and feature work.

No patch release permitted for this scope.

## WP-11 (v4.1.0): Harness regression pack framework
### Scope
- Define a repeatable “regression pack” set of templates / scenarios
- Standardize how artifacts are stored and compared

### Definition of Done
- Regression pack documented and runnable
- Artifact locations standardized
- Clear pass/fail criteria captured

WP-11 CLOSED — Harness regression pack framework implemented (tests-only).
Evidence: tests/wp-11/regression-pack/artifacts/compare/wp11_runA__wp11_runB/ (PACK_PASS=True)

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
