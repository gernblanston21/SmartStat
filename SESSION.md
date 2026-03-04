# SESSION - SmartStat Core Engine

## Release State

- Stable Baseline: v4.0.0_RC1 (FROZEN + TAGGED)
- RC1 Merge Commit: c7ee0e8 (merged back into v4_Dev)
- Active Development Branch: v4_Dev
- Current Development Track: v4.1.0 (Post-RC)
- Source of Truth: v4_Dev branch workspace
- Viz Trio Reference: docs/viz-trio/ (Method A — repo grounded)

RC1 is complete and must not be re-reviewed or modified.
All new work proceeds from the post-RC baseline.

---

## Current Phase: Post-RC Structured Development (v4.1.0)

We are operating under controlled, versioned Work Packages.

Active WP:
- WP-11 — Harness regression pack framework (tests/framework)

No opportunistic refactors.
No scope creep.
Each WP must be:
- Scoped
- Documented
- Regression validated
- Changelog recorded

---

## RC1 Stabilization Discipline (Active)

Path A is active for `v4.1.0_RC1` on `v4_Dev`.
WP-10 is closed.
Phase-4 pack evidence is archived under `/tests/wp-10/phase-4/`.

Allowed changes for RC1:
- Logging clarity improvements (no behavior change)
- Guardrail reinforcement (no behavior change)
- Determinism verification additions (tests only)
- Documentation corrections

Prohibited changes for RC1:
- Any resolver behavior changes
- Any scoring/threshold/tie/policy changes
- Any INI reordering

---

## Architectural Guardrails (Active)

These rules persist across all v4.x versions:

- Fail-closed ambiguity gating remains strict.
- FIRST_SEEN tie rule remains unchanged.
- No resolver scoring math changes unless explicitly versioned.
- No INI key reordering.
- No tabfield pattern redesign.
- No silent refactors.
- No placeholders.
- Determinism must be intentional and testable.
- All SmartStat behavior must align with docs/viz-trio/.

If a proposal conflicts with:
- Determinism
- Ambiguity gating
- Transaction integrity
- INI ordering
- Viz Trio documentation

Then:
1. Halt
2. Explain conflict
3. Propose safer alternative

Fail closed by default.

---

## Stability Guarantees (Inherited from RC1)

- Deterministic resolver behavior (as of RC1)
- STRICT harness diff gating validated
- No ambiguity leakage
- Transaction validation integrity preserved
- Output_map inference stable
- Override audit trail complete
- Governance discipline enforced

These guarantees form the regression baseline for v4.1.0.

---

## Determinism Focus (WP-10)

Objective:
Introduce explicit deterministic ordering where unordered iteration affects output stability.

Constraints:
- No logic/math changes
- No resolver scoring changes
- No behavior drift beyond ordering stability
- STRICT harness must confirm stability across repeated runs

---

## Near-Term Roadmap

v4.1.0 Work Packages:
- WP-10: Deterministic key sorting
- WP-11: Harness regression pack framework
- WP-12: Enhanced learn system validation
- WP-13: Resolver performance optimization
- WP-14: TrayApp alignment preparation

---

## Operating Discipline

- Architecture discussion occurs before implementation.
- Code mutation occurs in Codex.
- Governance review occurs before merge.
- Each WP results in:
  - Diff review
  - Regression validation
  - Changelog update

Release discipline enforced starting 2026-02-28.
