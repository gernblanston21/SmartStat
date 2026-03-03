# SmartStat Testing & Validation Discipline

This document defines the deterministic validation framework used for SmartStat development.

SmartStat follows a structured regression discipline during roadmap work packages (WP-*), especially when modifying resolution, ambiguity, or heuristic surfaces.

---

# 1. Test Artifact Structure

All harness and regression artifacts must live under:

/tests/
  wp-XX/
    target-YY/

Example:

/tests/wp-10/target-09/
/tests/wp-10/target-10/

Root-level tmp_* files are not permitted.

---

# 2. STRICT Harness Discipline

Any work that affects:

- Heuristic resolution
- Ambiguity handling
- Candidate enumeration
- Normalization logic
- Resolver ordering
- Fail-closed behavior

MUST provide STRICT validation evidence.

STRICT validation requires:

## A. Initialization Guard

ASSERT_STRICT_NORMAL_INIT_NO_KEY=True
STRICT_INIT_HAS_KEY=False

This confirms clean deterministic initialization.

---

## B. A/B Reversed Insertion Test

Used to detect nondeterministic enumeration behavior.

Two runs must be executed:

- Run A: candidate dictionary inserted in order A → B
- Run B: same dictionary inserted in reverse order B → A

Required parity checks:

- Same RET value
- Same ACCEPTED_BY
- Same SCORE
- Same AMB_COUNT
- Same AMB_SUMMARY
- No drift in accepted/rejected behavior

---

## C. Single Winner Parity

A non-ambiguous case must be tested pre/post change.

Required checks:

- Same STAT_OUT
- Same CONCEPT_OUT
- Same ACCEPTED_BY
- Same SCORE
- Same AMB_COUNT

---

# 3. Ordering-Only Stabilization Rule

For WP-10 Phase-3 surfaces:

Allowed:
- Deterministic TEXT_BINARY sorting of candidate arrays
- Normalization of enumerable inputs to arrays (ordering only)

Not allowed:
- Scoring changes
- Threshold changes
- Tie-rule changes (FIRST_SEEN remains)
- Ambiguity policy changes
- Resolver redesign
- INI key reordering
- Logging format changes

---

# 4. Learn-Only Stabilization

Changes that affect learn/pending suggestion paths must:

- Be explicitly scoped to learn-only helpers
- Not alter runtime resolver behavior
- Be validated for no ambiguity drift

---

# 5. Hybrid Evidence Model

When Trio runtime is unavailable:

- Local static proof may be used
- BUT STRICT-style output parity must still be generated via harness runner scripts

All tmp artifacts must follow the standard output format:

RET=
STAT_OUT=
CONCEPT_OUT=
IS_PITCHER=
USED_HEUR=
ACCEPTED_BY=
SCORE=
AMB_COUNT=
AMB_SUMMARY=

---

# 6. Evidence Preservation

Each target under a WP must preserve:

- Runner script
- strict_init output
- ab_pre
- ab_post_A
- ab_post_B
- single_pre
- single_post

These are considered deterministic regression evidence.

---

# 7. Definition of “Closed” for a Target

A target is considered CLOSED when:

- Code diff is minimal and scoped
- STRICT init passes
- A/B parity holds
- Single-winner parity holds
- No ambiguity drift
- No logging drift
- No resolver behavior drift

---

# 8. Future Work Packages

All future roadmap work that touches resolution, normalization, or ambiguity must follow this document.

No exceptions.

SmartStat determinism is a release-grade guarantee.
