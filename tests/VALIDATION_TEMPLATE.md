# SmartStat Validation Template

This template defines the REQUIRED harness structure for validating any SmartStat roadmap target (WP-*).

Use this template when creating new validation artifacts.

All validation files must live under:

/tests/wp-XX/target-YY/

Example:

/tests/wp-11/target-02/

---

# 1. Required Files

Each validated target MUST generate the following files:

tmp_targetXX_runner.vbs  
tmp_targetXX_strict_init.txt  
tmp_targetXX_ab_pre.txt  
tmp_targetXX_ab_post_A.txt  
tmp_targetXX_ab_post_B.txt  
tmp_targetXX_single_pre.txt  
tmp_targetXX_single_post.txt  

File naming must remain consistent across WPs.

---

# 2. Output Format (Mandatory)

All output artifacts must use EXACTLY this format:

RET=
STAT_OUT=
CONCEPT_OUT=
IS_PITCHER=
USED_HEUR=
ACCEPTED_BY=
SCORE=
AMB_COUNT=
AMB_SUMMARY=

Do not add or remove fields.

Do not change field names.

---

# 3. STRICT Initialization Check

The runner must emit:

ASSERT_STRICT_NORMAL_INIT_NO_KEY=True
STRICT_INIT_HAS_KEY=False

If either fails, the target cannot be closed.

---

# 4. A/B Reversed Insertion Test

Purpose: Detect nondeterministic enumeration.

Procedure:

Run A:
- Build candidate dictionary or enumerable in insertion order A → B.

Run B:
- Build the same structure in reversed insertion order B → A.

Required parity checks:
- Same RET
- Same ACCEPTED_BY
- Same SCORE
- Same AMB_COUNT
- Same AMB_SUMMARY
- No drift in accepted/rejected outcome

If any drift exists, target fails.

---

# 5. Single Winner Parity Test

Purpose: Ensure winner stability.

Procedure:
- Execute identical single-match case pre/post change.

Required parity:
- Same STAT_OUT
- Same CONCEPT_OUT
- Same ACCEPTED_BY
- Same SCORE
- Same AMB_COUNT

If winner changes, target fails.

---

# 6. Ambiguity Parity Test (If Applicable)

If the target touches ambiguity or tie logic:

- Include at least one tie scenario.
- Confirm ambiguity classification unchanged.
- Confirm fail-closed semantics unchanged.
- Confirm no new ambiguity introduced.

---

# 7. Ordering-Only Targets

If the target claims to be “ordering-only”:

You MUST verify:
- No scoring drift
- No tie-rule drift
- No ambiguity drift
- No acceptance/rejection drift

FIRST_SEEN rule must remain intact unless explicitly changed in roadmap.

---

# 8. Learn-Only Targets

If changes affect learn/pending suggestion logic:

You MUST verify:
- Runtime resolver behavior unchanged
- Learn output deterministic across A/B insertion
- No cross-surface leakage into resolution path

---

# 9. Duplicate Preservation Check (If Relevant)

If enumerable normalization is involved:

- Validate that duplicate values (when possible) are preserved.
- Confirm no implicit deduplication was introduced.

Note: VBScript Dictionary does not allow duplicate keys. Use arrays or collections if duplicate testing is required.

---

# 10. Target Closure Criteria

A target is CLOSED only when:

- STRICT init passes
- A/B reversed insertion parity holds
- Single-winner parity holds
- Ambiguity parity holds (if applicable)
- No logging drift
- No resolver behavior drift
- No unintended side effects

All tmp artifacts must remain in the target folder as regression evidence.

---

# 11. Engineering Discipline Rules

Do not:

- Modify SmartStat logging for validation convenience
- Modify production code to expose test-only logic
- Add hidden toggles
- Change thresholds unless roadmap explicitly states so
- Reorder INI keys unless explicitly required

All validation must occur externally via harness runners.

---

SmartStat determinism is a release-grade invariant.
All future WPs must follow this template.
