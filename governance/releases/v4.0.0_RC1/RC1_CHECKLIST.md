# SmartStat v4.0.0_RC1 — Release Checklist

This checklist must be fully completed before tagging `v4.0.0_RC1`.

---

# 1. Scope Verification

## 1.1 Change Classification
All commits since `v4.0.0_beta` freeze must be categorized as:

- [ ] LOG (logging-only)
- [ ] GUARD (defensive, no behavior change)
- [ ] HARNESS (artifact/verification-only)
- [ ] DOC (documentation-only)
- [ ] HYGIENE (non-runtime repo cleanup)

If any commit does not fall into one of these categories, it must be version-bumped to v4.1.0.

---

# 2. Runtime Behavior Lock

## 2.1 Resolver Integrity
- [ ] No changes to scoring math.
- [ ] No changes to thresholds.
- [ ] Tie rule remains FIRST_SEEN.
- [ ] No candidate ordering modifications in selection path.
- [ ] Resolver logging still shows `tie_rule=FIRST_SEEN`.

## 2.2 Ambiguity Enforcement
- [ ] Fail-closed behavior preserved.
- [ ] EARLY EXIT markers present:
  - [ ] TX: EARLY EXIT - AMBIGUOUS_GATE
  - [ ] TX: EARLY EXIT - AMBIGUOUS_CONTEXT_INVALID
- [ ] Ambiguity detail capped at 5 entries.
- [ ] No unbounded ambiguity dumps.

## 2.3 Output Map Enforcement
- [ ] `OUTMAP.EMPTY` still hard fails.
- [ ] No inference logic changes.
- [ ] No schema modifications.

## 2.4 Transaction Integrity
- [ ] Stage_ValidatePlan behavior unchanged.
- [ ] STRICT diff>0 blocks commit.
- [ ] STRICT diff=0 allows commit.
- [ ] No false "VALIDATION FAILURE - Transaction Aborted" logs.
- [ ] No transaction flow modifications.

---

# 3. Harness Validation

Run `HARNESS_STRICT` across 5–10 representative templates.

## 3.1 Required Coverage
- [ ] Qualifier-heavy template
- [ ] Override-triggering template
- [ ] Multi-column output_map template
- [ ] No-change (zero-diff) template
- [ ] Ambiguity-triggering template

## 3.2 Required Outcomes
- [ ] diffCount=0 runs allow commit.
- [ ] diffCount>0 runs block commit.
- [ ] No duplicate artifact emission.
- [ ] Snapshot and grouped diff alignment verified.
- [ ] Post-snapshot CP/value capture verified (no inversion).

Archive STRICT artifacts before tagging.

---

# 4. Execution Integrity

- [ ] Script loads via: `cscript //nologo SmartStat_v4.0.0_beta.vbs`
