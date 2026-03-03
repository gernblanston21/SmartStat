# WP-10 Determinism Burn-Down Tracker
Branch: v4_Dev
Version Track: v4.1.0
Scope: Determinism Surface Stabilization (Explicit Ordering)

This tracker operationalizes the Phase-1 audit from PLAN.md and tracks closure of HIGH and MEDIUM determinism surfaces.

---

# HIGH-RISK (Behavioral) Surfaces

These affect resolution outcome, ambiguity gating, commit behavior, or fragment selection.

| Target | Surface | Class | Status | Validation Artifact |
|--------|---------|-------|--------|---------------------|
| #1 | ApplyPlan.Keys commit ordering | Dictionary iteration | Complete | STRICT harness repeat-run hash match |
| #2 | Resolver tie determinism | First-match / fuzzy | Complete | Top-tie fail-closed validation |
| #3 | TRANSFORMS_REGEX load/apply order | INI traversal | Complete | Deterministic regex conflict test |
| #4 | AmbiguityContext lifecycle | Resolver gating | Complete | STRICT invariant guard test |
| #5 | TryCanonLookupFlexible normalize collision | First-match resolver | Complete | A/B normalize collision test |
| #6 | LoadIniSectionDictNormalized collision | INI traversal | Complete | Strict + non-strict collision A/B |
| #7 | SuggestQualifierMapping containment | First-hit scan | Complete | Multi-hit containment strict fail-closed |
| #8 | ResolveQualifierSmart fuzzy pool ordering | Candidate construction | Complete | A/B insertion SHA match |
| #9 | ResolveCategorySmart alias/canon merge determinism | Merge order + first-seen precedence | Complete | Mirror discipline of #8 |
| #10 | Heuristic scanner input normalization (global) | Candidate pool ordering | Complete | Eliminate hidden dict.Keys in fuzzy helpers |
| #11 | Residual normalize-first-match helpers | Normalize collision | Complete | Same pattern as #5 |

---

### Remaining HIGH Surfaces

HIGH surface count = 0.

| Target | Surface | Class | Status | Notes |
|--------|---------|-------|--------|-------|
| #11 | Residual normalize-first-match helpers | Normalize collision | CLOSED | Closed |

---

# MEDIUM-RISK Surfaces

These affect artifact reproducibility or cross-machine stability.

| Surface | Class | Status | Notes |
|---------|-------|--------|-------|
| Output_map emission ordering | Conditional behavioral | Phase-2 stabilized (no duplicate targets observed) |
| TemplateConfig traversal precedence | Behavioral (explicit default precedence) | To confirm stable |
| StaticOverrides intra-section key iteration | Behavioral (rare conflict path) | To evaluate |
| Harness fixture ordering | Presentation | Stabilized in Phase-2 |
| Socket payload tab ordering | Presentation | Out-of-scope for WP-10 determinism guarantee |

---

# LOW-RISK Surfaces

Presentation-only or already sorted.

(See PLAN.md Phase-1 audit for reference.)

---

# Burn-Down Criteria

WP-10 Behavioral Surface Hardening is complete when:

- All HIGH-risk surfaces marked
- No STRICT harness behavioral diffs vs RC1 baseline
- No new ambiguity leakage
- No implicit precedence remains in resolver logic
- Determinism Doctrine fully satisfied

---

# Determinism Doctrine Reference

- Identical input state + config - identical output
- STRICT harness repeat-run - identical artifacts
- No implicit precedence via iteration order
- No filesystem-order-dependent behavior
