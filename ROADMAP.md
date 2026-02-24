# SmartStat Development Roadmap — v4.0.0_beta (Core Engine)

## Scope
Core engine development only:
- SmartStat-Custom-Syntax-Generator (v4.0.0_beta script)
- SmartStat_Mappings.ini (+ NBA/NHL variants)
- SmartStat_StaticOverrides.ini
- SmartStat_TemplateConfig.ini

Explicitly excluded:
- Naming convention changes (forbidden)
- Viz Trio tabfield pattern redesign (forbidden)

External compatibility watch:
- SmartStatTrayApp (must flag breaking changes)

---

## Versioning (SmartStat vs VIZOR)
- SmartStat session/version baseline: user-declared (this roadmap targets v4.0.0_beta)
- VERSION.txt:
  - Value: 4.0.0
  - Purpose: VIZOR UI display only
  - Not used by SmartStat runtime; do not use as SmartStat gating/version source

---

## Current Baseline (this workspace)
- INI-based config system is present:
  - SmartStat_TemplateConfig.ini
  - SmartStat_Mappings.ini (+ NBA/NHL)
  - SmartStat_StaticOverrides.ini
- Learn layers exist:
  - SmartStat_Mappings.learn.ini
  - SmartStat_MappingsNBA.learn.ini
  - SmartStat_MappingsNHL.learn.ini
- Validator tool exists:
  - SmartStatValidator.exe
- Script target:
  - SmartStat_v4.0.0_beta.vbs

---

## Roadmap Themes (v4.0.0_beta)
1) Correctness & determinism
2) Coverage expansion (improve resolution rates, reduce empty output_map)
3) Regression discipline (repeatable validation + logs + breaking change awareness)

---

## In Progress (active work)
### 1) Ambiguity system: enable AMBIGUITY output and tracking
**Status:** NOT IMPLEMENTED (paused at analysis checkpoint)

**Last known checkpoint (where work paused):**
- You’re still not seeing AMBIGUITY because the script has:
  - No `Ambiguity_Add` function at all, and
  - No initialization for `CompilerContext("ambiguous")`.
- The shorthand guard calls a non-existent routine.
- Because `Main()` is running under `On Error Resume Next`, that call fails silently.
- The code falls through to the unresolved path.

**Planned implementation steps (smallest safe sequence):**
- [ ] Step 1: Add ambiguity context initialization (`CompilerContext("ambiguous")`) in the central context setup path.
- [ ] Step 2: Implement `Ambiguity_Add` (single responsibility: append standardized ambiguity records).
- [ ] Step 3: Replace/repair the shorthand guard so it calls a real routine and cannot silently fail.
- [ ] Step 4: Surface ambiguity in output/logs in a deterministic format.
- [ ] Step 5: Minimal validation pass with representative templates to confirm AMBIGUITY appears when expected.

**Regression focus for this item:**
- Avoid breaking unresolved-path behavior (only augment with ambiguity tracking).
- Eliminate silent missing-function calls under `On Error Resume Next`.

---

## Next Up (after ambiguity item)
- [ ] Establish “golden” validation runs (3–5 representative templates)
- [ ] Confirm and document normalization behavior (spaces ? underscores) end-to-end
- [ ] Audit TemplateConfig parsing rules for required fields and strict ordering
- [ ] Verify player-p override behavior remains correct and isolated

---

## Done (verified only)
- (none recorded yet)

---

## Validation Checklist (run after every core change)
- [ ] INI formatting unchanged (no key reordering, spacing preserved)
- [ ] Generator runs without script-stopping errors in Trio environment
- [ ] At least one NBA + one NHL template validated
- [ ] StaticOverrides still apply correctly (especially player-p)
- [ ] No new breaking assumptions introduced for SmartStatTrayApp

---

## Draft Changelog (append entries as you go)
### Unreleased — v4.0.0_beta
- (no entries yet)
