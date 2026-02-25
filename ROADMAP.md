# ROADMAP.md — SmartStat v4.0.0_beta (SmartStat_v4.0.0_beta.vbs)

This roadmap is written to be **directly actionable in Codex 5.3**: it’s broken into **milestones**, **work packages**, and **copy/paste tasks** with clear “Definition of Done” and regression checks.

---

## Role Summary
**Owner:** SmartStat Core Engine (VBScript + INI system)
**Primary goal:** Compile a reliable, operator-safe “apply plan” that converts template tabfields + INI config into correct Trio custom properties / values, with **fail-closed** ambiguity handling and strong diagnostics.

---

## Scope

### In scope
- `SmartStat_v4.0.0_beta.vbs`
- `SmartStat_Mappings.ini`
- `SmartStat_Mappings.learn.ini`
- `SmartStat_MappingsNBA.ini`
- `SmartStat_MappingsNBA.learn.ini`
- `SmartStat_MappingsNHL.ini`
- `SmartStat_MappingsNHL.learn.ini`
- `SmartStat_StaticOverrides.ini`
- `SmartStat_TemplateConfig.ini`

### Explicit non-goals (guardrails)
- **No naming convention changes** (template names, section names, tabfield patterns, etc.)
- **No redesign** of Viz Trio tabfield patterns / VTW conventions
- **No breaking changes** to external tools (flag anything that might impact SmartStatTrayApp or future GUI tooling)

---

## Current v4.0.0_beta architecture (what already exists)

SmartStat is already structured as a **phase pipeline**:

- `00.BOOT`
- `01.ENV_VALIDATE`
- `02.LOAD_CONFIG`
- `03.CLASSIFY_FIELDS`
- `04.APPLY_OVERRIDES`
- `05.DETECT_FILTERS_CATS`
- `06.BUILD_OUTPUT_MAP`
- `07.BUILD_SYNTAX`
- `08.PUSH_TO_TRIO`
- `99.DONE`

Key systems present in the script today:
- **DIAG logging** (operator + env check logs, file-size capped)
- **Fail-closed ambiguity gating** (ambiguity recorded + commit blocked unless allowed)
- **INI-driven config** (TemplateConfig + Mappings + StaticOverrides + learn)
- **Harness + Integrity** mode (capture/pre-state, diff, optional commit behaviors)
- **Transaction mindset** (build plan first, validate, then apply)

---

## Release philosophy (v4.x)

- **Patch releases (4.0.1, 4.0.2…)**: bug fixes, log clarity, resolver edge cases, safer defaults.
- **Minor releases (4.1.0…)**: new features that preserve existing behavior by default, new config keys (only if justified), new harness modes.
- **Structural releases (4.2.0+ / 5.0.0)**: large workflow changes, format changes, or compatibility impacts.

---

## Definition of Done (global)

A change is “done” only when all are true:

1. **No WSH runtime errors** with `cscript //nologo SmartStat_v4.0.0_beta.vbs`
2. **No new ambiguity regressions**: ambiguous cases still block commit unless explicitly allowed.
3. **No INI governance violations**:
   - do not reorder keys in existing sections
   - preserve spacing/formatting
4. **Logs remain readable** and do not spam per-line `Err.Number` noise.
5. **Harness integrity diff** remains valid and continues to work on both success + fail paths.

---

## Milestones (high-level)

### M0 — Baseline Stability & Repeatability (immediate)
**Goal:** Make v4.0.0_beta rock-solid for daily operator use and for Codex-driven iteration.

Deliverables:
- Consistent phase boundary logging
- Deterministic plan validation outcomes
- Clear operator-facing error messages (no popups that block workflows)
- Verified harness behavior (capture/diff/commit)

---

### M1 — Resolver Accuracy & Ambiguity Discipline (next)
**Goal:** Improve mapping resolution accuracy while keeping “fail-closed” behavior.

Deliverables:
- Better fuzzy resolve ordering (full-string before tokenization, stable scoring)
- Operator shorthand ambiguity guard coverage
- Clear “why it failed” logging: candidate list + rejection reason
- Learn logging writes only what it should (no garbage keys)

---

### M2 — Output Map Coverage (next)
**Goal:** Ensure `output_map` generation produces *usable mappings* on real templates, not empty output.

Deliverables:
- Strong detection of category?output columns (including “H0100 => H0110/H0120/H0130” patterns)
- Correct per-row filter application (no stacking all filters onto one path)
- Safer handling of missing/blank tabfields (skip, don’t poison plan)

---

### M3 — Overrides & Entity Subtype Behaviors (next)
**Goal:** Make overrides predictable and auditable.

Deliverables:
- StaticOverrides precedence rules locked + documented
- Player subtype overrides (like `player-p`) remain localized to the intended engine behaviors
- Cross-league override compatibility checks (NBA/NHL/MLB configs)

---

### M4 — Harness Expansion (later)
**Goal:** Turn Harness into a real regression safety net.

Deliverables:
- Golden snapshot capture mode (template + CP map + generated syntax)
- Diff reports that are operator-readable (what changed, where, why)
- Optional “dry-run only” mode with full plan output

---

### M5 — SmartStatTrayApp Readiness (later)
**Goal:** Make v4 config + plan model clean enough that a WinForms tray UI can load, modify, and apply safely.

Deliverables:
- A stable internal “plan schema” (dictionary keys consistent + documented)
- Predictable per-tabfield classification outputs (filters/categories/outputs)
- No “hidden magic” that the GUI can’t mirror

---

## Work Packages (Codex-ready)

Each work package below is sized so Codex can implement it without getting lost.

---

## WP-01: Phase Pipeline Hardening

### Objectives
- Ensure each phase has:
  - an entry log
  - a completion log
  - a clear reason on early exit
- Ensure early exits still produce:
  - Harness diff (if harness active)
  - Operator alert (non-blocking)

### Tasks
- [ ] Normalize phase enter/exit calls (`Diag_Step`, `Diag_Mark_*`)
- [ ] Centralize “early exit” reasons into a single helper (consistent text)
- [ ] Confirm `FinalizeAndRefresh` runs on both success and fail paths (as intended)

### Done when
- Logs show a clean phase progression OR a clear early-exit reason
- No duplicate/missing phase marks

---

## WP-02: INI Load & Path Discovery Robustness

### Objectives
- Hardening around:
  - missing files
  - nested folder case (`SmartStat\SmartStat\...`)
  - read failures
  - encoding / weird line endings

### Tasks
- [ ] Improve `Diag_Check_ConfigPresence` messaging (exact missing filename)
- [ ] Ensure load order is deterministic (TemplateConfig ? Mappings ? Overrides ? learn)
- [ ] Add a single summary block in logs: loaded file paths + counts (sections/keys)

### Done when
- Operator can immediately see what config files were used
- Missing config exits are clean and actionable

---

## WP-03: Field Classification Reliability

### Objectives
- Classification is the backbone for `filter_tabfields`, `category_tabfields`, and `output_map`.

### Tasks
- [ ] Ensure classification respects known prefix heuristics:
  - `A####` toggles (deprioritize)
  - `E####` sponsors (deprioritize)
  - `B####` qualifier/filter/toggle candidates
  - `C####` row/column controls
  - `H####`–`Y####` prioritize for player data + output_map
- [ ] Emit a “classification table” into diag logs (tabfield ? role ? reason)

### Done when
- Misclassified fields are easy to debug from logs without guessing

---

## WP-04: Ambiguity Gating Refinement (Fail-Closed)

### Objectives
- Keep the strong “block commit if ambiguous” behavior, but make it more transparent.

### Tasks
- [ ] Ensure ambiguity hits record:
  - input string
  - top candidates (top 2 is fine)
  - why #1 wasn’t confidently selected (if applicable)
- [ ] Ensure “allowed commit” conditions are explicit and logged

### Done when
- An operator can understand ambiguity without reading the code

---

## WP-05: Output Map Generation Coverage (Highest Value)

### Objectives
- Reduce cases where templates end up with `output_map=` empty or useless.

### Tasks
- [ ] Implement category?output column pairing patterns beyond sequential:
  - example: `H0100` category column maps to `H0110/H0120/H0130` outputs
- [ ] Ensure per-row mapping is preserved (no flattening)
- [ ] Ensure output_map ignores irrelevant tabfields

### Done when
- Real-world templates produce a populated `output_map` with correct targets

---

## WP-06: Syntax Builder Determinism

### Objectives
- Given the same inputs, generated syntax should always match (no ordering randomness).

### Tasks
- [ ] Stabilize dictionary iteration ordering where it impacts output
- [ ] Add a single “syntax preview” block in logs (per output tabfield)

### Done when
- Comparing two runs with identical inputs yields identical syntax output

---

## WP-07: StaticOverrides Precedence & Audit Trail

### Objectives
- Overrides must be obvious and traceable.

### Tasks
- [ ] Log: “override applied” with:
  - override source section
  - affected tabfield(s)
  - old ? new mapping target
- [ ] Ensure override precedence is consistent (documented in code + logs)

### Done when
- Operator can tell exactly why a mapping changed

---

## WP-08: Harness + Integrity Expansion (Regression Net)

### Objectives
- Harness becomes the go-to tool for safe iterative changes.

### Tasks
- [ ] Add “capture-only” mode that writes a structured snapshot
- [ ] Improve diff formatting (group by: CP changes vs visible value changes)
- [ ] Ensure harness runs even when config load fails (if enabled)

### Done when
- Harness can be used as a regression test across templates

---

## Testing Matrix (manual + harness)

### Minimum manual tests (per change)
- [ ] Template with known good config (no ambiguity) ? commits successfully
- [ ] Template with intentional ambiguity ? blocks commit, logs candidates
- [ ] Template missing config block ? exits clean, operator message is useful
- [ ] StaticOverrides template ? override is applied + logged
- [ ] Output_map heavy template ? output_map populated and correct

### Harness tests (recommended)
- [ ] HARNESS_CAPTURE: snapshot before/after a run
- [ ] HARNESS: diff-only (no commit)
- [ ] HARNESS_COMMIT: verify intended writes only

---

## Changelog policy (draft block format)

Use this structure for each release:

```md
## [4.0.X] - YYYY-MM-DD
### Added
- ...

### Changed
- ...

### Fixed
- ...

### Notes
- Harness impact:
- Compatibility impact (TrayApp / configs):
