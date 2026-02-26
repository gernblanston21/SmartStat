# User Changelog

This is a plain-language version of `CHANGELOG.md`.
It explains what changed in SmartStat and what that means in day-to-day use.

## [v4_Dev] - 2026-02-26

### What this release focused on
- Making SmartStat safer when something is unclear.
- Making logs easier to understand during live troubleshooting.
- Improving output mapping so templates are less likely to end up with empty or broken results.
- Improving harness/testing modes so changes can be verified before commit.

### What you will notice
- Phase logs are clearer. You can see where the pipeline started, ended, or exited early.
- If the script exits early, the reason is now explicit (for example: template missing, qualifier unresolved, or output map empty).
- Harness modes are more useful:
  - `HARNESS`: run checks without commit.
  - `HARNESS_COMMIT`: run with commit behavior.
  - `HARNESS_CAPTURE`: write snapshot fixtures.
  - `HARNESS_STRICT`: block commit when diff is not empty.
- Output map behavior is stronger:
  - Existing explicit map entries are kept.
  - Missing entries are inferred when it is safe to do so.
- Static overrides are easier to audit:
  - Logs now show old value to new value when an override changes a tabfield.
  - If overrides INI is unavailable, the skip reason is logged.

### Important fixes included
- Ambiguity context handling is more stable (fewer invalid object-state edge cases).
- `SmartStat_MappingsNBA.learn.ini` section header issue was fixed (`[ALIASES_REGEX]`).
- Duplicate `POINTS/GM` alias in `SmartStat_MappingsNHL.learn.ini` was removed.

### Cleanup done
- Old v3.92 script snapshots were removed from `Main_TrioScript`.
- Old legacy diagnostic scripts were removed from the active v4 path.

### Compatibility and operator impact
- Fail-closed ambiguity behavior is still strict by default (`allow_ambiguous_apply=false`).
- Viz Trio naming and tabfield patterns did not change.
- No blocking popups were introduced.
- SmartStatTrayApp/socket consumers should still work with ambiguity context lines in messages.

### Recommended checks
- Validate all harness modes on a known-good template.
- Confirm unresolved qualifiers still block unsafe apply.
- Confirm empty output-map conditions still block apply.
- Confirm override logs only report real value changes.

## [v4.0.0_beta] - 2026-02-24

### What changed in plain terms
- This was the big move to the v4 engine model.
- Writes became transaction-based:
  - Build a plan first.
  - Validate it.
  - Commit only when safe.
- Ambiguous or unresolved input started blocking commit by default.
- Diagnostics became much more structured and useful for support.
- Multi-sport mapping lookup improved (MLB/NBA/NHL mapping file resolution).
- Dynamic usage paths were added for pitch-related stats.

### Why it mattered
- Lower risk of partial or unsafe writes.
- Better visibility into why a run failed.
- More consistent behavior across leagues and templates.

### Notable fixes from this baseline
- Better INI parsing reliability (including BOM handling).
- Better handling when mapping files are missing/unreadable.
- Better fuzzy-match resilience for small input mistakes.

### Config and version notes
- Core config family remained:
  - `SmartStat_Mappings*.ini`
  - `SmartStat_Mappings*.learn.ini`
  - `SmartStat_StaticOverrides.ini`
  - `SmartStat_TemplateConfig.ini`
- Security marker exists in `SmartStat_TemplateConfig.ini` under `[SECURITY]`.
- Runtime version comes from `SmartStat_v4.0.0_beta.vbs`.
- `VERSION.txt` is for VIZOR UI display, not SmartStat runtime gating.

## [v3.92] - 2025-12-23

### Plain summary
- Last major pre-v4 line.
- v4 introduced the larger safety, validation, and ambiguity-gating model after this version.
