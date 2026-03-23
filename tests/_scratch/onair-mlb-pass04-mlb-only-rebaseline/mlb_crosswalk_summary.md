# MLB Crosswalk Rebaseline Summary (PASS_04)

- Pass: `MLB_ONAIR_TO_SMARTSTAT_MLB_ONLY_CANDIDATE_REBASELINE_PASS_04`
- Mode: Read-Only Tooling / Semantic Crosswalk (MLB-only)

## Scope

- OnAir source set: PASS_02 full-index artifacts (deterministic fallback in `docs/onair/` because `tests/_scratch/onair-mlb-pass02-full-index/` does not exist in this workspace).
- MLB SmartStat sources used:
  - `SmartStat_Mappings.ini`
  - `SmartStat_Mappings.learn.ini`
- Shared SmartStat sources used:
  - `SmartStat_StaticOverrides.ini`
  - `SmartStat_TemplateConfig.ini`
- Explicit exclusions enforced:
  - `SmartStat_MappingsNBA.ini`
  - `SmartStat_MappingsNBA.learn.ini`
  - `SmartStat_MappingsNHL.ini`
  - `SmartStat_MappingsNHL.learn.ini`

## Deterministic Result Counts

- Measure crosswalk rows: **771**
  - `alias_candidate`: 60
  - `composite_candidate`: 137
  - `direct_match`: 75
  - `missing_candidate`: 499
- User dictionary crosswalk rows: **187**
  - `alias_candidate`: 1
  - `composite_candidate`: 30
  - `direct_match`: 57
  - `missing_candidate`: 99
- Composite definitions: **67**
- High-confidence missing measure candidates: **265**

## Boundary Confirmation

- No `.vbs` files modified
- No `.ini` files modified
- No runtime integration or execution-surface change performed
- Outputs are deterministic analysis artifacts only
