# OnAir Semantic Grammar (Phase 0 Synthesized Reference)

## Purpose / Overview

Phase 0 synthesized semantic reference derived from MLB/NHL OnAir v3 workbook documentation.

## Phase Status

This document represents a **Phase 0 synthesized semantic reference model** derived from the OnAir workbook documentation.

It reflects observed patterns in the reference sheets but does not claim to be the final canonical grammar of the OnAir language.

Future SmartStat planner work may refine or formalize these structures.

## Source workbook coverage

- MLB workbook: `MLB_OnAir_v3_Stat_Syntax_v3.xlsx`
- NHL workbook: `NHL_OnAir_v3_Stat_Syntax_v3.xlsx`
- Source sheets used: `Entities`, `CATEGORY_TO_MEASURE*`, `QUALIFIER_TO_FILTER`, `Player-Coach Attributes`, `Team Attributes`, `Time Attribute`, `Formatters`, `Available Queries`

## Observed Query Families

### Core semantic families

- `info`
- `stats`

### Workbook-observed families

- `calendar`
- `conditional`
- `custom`
- `game high` (also observed as `game_high`)
- `games with` (also observed as `games_with`)
- `leader`
- `math`
- `previous`
- `rank`
- `streak`

Workbook-observed families in this section may represent higher-order query operators, workbook-defined categories, or macro-style queries. They should not be assumed to be canonical semantic families.

## Operator Classes

Workbook query examples indicate an operator layer that is distinct from base families like `info` and `stats`.

- Analytical operators (workbook-observed): `rank(asc)`, `leader(1, home_runs)`, `streak(runs, >=3)`
- Dataset-selection operators (workbook-observed): `previous`, `game_high(at_bats)`, `games_with`
- Utility/calendar operators (workbook-observed): `calendar(1, sun)`
- Workbook-defined operator-style constructs: `conditional`, `custom`, `math`

These are documented as observed operator-style constructs from workbook/query evidence. They are not yet formalized as canonical planner nodes in this Phase 0 reference.

## Semantic Component Classes

- `entity`: actor/scope token from `Entities` (examples: `player`, `team`, `coach`, `time`).
- `attribute`: `info`-oriented field tokens from attribute sheets.
- `measure`: stats-valued token from `CATEGORY_TO_MEASURE*` and `Additional Measures`.
- `filter`: qualifier token from `QUALIFIER_TO_FILTER`.
- `formatter`: post-expression transform token from `Formatters`.

## Baseline Composition Patterns

Baseline canonical patterns observed in workbook examples:

- `{{info.entity.attribute}}`
- `{{stats.entity.filter.measure}}`

Observed workbook examples also show formatter pipes (for example `{{ ... | formatter }}`) and higher-order wrappers (for example `leader`, `rank`, `streak`, `previous`). Additional compositions may exist beyond this Phase 0 baseline.

## Slot-Oriented Interpretation

Observed OnAir expressions can be interpreted as typed semantic slots instead of a single unresolved string.

Example A: `{{stats.player.career.month(april).innings(7-9).hits}}`

- family slot = `stats`
- entity slot = `player`
- scope/filter slot = `career`
- filter slot = `month(april)`
- filter slot = `innings(7-9)`
- terminal measure slot = `hits`

Example B: `{{leader(1, home_runs).player.season(2023).location(away).full_name}}`

- operator slot = `leader(1, home_runs)`
- entity slot = `player`
- filter slot = `season(2023)`
- filter slot = `location(away)`
- terminal attribute slot = `full_name`

SmartStat Phase 9 candidate-resolution scaffolding can be interpreted as an early model for resolving candidates within a semantic slot.

Future planner work may apply slot-aware candidate resolution rather than treating an entire query as one unresolved blob.

This is an architecture interpretation note for direction-setting only, not implemented runtime behavior.

See [slot-resolution-model.md](e:/EDRIVE/UNIVERSAL/SmartStat/docs/onair/slot-resolution-model.md) for the slot-class architecture bridge between grammar interpretation, candidate resolution, and future plan capture.

## Argument Shape Taxonomy

### Literal placeholders

- `TRICODE`: team shorthand code parameter in entities and team-scoped filters.
- `##`: integer slot (for example entity secondary parameter).
- `##:##`: clock-time literal for time-threshold comparisons.

### Comparison operators

- `<#`, `>#`, `=#`, `<=#`, `=>#`: numeric threshold operators used in comparison-style filters.

### Range forms

- `#-#`: inclusive range-style value form used in inning/range examples.

### Relative forms

- `last#`: relative recent-window selector.
- `prev#`: previous-window selector.

### Enumerated token forms

- Explicit token lists (for example `home/away`, `yes/no`, position codes, period/state tokens).

## League/Subtype Differences

- MLB includes pitcher-specific measure sheet: `CATEGORY_TO_MEASURE_PITCHER`.
- NHL includes goalie-specific measure sheet: `CATEGORY_TO_MEASURE_GOALIE`.
- MLB entities include `batter` and `pitcher`; NHL entity sheet omits those tokens.
- Header naming differs across leagues (`Measure Syntax` vs `Measure`, `Filter (Qualifier) Syntax` vs `Qualifier/Filter`) and is normalized in extracted references.

## SmartStat Interpretation Notes

- Supports semantic dictionaries for entities, attributes, measures, filters, and formatters.
- Supports alias normalization scaffolding with explicit provenance labels (`explicit` vs inferred mappings).
- Supports deterministic candidate-resolution explainability inputs.
- Provides Phase 0 grammar interpretation notes for future planner grammar formalization.
- This document is not the final canonical SmartStat planner grammar.
