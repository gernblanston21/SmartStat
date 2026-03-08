# OnAir Semantic Grammar (Canonical Phase 0 Reference)

## Purpose / Overview

Canonical semantic grammar reference synthesized from the MLB/NHL OnAir v3 workbook sheets. This document defines stable query/component composition patterns for SmartStat semantic tooling.

## Source workbook coverage

- MLB workbook: `MLB_OnAir_v3_Stat_Syntax_v3.xlsx`
- NHL workbook: `NHL_OnAir_v3_Stat_Syntax_v3.xlsx`
- Source sheets used: `Entities`, `CATEGORY_TO_MEASURE*`, `QUALIFIER_TO_FILTER`, `Player-Coach Attributes`, `Team Attributes`, `Time Attribute`, `Formatters`, `Available Queries`

## Extraction notes / normalization notes

- Grammar statements are grounded in explicit workbook syntax/examples.
- Core SmartStat planning focus remains `info` and `stats` query families; other observed families are listed as workbook-observed extensions.
- No runtime resolution/execution semantics are inferred in this phase.

## Query families

### Core families

- `info`
- `stats`

### Additional workbook-observed families

- `calendar`
- `conditional`
- `custom`
- `game high`
- `game_high`
- `games with`
- `games_with`
- `leader`
- `math`
- `previous`
- `rank`
- `streak`

## Semantic component classes

- `entity`: actor/scope token from `Entities` (examples: `player`, `team`, `coach`, `time`).
- `attribute`: `info`-oriented field tokens from attribute sheets.
- `measure`: stats-valued token from `CATEGORY_TO_MEASURE*` and `Additional Measures`.
- `filter`: qualifier token from `QUALIFIER_TO_FILTER`.
- `formatter`: post-expression transform token from `Formatters`.

## Canonical composition patterns

- `{{info.entity.attribute}}`
- `{{stats.entity.filter.measure}}`
- Formatter application observed in workbook examples: `{{ ... | formatter }}`

Representative workbook examples:

- `{{ info.coach(DET).first_name }}` (MLB)
- `{{ info.league.alias }}` (MLB)
- `{{ info.player.full_name }}` (MLB)
- `{{ info.team.name }}` (MLB)
- `{{ info.time.month(prev) }}` (MLB)
- `{{ info.time.season(last) }}` (MLB)
- `{{info.entity.attribute}}` (MLB)
- `{{ stats.away.season(2015).venue(DET).extra_inning_games(yes).team_wins }}` (MLB)

## Argument grammar (source-grounded)

| argument_shape | description | source evidence |
| --- | --- | --- |
| `TRICODE` | Team shorthand code parameter | Entities `team(TRICODE)`; filters like `on_team(TRICODE)` |
| `##` | Integer slot (example: player index/id parameter) | Entities `player(TRICODE, ##)` |
| `##:##` | Clock-time parameter | Filter examples for `game_clock` |
| `#-#` | Range interval | Filters like `innings(7-9)` examples |
| `<#`, `>#`, `=#`, `<=#`, `=>#` | Comparison operators | Margin/comparison filter examples |
| `last#`, `prev#` | Relative time/count selectors | `season`, `month`, `postseason` available parameters |
| enum lists | Explicit token sets | Available-parameter lists in `QUALIFIER_TO_FILTER` |

## League / subtype deltas

- MLB adds pitcher-specific measure sheet: `CATEGORY_TO_MEASURE_PITCHER`.
- NHL adds goalie-specific measure sheet: `CATEGORY_TO_MEASURE_GOALIE`.
- MLB entities include `batter` and `pitcher`; NHL entity sheet omits those tokens.
- Header naming differs across leagues (`Measure Syntax` vs `Measure`, `Filter (Qualifier) Syntax` vs `Qualifier/Filter`) and is normalized in extraction output.

## SmartStat Relevance

- Provides canonical semantic dictionaries for entities/measures/filters/attributes/formatters.
- Provides source-grounded qualifier alias normalization scaffolding.
- Provides deterministic input structure for candidate-resolution explainability.
- Provides grammar scaffolding for later planner and Plan Engine phases.
- Keeps runtime behavior unchanged while formalizing the semantic language surface.

## Source evidence snapshots

- Entity parameter forms observed: (info requests only), TRICODE, TRICODE, ##, TRICODE, home, away, us, opp, them, only available in Batter vs Pitcher category Playbook pages, player, team
- Filter parameter evidence examples (sample): # | innings(3), innings(7-25), # | last_game(1), last_game(5), (REG, [year]) or (PST, [year]) year = cutoff for stats | all_time, all_time(REG, 2002), all_time(2009), (REG, [year]), (PST, [year]) 'year = cutoff for stats eg. 2004' | all_time, all_time(REG, 2002), all_time(2009), 0, 1, 2 | outs(2), 0, 1, 2 | strikes(2), 0-0, 1-0, 2-0, 3-0, 0-1, 0-2, 1-1, 1-2, 2-1, 3-1, 3-2, 2-2 | count(0-0), count(0-2), 1, 2, 3, OT1, OT2, OT3, OT4, OT5, SO | period(2), period(OT1)
- Measure rows extracted: 1323
- Attribute rows extracted: 80
- Formatter rows extracted: 38
