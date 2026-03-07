# OnAir Entity Dictionary

## Purpose / Overview

Entity references with supported parameters and source examples.

## Source workbook coverage

- Workbook(s): MLB_OnAir_v3_Stat_Syntax_v2.xlsx, NHL_OnAir_v3_Stat_Syntax_v2.xlsx
- Sheets used: Entities (4), Entities (1)

## Extraction notes / normalization notes

- NHL includes extra `Key`/example columns; MLB is compact.
- Example fields are merged from source `Example A..D` columns when present.

## Extracted reference

| entity_name | allowed_parameters | example_usage | league | source_sheet | notes |
| --- | --- | --- | --- | --- | --- |
| away |  |  | MLB | Entities |  |
| batter | only available in Batter vs Pitcher category Playbook pages |  | MLB | Entities |  |
| coach | TRICODE |  | MLB | Entities |  |
| home |  |  | MLB | Entities |  |
| league | player, team |  | MLB | Entities |  |
| pitcher | only available in Batter vs Pitcher category Playbook pages |  | MLB | Entities |  |
| player | TRICODE, ## |  | MLB | Entities |  |
| team | TRICODE, home, away, us, opp, them |  | MLB | Entities |  |
| them |  |  | MLB | Entities |  |
| time | (info requests only) |  | MLB | Entities |  |
| us |  |  | MLB | Entities |  |
| away |  |  | NHL | Entities | AWAY |
| coach | TRICODE |  | NHL | Entities | COACH |
| home |  |  | NHL | Entities | HOME |
| league | player, team |  | NHL | Entities | LEAGUE |
| player | TRICODE, ## |  | NHL | Entities | PLAYER |
| team | TRICODE, home, away, us, opp, them |  | NHL | Entities | TEAM |
| them |  |  | NHL | Entities | OPPONENT |
| time | (info requests only) |  | NHL | Entities | TIME |
| us |  |  | NHL | Entities | PREFERRED |

## SmartStat Relevance

- Supports semantic entity dictionaries.
- Supports deterministic entity selection context for explainability.
- Supports planner entity grammar slots without runtime integration.
