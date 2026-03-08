# OnAir Entity Dictionary

## Purpose / Overview

Entity token reference with allowed parameter forms extracted from Entities sheets.

## Source workbook coverage

- Workbook(s): MLB_OnAir_v3_Stat_Syntax_v3.xlsx, NHL_OnAir_v3_Stat_Syntax_v3.xlsx
- Sheets used: Entities

## Extraction notes / normalization notes

- Header normalization: `Entity Syntax` and `Entity` are normalized to `entity_name`.
- Example usage is rendered deterministically from entity name and source parameter text.

## Extracted reference

| entity_name | allowed_parameters | example_usage | league |
| --- | --- | --- | --- |
| away |  | away | MLB |
| batter | only available in Batter vs Pitcher category Playbook pages | batter(only available in Batter vs Pitcher category Playbook pages) | MLB |
| coach | TRICODE | coach(TRICODE) | MLB |
| home |  | home | MLB |
| league | player, team | league(player, team) | MLB |
| pitcher | only available in Batter vs Pitcher category Playbook pages | pitcher(only available in Batter vs Pitcher category Playbook pages) | MLB |
| player | TRICODE, ## | player(TRICODE, ##) | MLB |
| team | TRICODE, home, away, us, opp, them | team(TRICODE, home, away, us, opp, them) | MLB |
| them |  | them | MLB |
| time | (info requests only) | time((info requests only)) | MLB |
| us |  | us | MLB |
| away |  | away | NHL |
| coach | TRICODE | coach(TRICODE) | NHL |
| home |  | home | NHL |
| league | player, team | league(player, team) | NHL |
| player | TRICODE, ## | player(TRICODE, ##) | NHL |
| team | TRICODE, home, away, us, opp, them | team(TRICODE, home, away, us, opp, them) | NHL |
| them |  | them | NHL |
| time | (info requests only) | time((info requests only)) | NHL |
| us |  | us | NHL |

## SmartStat Relevance

- Defines entity tokens for semantic dictionaries.
- Supports deterministic semantic selection and context scoping.
- Feeds planner grammar entity positions without runtime behavior changes.
