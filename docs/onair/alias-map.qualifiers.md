# OnAir Qualifier Phrase Alias Map

## Purpose / Overview

Alias phrases mapped to canonical qualifier/filter forms where alias columns are explicitly available.

## Source workbook coverage

- Workbook(s): MLB_OnAir_v3_Stat_Syntax_v2.xlsx, NHL_OnAir_v3_Stat_Syntax_v2.xlsx
- Sheets used: QUALIFIER_TO_FILTER (3), QUALIFIER_TO_FILTER (2)

## Extraction notes / normalization notes

- Only source-grounded alias rows are emitted; no free-form alias invention.
- Canonical filter/parameter parse supports structured forms like `all_star_break(after)`.

## Extracted reference

| alias_phrase | canonical_filter | canonical_parameter | league | source_sheet | example | notes |
| --- | --- | --- | --- | --- | --- | --- |
| AFTER ALL-STAR BREAK | all_star_break | after | NHL | QUALIFIER_TO_FILTER |  | Grounded in `Aliases` column. |
| AFTER ALL-STAR GAME | all_star_break | after | NHL | QUALIFIER_TO_FILTER |  | Grounded in `Aliases` column. |
| BEFORE ALL-STAR BREAK | all_star_break | before | NHL | QUALIFIER_TO_FILTER |  | Grounded in `Aliases` column. |
| BEFORE ALL-STAR GAME | all_star_break | before | NHL | QUALIFIER_TO_FILTER |  | Grounded in `Aliases` column. |
| POST ALL-STAR BREAK | all_star_break | after | NHL | QUALIFIER_TO_FILTER |  | Grounded in `Aliases` column. |
| POST ALL-STAR GAME | all_star_break | after | NHL | QUALIFIER_TO_FILTER |  | Grounded in `Aliases` column. |
| PRE ALL-STAR BREAK | all_star_break | before | NHL | QUALIFIER_TO_FILTER |  | Grounded in `Aliases` column. |
| PRE ALL-STAR GAME | all_star_break | before | NHL | QUALIFIER_TO_FILTER |  | Grounded in `Aliases` column. |
| SINCE ALL-STAR BREAK | all_star_break | after | NHL | QUALIFIER_TO_FILTER |  | Grounded in `Aliases` column. |
| SINCE ALL-STAR GAME | all_star_break | after | NHL | QUALIFIER_TO_FILTER |  | Grounded in `Aliases` column. |

## SmartStat Relevance

- Supports qualifier alias normalization.
- Improves deterministic candidate-resolution wording.
- Supplies planner synonym scaffolding without planner execution behavior.
