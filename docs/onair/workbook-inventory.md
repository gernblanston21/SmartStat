# OnAir Workbook Inventory

## Purpose / Overview

Deterministic inventory of workbook sheet order, headers, and sheet-level purpose for the v3 reference sources.

## MLB - `MLB_OnAir_v3_Stat_Syntax_v3.xlsx`

| order | sheet_name | header_row | headers | data_row_count | purpose |
| --- | --- | --- | --- | --- | --- |
| 1 | Entities | 1 | Entity Syntax, Available Parameters | 11 | Entity tokens and parameter signatures |
| 2 | CATEGORY_TO_MEASURE | 1 | Measure Syntax, Description | 775 | General measure syntax catalog |
| 3 | CATEGORY_TO_MEASURE_PITCHER | 1 | Measure Syntax, Description | 242 | Pitcher-specific measure syntax catalog (MLB) |
| 4 | QUALIFIER_TO_FILTER | 1 | Filter (Qualifier) Syntax, Available Parameters, Example(s) | 50 | Filter/qualifier syntax and parameter forms |
| 5 | Player-Coach Attributes | 1 | Main, Player/Coach Attribute, Entity Supported | 28 | Player and coach attribute names |
| 6 | Team Attributes | 1 | Main, Team Attribute,  | 9 | Team attribute names |
| 7 | Time Attribute | 1 | Formatter, Function | 4 | Time attribute and relative selectors |
| 8 | Formatters | 1 | Formatter, Function | 20 | Formatter names and descriptions |
| 9 | Additional Measures | 1 | Measure,  | 2 | Extra standalone measure tokens |
| 10 | Available Queries | 1 | Type, Query, Info | 62 | Query-family skeletons and examples |

## NHL - `NHL_OnAir_v3_Stat_Syntax_v3.xlsx`

| order | sheet_name | header_row | headers | data_row_count | purpose |
| --- | --- | --- | --- | --- | --- |
| 1 | Entities | 1 | Entity, Available Parameters | 9 | Entity tokens and parameter signatures |
| 2 | CATEGORY_TO_MEASURE | 1 | Measure, Description | 274 | General measure syntax catalog |
| 3 | CATEGORY_TO_MEASURE_GOALIE | 1 | Measure, Description | 45 | Goalie-specific measure syntax catalog (NHL) |
| 4 | QUALIFIER_TO_FILTER | 1 | Qualifier/Filter, Available Parameters, Example(s) | 39 | Filter/qualifier syntax and parameter forms |
| 5 | Player-Coach Attributes | 1 | Main, Player/Coach Measure, Entity Supported | 26 | Player and coach attribute names |
| 6 | Team Attributes | 1 | Main, Team Measure,  | 9 | Team attribute names |
| 7 | Time Attribute | 1 | Formatter, Function | 4 | Time attribute and relative selectors |
| 8 | Formatters | 1 | Formatter, Function | 18 | Formatter names and descriptions |
| 9 | Additional Measures | 1 | Measure,  | 2 | Extra standalone measure tokens |
| 10 | Available Queries | 1 | Type, Query, Info | 62 | Query-family skeletons and examples |

## MLB vs NHL structural differences

- MLB `Entities` header is `Entity Syntax`; NHL uses `Entity`.
- MLB measure headers use `Measure Syntax`; NHL uses `Measure`.
- MLB filter header is `Filter (Qualifier) Syntax`; NHL uses `Qualifier/Filter`.
- Player/coach attribute header differs: MLB `Player/Coach Attribute`, NHL `Player/Coach Measure`.
- Team attribute header differs: MLB `Team Attribute`, NHL `Team Measure`.
- MLB has `CATEGORY_TO_MEASURE_PITCHER`; NHL has `CATEGORY_TO_MEASURE_GOALIE`.
