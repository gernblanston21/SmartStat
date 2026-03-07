# OnAir Workbook Inventory

## Purpose / Overview

Deterministic inventory of actual workbook sheets, order, and header structures.

## MLB - `MLB_OnAir_v3_Stat_Syntax_v2.xlsx`

| order | sheet_name | header_row | headers | data_row_count |
| --- | --- | --- | --- | --- |
| 1 | CATEGORY_TO_MEASURE | 1 | VALUE, DESCRIPTION | 775 |
| 2 | CATEGORY_TO_MEASURE_PITCHER | 1 | SYNTAX, DESCRIPTION | 242 |
| 3 | QUALIFIER_TO_FILTER | 1 | Filter, Available Parameters | 50 |
| 4 | Entities | 1 | Entity, Available Parameters | 11 |
| 5 | Player-Coach Attributes | 1 | Player/Coach Attribute, Entity Supported | 28 |
| 6 | Team Attributes | 1 | Team Attribute,  | 9 |
| 7 | Time Attribute | 1 | Formatter, Function | 4 |
| 8 | Formatters | 1 | Formatter, Function | 20 |
| 9 | Additional Measures | 1 | Measure,  | 2 |
| 10 | Available Queries | 1 | Type, Query, Info | 62 |

## NHL - `NHL_OnAir_v3_Stat_Syntax_v2.xlsx`

| order | sheet_name | header_row | headers | data_row_count |
| --- | --- | --- | --- | --- |
| 1 | Entities | 1 | Key, Entity, Available Parameters, Example A, Example B, Example C, Example D | 9 |
| 2 | QUALIFIER_TO_FILTER | 1 | Aliases, Qualifier/Filter, Available Parameters, Example A, Example B, Example C, Example D | 39 |
| 3 | CATEGORY_TO_MEASURE | 1 | Measure, Description | 274 |
| 4 | CATEGORY_TO_MEASURE_GOALIE | 1 | Measure, Description | 45 |
| 5 | Player-Coach Attributes | 1 | Player/Coach Measure, Entity Supported | 26 |
| 6 | Team Attributes | 1 | Team Measure,  | 9 |
| 7 | Time Attribute | 1 | Formatter, Function | 4 |
| 8 | Formatters | 1 | Formatter, Function | 18 |
| 9 | Additional Measures | 1 | Measure,  | 2 |
| 10 | Available Queries | 1 | Type, Query, Info | 62 |

## Structural differences noted

- NHL `Entities` includes `Key` and example columns; MLB `Entities` is compact.
- MLB `CATEGORY_TO_MEASURE_PITCHER` uses header `SYNTAX`; other measure sheets use `VALUE`/`Measure`.
- NHL `QUALIFIER_TO_FILTER` has explicit `Aliases`; MLB does not.
- `Available Queries` includes shifted rows where query text appears under `Type`.
- `Time Attribute` uses `Formatter`/`Function` headers and is treated as attribute reference input.
