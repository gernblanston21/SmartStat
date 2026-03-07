# OnAir Formatter Dictionary

## Purpose / Overview

Formatter references and usage context metadata from Formatters sheets.

## Source workbook coverage

- Workbook(s): MLB_OnAir_v3_Stat_Syntax_v2.xlsx, NHL_OnAir_v3_Stat_Syntax_v2.xlsx
- Sheets used: Formatters (8), Formatters (8)

## Extraction notes / normalization notes

- Usage context is tagged as `pipe_formatter` for pipe-prefixed formatter names; otherwise `measure_modifier`.

## Extracted reference

| formatter_name | description | usage_context | league | source_sheet | notes |
| --- | --- | --- | --- | --- | --- |
| (ingame) | Add after a measure to include stats from the current game | measure_modifier | MLB | Formatters |  |
| (pregame) | Add after a measure to exclude current game stats | measure_modifier | MLB | Formatters |  |
| *N | Multiply N to Measure | measure_modifier | MLB | Formatters |  |
| +N | Add N to Measure | measure_modifier | MLB | Formatters |  |
| -N | Subtract N to Measure | measure_modifier | MLB | Formatters |  |
| /N | Divide N to Measure | measure_modifier | MLB | Formatters |  |
| \| day_long | applies to any "day_of_week" attribute | pipe_formatter | MLB | Formatters |  |
| \| day_short | applies to any "day_of_week" attribute | pipe_formatter | MLB | Formatters |  |
| \| file_path | Replaces spaces with _ and removes any special characters | pipe_formatter | MLB | Formatters |  |
| \| flex_long | returns year or noyear format based on how long ago the date was and applies to any "date" attribute | pipe_formatter | MLB | Formatters |  |
| \| flex_short | returns year or noyear format based on how long ago the date was and applies to any "date" attribute | pipe_formatter | MLB | Formatters |  |
| \| long_noyear | applies to any "date" attribute | pipe_formatter | MLB | Formatters |  |
| \| long_year | applies to any "date" attribute | pipe_formatter | MLB | Formatters |  |
| \| lowercase | overrides any formatting set in the Text Formatting menu | pipe_formatter | MLB | Formatters |  |
| \| ordinal | Adds proper ordinal to number | pipe_formatter | MLB | Formatters |  |
| \| short_noyear | applies to any "date" attribute | pipe_formatter | MLB | Formatters |  |
| \| short_year | applies to any "date" attribute | pipe_formatter | MLB | Formatters |  |
| \| smallcaps | overrides any formatting set in the Text Formatting menu | pipe_formatter | MLB | Formatters |  |
| \| title | overrides any formatting set in the Text Formatting menu | pipe_formatter | MLB | Formatters |  |
| \| uppercase | overrides any formatting set in the Text Formatting menu | pipe_formatter | MLB | Formatters |  |
| *N | Multiply N to Measure | measure_modifier | NHL | Formatters |  |
| +N | Add N to Measure | measure_modifier | NHL | Formatters |  |
| -N | Subtract N to Measure | measure_modifier | NHL | Formatters |  |
| /N | Divide N to Measure | measure_modifier | NHL | Formatters |  |
| \| day_long | applies to any "day_of_week" attribute | pipe_formatter | NHL | Formatters |  |
| \| day_short | applies to any "day_of_week" attribute | pipe_formatter | NHL | Formatters |  |
| \| file_path | Replaces spaces with _ and removes any special characters | pipe_formatter | NHL | Formatters |  |
| \| flex_long | returns year or noyear format based on how long ago the date was and applies to any "date" attribute | pipe_formatter | NHL | Formatters |  |
| \| flex_short | returns year or noyear format based on how long ago the date was and applies to any "date" attribute | pipe_formatter | NHL | Formatters |  |
| \| long_noyear | applies to any "date" attribute | pipe_formatter | NHL | Formatters |  |
| \| long_year | applies to any "date" attribute | pipe_formatter | NHL | Formatters |  |
| \| lowercase | overrides any formatting set in the Text Formatting menu | pipe_formatter | NHL | Formatters |  |
| \| ordinal | Adds proper ordinal to number | pipe_formatter | NHL | Formatters |  |
| \| short_noyear | applies to any "date" attribute | pipe_formatter | NHL | Formatters |  |
| \| short_year | applies to any "date" attribute | pipe_formatter | NHL | Formatters |  |
| \| smallcaps | overrides any formatting set in the Text Formatting menu | pipe_formatter | NHL | Formatters |  |
| \| title | overrides any formatting set in the Text Formatting menu | pipe_formatter | NHL | Formatters |  |
| \| uppercase | overrides any formatting set in the Text Formatting menu | pipe_formatter | NHL | Formatters |  |

## SmartStat Relevance

- Supports semantic formatter dictionaries.
- Supports explainability of expression post-processing hints.
- Supports planner grammar formatter references without runtime behavior.
