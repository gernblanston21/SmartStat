# OnAir Attribute Dictionary

## Purpose / Overview

Attribute token reference across player/coach, team, and time attribute source sheets.

## Source workbook coverage

- Workbook(s): MLB_OnAir_v3_Stat_Syntax_v3.xlsx, NHL_OnAir_v3_Stat_Syntax_v3.xlsx
- Sheets used: Player-Coach Attributes, Team Attributes, Time Attribute

## Extraction notes / normalization notes

- Header normalization aligns `Player/Coach Attribute` vs `Player/Coach Measure` and `Team Attribute` vs `Team Measure`.
- `Time Attribute` rows are included under `entity_type=time`.

## Extracted reference

| attribute_name | entity_type | league | notes |
| --- | --- | --- | --- |
| abbr_name | player | MLB | info |
| age | player | MLB | info |
| bat_hand | player | MLB | info |
| birthdate | player | MLB | info |
| college | player | MLB | info |
| draft_pick | player | MLB | info |
| draft_round | player | MLB | info |
| draft_team_alias | player | MLB | info |
| draft_team_name | player | MLB | info |
| draft_year | player | MLB | info |
| experience | player | MLB | info |
| height | player | MLB | info |
| high_school | player | MLB | info |
| hometown | player | MLB | info |
| jersey_number | player | MLB | info |
| pitcher_hand | player | MLB | info |
| position | player | MLB | info |
| primary_position | player | MLB | info |
| pro_debut | player | MLB | info |
| throw_hand | player | MLB | info |
| weight | player | MLB | info |
| first_name | player/coach | MLB | info |
| full_name | player/coach | MLB | info |
| last_name | player/coach | MLB | info |
| other_team_alias_list | player/coach | MLB | info |
| other_team_list | player/coach | MLB | info |
| team_alias_list | player/coach | MLB | info |
| team_name_list | player/coach | MLB | info |
| alias | team | MLB | info |
| altcity | team | MLB | info |
| altname | team | MLB | info |
| city | team | MLB | info |
| name | team | MLB | info |
| venue_city | team | MLB | info |
| venue_country | team | MLB | info |
| venue_name | team | MLB | info |
| venue_state | team | MLB | info |
| day_of_week | time | MLB | prev# |
| month | time | MLB | prev# |
| season | time | MLB | prev# |
| year | time | MLB | prev# |
| abbr_name | player | NHL | info |
| age | player | NHL | info |
| birthdate | player | NHL | info |
| college | player | NHL | info |
| draft_pick | player | NHL | info |
| draft_round | player | NHL | info |
| draft_team_alias | player | NHL | info |
| draft_team_name | player | NHL | info |
| draft_year | player | NHL | info |
| experience | player | NHL | info |
| handedness | player | NHL | info |
| height | player | NHL | info |
| high_school | player | NHL | info |
| hometown | player | NHL | info |
| jersey_number | player | NHL | info |
| position | player | NHL | info |
| primary_position | player | NHL | info |
| rookie_year | player | NHL | info |
| weight | player | NHL | info |
| first_name | player/coach | NHL | info |
| full_name | player/coach | NHL | info |
| last_name | player/coach | NHL | info |
| other_team_alias_list | player/coach | NHL | info |
| other_team_list | player/coach | NHL | info |
| team_alias_list | player/coach | NHL | info |
| team_name_list | player/coach | NHL | info |
| alias | team | NHL | info |
| altcity | team | NHL | info |
| altname | team | NHL | info |
| city | team | NHL | info |
| name | team | NHL | info |
| venue_city | team | NHL | info |
| venue_country | team | NHL | info |
| venue_name | team | NHL | info |
| venue_state | team | NHL | info |
| day_of_week | time | NHL | prev# |
| month | time | NHL | prev# |
| season | time | NHL | prev# |
| year | time | NHL | prev# |

## SmartStat Relevance

- Defines attribute vocabulary for semantic dictionaries.
- Supports explainable entity->attribute resolution paths.
- Supplies planner grammar attribute slots with source provenance.
