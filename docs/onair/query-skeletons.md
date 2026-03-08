# OnAir Query Skeleton Grammar

## Purpose / Overview

Workbook-grounded query reference that separates baseline canonical skeleton patterns, workbook family labels, and full example evidence rows.

## Source workbook coverage

- Workbook(s): MLB_OnAir_v3_Stat_Syntax_v3.xlsx, NHL_OnAir_v3_Stat_Syntax_v3.xlsx
- Sheets used: Available Queries

## Extraction notes / normalization notes

- Canonical skeleton patterns below are normalized from workbook evidence and intended as Phase 0 baseline patterns.
- Workbook family labels are preserved from the `Type` column and should be interpreted as workbook categories, not guaranteed canonical semantic families.
- Example query strings are preserved exactly as they appear in workbook query rows.

## Section 1: Canonical Skeleton Patterns

Baseline canonical patterns observed in workbook examples:

| normalized_skeleton | evidence_basis |
| --- | --- |
| `{{info.entity.attribute}}` | `INFO` family template row in both workbooks |
| `{{stats.entity.filter.measure}}` | `STATS` family template row in both workbooks |

Additional operator-style templates are present in workbook categories (`leader`, `rank`, `previous`, `streak`, etc.) and are preserved in Section 3 as evidence rows.

## Section 2: Workbook Query Families

### Base query families

- `info`
- `stats`

### Operator-style/workbook-observed families

- `leader`
- `rank`
- `previous`
- `streak`
- `game high`
- `game_high`
- `games with`
- `games_with`
- `calendar`
- `conditional`
- `custom`
- `math`

These labels represent workbook-defined categories and should not be assumed to be canonical grammar families.

## Section 3: Example Queries

Workbook categories in this section are preserved as evidence labels from source sheets and are not necessarily the final SmartStat semantic family taxonomy.

| workbook_family_label | query_string | explanatory_notes | league |
| --- | --- | --- | --- |
| calendar | {{ calendar(1, sun).team.game_vs }} |  | MLB |
| calendar | {{ calendar(1, sun).time.day_of_month }} |  | MLB |
| calendar | {{ calendar(2, mon).team.opp_alias }} |  | MLB |
| calendar | {{ calendar(2, tue).team.opp_name }} |  | MLB |
| calendar | {{ calendar(3, thu).team.runs }} |  | MLB |
| calendar | {{ calendar(3, thu).team.venue_type }} | Will return 0 for no game, 1 for away games, and 2 for home games | MLB |
| calendar | {{ calendar(3, wed).team.game_result }} |  | MLB |
| calendar | {{ calendar(4, fri).team.runs_allowed }} |  | MLB |
| calendar | {{ calendar(5, sat).time.day_of_month }} |  | MLB |
| calendar | {{calendar(#,day).entity.filter.measure}} |  | MLB |
| conditional | {{ %if [query] = [query or value] %then [query or value] %else [query or value] %endif }} |  | MLB |
| custom | {{custom.user-dictionary-variable}} |  | MLB |
| game high | {{game_high.entity.filter.measure}} |  | MLB |
| game_high | {{ game_high(at_bats).player(LAD, 5).season(last5).at_bats }} | The most at bats Freddie Freeman has had in a game in the last 5 seasons | MLB |
| game_high | {{ game_high(runs).team(SF).all_time(REG, 1876).runs }} | The most runs the Giants have ever scored in a game | MLB |
| games with | {{games_with.entity.filter.measure}} |  | MLB |
| games_with | {{ games_with.player(NYY, 99).season(2024).minimum(home_runs, 2).games_played }} | The amount of multi-home run Aaron Judge has this season | MLB |
| games_with | {{ games_with.team(LAD).season(2024).minimum(runs, 10).home_runs }} | The amount of overall home runs the Dodgers have hit in the games that they have scored at least 10 runs this season | MLB |
| info | {{ info.coach(DET).first_name }} | First Name of Detroit's Coach | MLB |
| info | {{ info.league.alias }} | League Name Abbreviation | MLB |
| info | {{ info.player.full_name }} | Player's Full Name | MLB |
| info | {{ info.team.name }} | Team's Name | MLB |
| info | {{ info.time.month(prev) }} | Previous Month | MLB |
| info | {{ info.time.season(last) }} | Last Season | MLB |
| info | {{info.entity.attribute}} |  | MLB |
| leader | {{ leader(1, fielding_steals_against, desc).team.season.alias }} | The alias of the team that allowed the most steals this season (overrides the "asc" sort order in the Stat Defaults menu) | MLB |
| leader | {{ leader(1, home_runs).player.season(2023).location(away).full_name }} | The league leader in home runs on the road in 2023 | MLB |
| leader | {{ leader(1, home_runs).player.season(2023).location(away).home_runs }} | The amount of HRs the league leader in home runs on the road in 2023 hit | MLB |
| leader | {{ leader(1, home_runs).player.season(2023).location(away).team_alias }} | The team alias of the leader in home runs on the road in 2023 | MLB |
| leader | {{ leader(1, team_era).team.season(2023).inning(7-9).altcity }} | The altcity of the team that lead the league in ERA in the 7-9th innings in 2023 | MLB |
| leader | {{ leader(3, team_era).team.season(2023).inning(7-9).alias }} | The alias of the 3rd ranked team in the league in ERA in the 7-9th innings in 2023 | MLB |
| leader | {{ leader(4, team_era).team.season(2023).location(away).team_era }} | The ERA of the 4th ranked team in ERA in the 7-9th innings in 2023 | MLB |
| leader | {{leader(#,measure).entity.filter.attribute_or_stat}} |  | MLB |
| math | {{query.entity.filter.measure 'operator'# 'formatter'}} |  | MLB |
| previous | {{ previous.player(LAA, 27).minimum(walkoff_hits, 1).ab_result }} | The result of Mike Trout's last walk-off hit | MLB |
| previous | {{ previous.team(MIL).minimum(runs_allowed, 20).game_date \| long_year }} | Last time the Brewers allowed 20 runs in a game | MLB |
| previous | {{ previous.team_player(LAD).minimum(home_runs_grand_slam, 1).full_name }} | The name of the last Dodger to hit a grand slam | MLB |
| previous | {{previous.entity.filter.measure}} |  | MLB |
| rank | {{ rank(asc).player(TB, 56).season(2023).is_qualified_hitting.strikeouts \| ordinal }} | Where Randy Arozarena ranked among qualified hitters in strikeouts in 2023 from least to most - overrides default sort order set in the Stat Defaults menu | MLB |
| rank | {{ rank(desc).player.season(2016, 2017, 2018, 2019, 2020, 2021, 2022, 2023).location(home).is_qualified_pitching.pitcher_era \| ordinal }} | Where a player ranked among qualified pitchers in ERA at home from 2016-2023 from highest to lowest - overrides default sort order set in the Stat Defaults menu | MLB |
| rank | {{ rank.team(LAD).season.innings(extra).asRHB.exit_velocity_ground_balls_average \| ordinal }} | Where the Dodgers rank in average exit velocity of right-handed batters in extra innings this season | MLB |
| rank | {{ rank.team.postseason(last10).game_extra_innings(no).whip_road }} | Where a team ranks in WHIP on the road in non-extra innings games in the playoffs over the last 10 years (Ranks for coaches are not available at the moment) | MLB |
| rank | {{rank.entity.filter.measure}} |  | MLB |
| stats | {{ stats.away.season(2015).venue(DET).extra_inning_games(yes).team_wins }} | Away Team's Extra Inning Game Wins at Detroit in the 2015 Season | MLB |
| stats | {{ stats.coach.season(last5).all_star_break(after).team_wins }} | Coach's Team Wins in the Last 5 Seasons after the All Star Break | MLB |
| stats | {{ stats.home.season.last_game(5).location(home).team_wins }} | Teams Wins in the Last 5 Games At Home This Season | MLB |
| stats | {{ stats.player.career.month(april).innings(7-9).hits }} | Player's Career Hits in April in Innings 7 to 9 | MLB |
| stats | {{ stats.team(DET).season(prev2).vs(HOU).team_wins }} | Detroit's Wins vs. Houston in the Last 2 Seasons | MLB |
| stats | {{ stats.us.postseason.team_wins }} | My Team's Wins in the Post Season | MLB |
| stats | {{stats.entity.filter.measure}} |  | MLB |
| streak | {{ streak(quality_starts, 1).player(ATL, 56).season(2024).equal(games_started, 1).games_played }} | Spencer Schwellenbach number of quality starts in a row | MLB |
| streak | {{ streak(runs, >3).team(BAL).season(2024).games_played }} | The number of times in a row that the Orioles have scored at least 3 runs in a game | MLB |
| streak | {{streak.entity.filter.measure}} |  | MLB |
| calendar | {{ calendar(1, sun).team.game_vs }} |  | NHL |
| calendar | {{ calendar(1, sun).time.day_of_month }} |  | NHL |
| calendar | {{ calendar(2, mon).team.opp_alias }} |  | NHL |
| calendar | {{ calendar(2, tue).team.opp_name }} |  | NHL |
| calendar | {{ calendar(3, thu).team.runs }} |  | NHL |
| calendar | {{ calendar(3, thu).team.venue_type }} | Will return 0 for no game, 1 for away games, and 2 for home games | NHL |
| calendar | {{ calendar(3, wed).team.game_result }} |  | NHL |
| calendar | {{ calendar(4, fri).team.runs_allowed }} |  | NHL |
| calendar | {{ calendar(5, sat).time.day_of_month }} |  | NHL |
| calendar | {{calendar(#,day).entity.filter.measure}} |  | NHL |
| conditional | {{ %if [query] = [query or value] %then [query or value] %else [query or value] %endif }} |  | NHL |
| custom | {{custom.user-dictionary-variable}} |  | NHL |
| game high | {{game_high.entity.filter.measure}} |  | NHL |
| game_high | {{ game_high(at_bats).player(LAD, 5).season(last5).at_bats }} | The most at bats Freddie Freeman has had in a game in the last 5 seasons | NHL |
| game_high | {{ game_high(runs).team(SF).all_time(REG, 1876).runs }} | The most runs the Giants have ever scored in a game | NHL |
| games with | {{games_with.entity.filter.measure}} |  | NHL |
| games_with | {{ games_with.player(NYY, 99).season(2024).minimum(home_runs, 2).games_played }} | The amount of multi-home run Aaron Judge has this season | NHL |
| games_with | {{ games_with.team(LAD).season(2024).minimum(runs, 10).home_runs }} | The amount of overall home runs the Dodgers have hit in the games that they have scored at least 10 runs this season | NHL |
| info | {{ info.coach(DET).first_name }} | First Name of Detroit's Coach | NHL |
| info | {{ info.league.alias }} | League Name Abbreviation | NHL |
| info | {{ info.player.full_name }} | Player's Full Name | NHL |
| info | {{ info.team.name }} | Team's Name | NHL |
| info | {{ info.time.month(prev) }} | Previous Month | NHL |
| info | {{ info.time.season(last) }} | Last Season | NHL |
| info | {{info.entity.attribute}} |  | NHL |
| leader | {{ leader(1, fielding_steals_against, desc).team.season.alias }} | The alias of the team that allowed the most steals this season (overrides the "asc" sort order in the Stat Defaults menu) | NHL |
| leader | {{ leader(1, home_runs).player.season(2023).location(away).full_name }} | The league leader in home runs on the road in 2023 | NHL |
| leader | {{ leader(1, home_runs).player.season(2023).location(away).home_runs }} | The amount of HRs the league leader in home runs on the road in 2023 hit | NHL |
| leader | {{ leader(1, home_runs).player.season(2023).location(away).team_alias }} | The team alias of the leader in home runs on the road in 2023 | NHL |
| leader | {{ leader(1, team_era).team.season(2023).inning(7-9).altcity }} | The altcity of the team that lead the league in ERA in the 7-9th innings in 2023 | NHL |
| leader | {{ leader(3, team_era).team.season(2023).inning(7-9).alias }} | The alias of the 3rd ranked team in the league in ERA in the 7-9th innings in 2023 | NHL |
| leader | {{ leader(4, team_era).team.season(2023).location(away).team_era }} | The ERA of the 4th ranked team in ERA in the 7-9th innings in 2023 | NHL |
| leader | {{leader(#,measure).entity.filter.attribute_or_stat}} |  | NHL |
| math | {{query.entity.filter.measure 'operator'# 'formatter'}} |  | NHL |
| previous | {{ previous.player(LAA, 27).minimum(walkoff_hits, 1).ab_result }} | The result of Mike Trout's last walk-off hit | NHL |
| previous | {{ previous.team(MIL).minimum(runs_allowed, 20).game_date \| long_year }} | Last time the Brewers allowed 20 runs in a game | NHL |
| previous | {{ previous.team_player(LAD).minimum(home_runs_grand_slam, 1).full_name }} | The name of the last Dodger to hit a grand slam | NHL |
| previous | {{previous.entity.filter.measure}} |  | NHL |
| rank | {{ rank(asc).player(TB, 56).season(2023).is_qualified_hitting.strikeouts \| ordinal }} | Where Randy Arozarena ranked among qualified hitters in strikeouts in 2023 from least to most - overrides default sort order set in the Stat Defaults menu | NHL |
| rank | {{ rank(desc).player.season(2016, 2017, 2018, 2019, 2020, 2021, 2022, 2023).location(home).is_qualified_pitching.pitcher_era \| ordinal }} | Where a player ranked among qualified pitchers in ERA at home from 2016-2023 from highest to lowest - overrides default sort order set in the Stat Defaults menu | NHL |
| rank | {{ rank.team(LAD).season.innings(extra).asRHB.exit_velocity_ground_balls_average \| ordinal }} | Where the Dodgers rank in average exit velocity of right-handed batters in extra innings this season | NHL |
| rank | {{ rank.team.postseason(last10).game_extra_innings(no).whip_road }} | Where a team ranks in WHIP on the road in non-extra innings games in the playoffs over the last 10 years (Ranks for coaches are not available at the moment) | NHL |
| rank | {{rank.entity.filter.measure}} |  | NHL |
| stats | {{ stats.away.season(2015).venue(DET).extra_inning_games(yes).team_wins }} | Away Team's Extra Inning Game Wins at Detroit in the 2015 Season | NHL |
| stats | {{ stats.coach.season(last5).all_star_break(after).team_wins }} | Coach's Team Wins in the Last 5 Seasons after the All Star Break | NHL |
| stats | {{ stats.home.season.last_game(5).location(home).team_wins }} | Teams Wins in the Last 5 Games At Home This Season | NHL |
| stats | {{ stats.player.career.month(april).innings(7-9).hits }} | Player's Career Hits in April in Innings 7 to 9 | NHL |
| stats | {{ stats.team(DET).season(prev2).vs(HOU).team_wins }} | Detroit's Wins vs. Houston in the Last 2 Seasons | NHL |
| stats | {{ stats.us.postseason.team_wins }} | My Team's Wins in the Post Season | NHL |
| stats | {{stats.entity.filter.measure}} |  | NHL |
| streak | {{ streak(quality_starts, 1).player(ATL, 56).season(2024).equal(games_started, 1).games_played }} | Spencer Schwellenbach number of quality starts in a row | NHL |
| streak | {{ streak(runs, >3).team(BAL).season(2024).games_played }} | The number of times in a row that the Orioles have scored at least 3 runs in a game | NHL |
| streak | {{streak.entity.filter.measure}} |  | NHL |

## SmartStat Relevance

- Separates baseline semantic skeleton interpretation from workbook query category labels.
- Preserves complete workbook query evidence for deterministic semantic/planner analysis.
- Supports candidate-resolution explainability and future grammar formalization without asserting runtime behavior.
