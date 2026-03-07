# OnAir Measure Dictionary

## Purpose / Overview

Reference list of measure tokens from category/measure sheets and additional measure sheets.

## Source workbook coverage

- Workbook(s): MLB_OnAir_v3_Stat_Syntax_v2.xlsx, NHL_OnAir_v3_Stat_Syntax_v2.xlsx
- Sheets used: CATEGORY_TO_MEASURE (1), CATEGORY_TO_MEASURE_PITCHER (2), CATEGORY_TO_MEASURE (3), CATEGORY_TO_MEASURE_GOALIE (4), Additional Measures (9), Additional Measures (9)

## Extraction notes / normalization notes

- Header normalization: `VALUE`, `SYNTAX`, and `Measure` -> `measure_name`.
- Subtype derives from sheet role (`general`, `pitcher`, `goalie`, `additional`).

## Extracted reference

| measure_name | description | league | subtype | source_sheet | notes |
| --- | --- | --- | --- | --- | --- |
| pitcher_record |  | MLB | additional | Additional Measures |  |
| team_record |  | MLB | additional | Additional Measures |  |
| air_balls | Batters, number of balls in play that were hit in the air. (Fly Balls + Line Drives + Pop-Ups) | MLB | general | CATEGORY_TO_MEASURE |  |
| air_balls_percentage | Batters, percentage of balls in play that were hit in the air. (Fly Balls + Line Drives + Pop-Ups) | MLB | general | CATEGORY_TO_MEASURE |  |
| arsenal_changeups | Pitcher, number of changeups thrown | MLB | general | CATEGORY_TO_MEASURE |  |
| arsenal_changeups_percentage | Pitcher, percentage of total pitches thrown that were changeups (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |  |
| arsenal_curveballs | Pitcher, number of curveballs thrown | MLB | general | CATEGORY_TO_MEASURE |  |
| arsenal_curveballs_percentage | Pitcher, percentage of total pitches thrown that were curveballs (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |  |
| arsenal_cutters | Pitcher, number of cutters thrown (aka cut fastballs) | MLB | general | CATEGORY_TO_MEASURE |  |
| arsenal_cutters_percentage | Pitcher, percentage of total pitches thrown that were cutters (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |  |
| arsenal_eephuses | Pitcher, number of pitches thrown classified as Eephus | MLB | general | CATEGORY_TO_MEASURE |  |
| arsenal_eephuses_percentage | Pitcher, percentage of total pitches thrown that were eephus pitches (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |  |
| arsenal_fastballs | Pitcher, number of 4-seam fastballs thrown | MLB | general | CATEGORY_TO_MEASURE |  |
| arsenal_fastballs_percentage | Pitcher, percentage of total pitches thrown that were 4-seam fastballs (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |  |
| arsenal_forkballs | Pitcher, number of forkballs thrown | MLB | general | CATEGORY_TO_MEASURE |  |
| arsenal_forkballs_percentage | Pitcher, percentage of total pitches thrown that were forkballs (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |  |
| arsenal_knuckleballs | Pitcher, number of knuckleballs thrown | MLB | general | CATEGORY_TO_MEASURE |  |
| arsenal_knuckleballs_percentage | Pitcher, percentage of total pitches thrown that were knuckleballs (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |  |
| arsenal_knucklecurves | Pitcher, number of pitches thrown classified as knuckle curve | MLB | general | CATEGORY_TO_MEASURE |  |
| arsenal_knucklecurves_percentage | Pitcher, percentage of total pitches thrown that were knuckle curves (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |  |
| arsenal_screwballs | Pitcher, number of screwballs thrown | MLB | general | CATEGORY_TO_MEASURE |  |
| arsenal_screwballs_percentage | Pitcher, percentage of total pitches thrown that were screwballs (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |  |
| arsenal_sinkers | Pitcher, number of 2-seam fastballs thrown (aka sinkers) | MLB | general | CATEGORY_TO_MEASURE |  |
| arsenal_sinkers_percentage | Pitcher, percentage of total pitches thrown that were 2-seam fastballs (aka sinkers) (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |  |
| arsenal_sliders | Pitcher, number of sliders thrown | MLB | general | CATEGORY_TO_MEASURE |  |
| arsenal_sliders_percentage | Pitcher, percentage of total pitches thrown that were sliders (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |  |
| arsenal_slurves | Pitcher, number of slurves thrown | MLB | general | CATEGORY_TO_MEASURE |  |
| arsenal_slurves_percentage | Pitcher, percentage of total pitches thrown that were slurves (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |  |
| arsenal_splitters | Pitcher, number of splitters thrown | MLB | general | CATEGORY_TO_MEASURE |  |
| arsenal_splitters_percentage | Pitcher, percentage of total pitches thrown that were splitters (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |  |
| arsenal_sweepers | Pitcher, number of sweepers thrown | MLB | general | CATEGORY_TO_MEASURE |  |
| arsenal_sweepers_percentage | Pitcher, percentage of total pitches thrown that were sweepers (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |  |
| assists_outfield | Fielding, Outfield Assists | MLB | general | CATEGORY_TO_MEASURE |  |
| at_bats | Batters, number of at-bats | MLB | general | CATEGORY_TO_MEASURE |  |
| at_bats_bases_empty | At-Bats with no runners on base. | MLB | general | CATEGORY_TO_MEASURE |  |
| at_bats_bases_loaded | At-bats with the bases loaded. | MLB | general | CATEGORY_TO_MEASURE |  |
| at_bats_home | Batters, number of at-bats in home games | MLB | general | CATEGORY_TO_MEASURE |  |
| at_bats_per_home_run | Batters, Ratio of AB to HR (lower number is better, meaning a higher rate of HR hit) | MLB | general | CATEGORY_TO_MEASURE |  |
| at_bats_per_rbi | Batters, Ratio of AB to RBI (lower number is better, meaning a higher rate of RBI hit) | MLB | general | CATEGORY_TO_MEASURE |  |
| at_bats_per_strikeout | Batters, Ratio of AB to K (higher number is better, meaning a lower rate of striking out) | MLB | general | CATEGORY_TO_MEASURE |  |
| at_bats_road | Batters, number of at-bats in road games | MLB | general | CATEGORY_TO_MEASURE |  |
| at_bats_runners | At-bats with runners on base. | MLB | general | CATEGORY_TO_MEASURE |  |
| at_bats_runners_in_scoring_position | At-Bats with Runners in Scoring Position | MLB | general | CATEGORY_TO_MEASURE |  |
| at_bats_runners_in_scoring_position_two_outs | At-Bats with Runners in Scoring Position and two outs. | MLB | general | CATEGORY_TO_MEASURE |  |
| b_hr_multi_pct | Batters, Pct. of Home Runs hit that had men on base. | MLB | general | CATEGORY_TO_MEASURE |  |
| b_hr_multirun | Batters, HR hit with men on base | MLB | general | CATEGORY_TO_MEASURE |  |
| balks | Pitcher Balks | MLB | general | CATEGORY_TO_MEASURE |  |
| balls_in_field | Batters, number of balls in play (all balls in play minus HR) | MLB | general | CATEGORY_TO_MEASURE |  |
| balls_in_play | Batters, number of balls in play (all balls in play including HR, SacBunts, SacFlies, etc...) | MLB | general | CATEGORY_TO_MEASURE |  |
| balls_in_play_per_game | Batters, number of balls in play per game. | MLB | general | CATEGORY_TO_MEASURE |  |
| batted_balls_events | The number of balls that were put in play. | MLB | general | CATEGORY_TO_MEASURE |  |
| batter_chase_percentage | (also known as Out-of-Zone Swing Pct.) Batters, swing percentage on pitches out of the zone. (Zones as defined by MLB, 1->9 is in-zone, 11->14 is out-of-zone) | MLB | general | CATEGORY_TO_MEASURE |  |
| batter_contact_percentage | Batters, percentage of swings where contact was made (swings with contact / swings) | MLB | general | CATEGORY_TO_MEASURE |  |
| batter_foul_balls | Number of foul balls a batter or batting team hit | MLB | general | CATEGORY_TO_MEASURE |  |
| batter_home_run_to_fly_ball_percentage | Batters, Pct. of fly balls + line drives that were Home Runs | MLB | general | CATEGORY_TO_MEASURE |  |
| batter_inzone_contact_percentage | Batters, percentage of swings where contact was made on pitches thrown in the zone. (Zones as defined by MLB, 1->9 is in-zone, 11->14 is out-of-zone) | MLB | general | CATEGORY_TO_MEASURE |  |
| batter_inzone_swing_percentage | Batters, swing percentage on pitches which were thrown in the zone. (Zones as defined by MLB, 1->9 is in-zone, 11->14 is out-of-zone) | MLB | general | CATEGORY_TO_MEASURE |  |
| batter_inzone_whiff_percentage | Batters, whiff percentage on pitches thrown in the zone. (Zones as defined by MLB, 1->9 is in-zone, 11->14 is out-of-zone) | MLB | general | CATEGORY_TO_MEASURE |  |
| batter_outofzone_contact_percentage | Batters, percentage of swings where contact was made on pitches thrown out of the zone. (Zones as defined by MLB, 1->9 is in-zone, 11->14 is out-of-zone) | MLB | general | CATEGORY_TO_MEASURE |  |
| batter_outofzone_whiff_percentage | Batters, whiff percentage on pitches thrown out of the zone. (Zones as defined by MLB, 1->9 is in-zone, 11->14 is out-of-zone) | MLB | general | CATEGORY_TO_MEASURE |  |
| batter_percentage_team_runs | Players percentage of team's runs. (Player R / Team R) - NEEDS PLAYER AND TEAM GROUPING INCLUDED IN REPORTS | MLB | general | CATEGORY_TO_MEASURE |  |
| batter_percentage_team_walks | Batters, percentage of Teams Total Walks - NEEDS PLAYER AND TEAM GROUPING INCLUDED IN REPORTS | MLB | general | CATEGORY_TO_MEASURE |  |
| batter_percentage_teams_ab | Players percentage of team's at-bats. (Player AB / Team AB) - NEEDS PLAYER AND TEAM GROUPING INCLUDED IN REPORTS | MLB | general | CATEGORY_TO_MEASURE |  |
| batter_percentage_teams_extra_base_hits | Players percentage of team's extra base hits. (Player XBH / Team XBH) - NEEDS PLAYER AND TEAM GROUPING INCLUDED IN REPORTS | MLB | general | CATEGORY_TO_MEASURE |  |
| batter_percentage_teams_hits | Players percentage of team's hits. (Player H / Team H) - NEEDS PLAYER AND TEAM GROUPING INCLUDED IN REPORTS | MLB | general | CATEGORY_TO_MEASURE |  |
| batter_percentage_teams_home_runs | Players percentage of team's home runs. (Player HR / Team HR) - NEEDS PLAYER AND TEAM GROUPING INCLUDED IN REPORTS | MLB | general | CATEGORY_TO_MEASURE |  |
| batter_percentage_teams_pa | Players percentage of team's plate appearances. (Player PA / Team PA) - NEEDS PLAYER AND TEAM GROUPING INCLUDED IN REPORTS | MLB | general | CATEGORY_TO_MEASURE |  |
| batter_percentage_teams_rbis | Players percentage of team's RBI. (Player RBI / Team RBI) - NEEDS PLAYER AND TEAM GROUPING INCLUDED IN REPORTS | MLB | general | CATEGORY_TO_MEASURE |  |
| batter_percentage_teams_strikeouts | Players percentage of team's strikeouts batting. (Player SO / Team SO) - NEEDS PLAYER AND TEAM GROUPING INCLUDED IN REPORTS | MLB | general | CATEGORY_TO_MEASURE |  |
| batter_swing_percentage | Batters, percentage of pitches that are swung at. | MLB | general | CATEGORY_TO_MEASURE |  |
| batter_swings | Batters, number of times swinging at the pitch | MLB | general | CATEGORY_TO_MEASURE |  |
| batter_swings_contact | Batters, number of times swinging at the pitch and making contact | MLB | general | CATEGORY_TO_MEASURE |  |
| batter_times_chased | Batters, number of times swinging at a pitch out of the zone. (Zones as defined by MLB, 1->9 is in-zone, 11->14 is out-of-zone) | MLB | general | CATEGORY_TO_MEASURE |  |
| batter_whiff_percentage | Batters, percentage of swings where contact was not made (swing+miss / swings) | MLB | general | CATEGORY_TO_MEASURE |  |
| batter_whiffs | Batters, number of times swinging at the pitch without making contact | MLB | general | CATEGORY_TO_MEASURE |  |
| batters_faced | Pitchers number of batters faced | MLB | general | CATEGORY_TO_MEASURE |  |
| batting_average | Batting Average (hits / at-bats) | MLB | general | CATEGORY_TO_MEASURE |  |
| batting_average_balls_in_play | Batters, batting average on balls in play. (H - HR) / (AB - K - HR + SF) | MLB | general | CATEGORY_TO_MEASURE |  |
| batting_average_balls_in_play_number | Batters, number of "balls in play", as a result of SacFlies or AB only, which have a chance to be fielded (AB - HR - K + SacF) | MLB | general | CATEGORY_TO_MEASURE |  |
| batting_average_bases_empty | Batters, batting average with no runners on base. | MLB | general | CATEGORY_TO_MEASURE |  |
| batting_average_bases_loaded | Batting average when the bases are loaded. | MLB | general | CATEGORY_TO_MEASURE |  |
| batting_average_differential | Teams, difference between batting average and opponent batting average | MLB | general | CATEGORY_TO_MEASURE |  |
| batting_average_home | Batters, batting average in home games | MLB | general | CATEGORY_TO_MEASURE |  |
| batting_average_home_road_differential | Difference between batting average in home games vs. road games, negative indicates better on the road | MLB | general | CATEGORY_TO_MEASURE |  |
| batting_average_lhp | Batting average vs. left-handed pitchers | MLB | general | CATEGORY_TO_MEASURE |  |
| batting_average_lhp_rhp_differential | Difference between batting average against left-handed pitchers vs. right-handed pitchers, negative indicates better vs. right-handed pitchers. | MLB | general | CATEGORY_TO_MEASURE |  |
| batting_average_rhp | Batting average vs. right-handed pitchers | MLB | general | CATEGORY_TO_MEASURE |  |
| batting_average_road | Batters, batting average in road games | MLB | general | CATEGORY_TO_MEASURE |  |
| batting_average_runners_in_scoring_position | Batting Average (hits / at-bats) with Runners in Scoring position. | MLB | general | CATEGORY_TO_MEASURE |  |
| batting_average_runners_in_scoring_position_two_outs | Batting Average (hits / at-bats) with runners in scoring position and two outs. | MLB | general | CATEGORY_TO_MEASURE |  |
| batting_average_two_outs | Batting Average (hits / at-bats) with two outs. | MLB | general | CATEGORY_TO_MEASURE |  |
| batting_averages_runners_on_base | Batting average with runners on base. | MLB | general | CATEGORY_TO_MEASURE |  |
| batting_games_played | Players, number of games appearing as a batter | MLB | general | CATEGORY_TO_MEASURE |  |
| batting_summary | Doubles and Triples are spelled out | MLB | general | CATEGORY_TO_MEASURE |  |
| batting_summary_abb | 2B and 3B for Doubles and Triples | MLB | general | CATEGORY_TO_MEASURE |  |
| batting_summary_abb_noruns | 2B and 3B for Doubles and Triples (Does not include Runs) | MLB | general | CATEGORY_TO_MEASURE |  |
| batting_summary_noruns | Doubles and Triples are spelled out (Does not include Runs) | MLB | general | CATEGORY_TO_MEASURE |  |
| bet_t_w_fav | Games won where the team was favored based on the pre-game moneyline odds | MLB | general | CATEGORY_TO_MEASURE |  |
| bet_t_w_udog | Wins as the moneyline underdog | MLB | general | CATEGORY_TO_MEASURE |  |
| called_strike | Total number of called strikes plus whiffs divided by total pitches. | MLB | general | CATEGORY_TO_MEASURE |  |
| catcher_passed_balls | Catchers, number of passed balls allowed while catching | MLB | general | CATEGORY_TO_MEASURE |  |
| caught_stealing | Runners, Times caught stealing when attempting to steal a base | MLB | general | CATEGORY_TO_MEASURE |  |
| combined_extra_base_hits | Combined extra base hits (doubles/triples/home runs) between team and opponent | MLB | general | CATEGORY_TO_MEASURE |  |
| combined_extra_base_hits_per_game | Combined extra base hits (doubles/triples/home runs) between team and opponent, per game. | MLB | general | CATEGORY_TO_MEASURE |  |
| combined_hits | Combined Hits between Team and Opponent | MLB | general | CATEGORY_TO_MEASURE |  |
| combined_hits_per_game | Combined Hits between Team and Opponent, per game. | MLB | general | CATEGORY_TO_MEASURE |  |
| combined_home_runs | Combined total Home Runs between Team and Opponent. | MLB | general | CATEGORY_TO_MEASURE |  |
| combined_home_runs_per_game | Combined Home Runs between Team and Opponent, per game. | MLB | general | CATEGORY_TO_MEASURE |  |
| combined_runs | Combined Runs between Team and Opponent. | MLB | general | CATEGORY_TO_MEASURE |  |
| combined_runs_per_game | Combined Runs Scored between team and opponent, per game (include Team Grouping/Split) | MLB | general | CATEGORY_TO_MEASURE |  |
| combined_strikeouts | Combined strikeouts (thrown and from batters) by both teams. | MLB | general | CATEGORY_TO_MEASURE |  |
| combined_strikeouts_per_game | Combined strikeouts (thrown and from batters) by both teams, per game | MLB | general | CATEGORY_TO_MEASURE |  |
| doubles | Batters number of hits that were doubles | MLB | general | CATEGORY_TO_MEASURE |  |
| doubles_per_game | Doubles per game | MLB | general | CATEGORY_TO_MEASURE |  |
| enforced_balls_hitter | The number of times a hitter was the beneficiary of an enforced ball | MLB | general | CATEGORY_TO_MEASURE |  |
| enforced_balls_pitcher | The number of times a pitcher was the victim of an enforced ball | MLB | general | CATEGORY_TO_MEASURE |  |
| enforced_strikes_hitter | The number of times a hitter was the victim of an enforced strike | MLB | general | CATEGORY_TO_MEASURE |  |
| enforced_strikes_pitcher | The number of times a pitcher was the beneficiary of an enforced strike | MLB | general | CATEGORY_TO_MEASURE |  |
| errors_fielding | Fielders, number of errors committed classified as "fielding errors" (NOT throwing or interference errors) | MLB | general | CATEGORY_TO_MEASURE |  |
| errors_interference | Fielders, number of errors committed classified as "interference errors" (NOT fielding or throwing errors) | MLB | general | CATEGORY_TO_MEASURE |  |
| errors_throwing | Fielders, number of errors committed classified as throwing errors (NOT fielding or interference errors) | MLB | general | CATEGORY_TO_MEASURE |  |
| errors_total | Fielders, number of errors committed (fielding + throwing + interference) | MLB | general | CATEGORY_TO_MEASURE |  |
| exit_velocity_average | Batters, average exit velocity on balls in play (among pitches with recorded exit velocity only) | MLB | general | CATEGORY_TO_MEASURE |  |
| exit_velocity_fly_balls_line_drives_average | Batters, average exit velocity on flyballs and linedrives (among pitches with recorded exit velocity only) | MLB | general | CATEGORY_TO_MEASURE |  |
| exit_velocity_ground_balls_average | Batters, average exit velocity on ground balls (among pitches with recorded exit velocity only) | MLB | general | CATEGORY_TO_MEASURE |  |
| extra_base_hits | Batters, number of hits for extra bases (2B + 3B + HR) | MLB | general | CATEGORY_TO_MEASURE |  |
| extra_base_hits_differential | Teams, difference between batting XBH and pitching XBH allowed | MLB | general | CATEGORY_TO_MEASURE |  |
| extra_base_hits_per_game | The number of extra-base hits per game. | MLB | general | CATEGORY_TO_MEASURE |  |
| fantasy_points_draftkings_batting | Draftkings MLB Batting Fantasy Points | MLB | general | CATEGORY_TO_MEASURE |  |
| fantasy_points_draftkings_pitching | Draftkings MLB Pitching Fantasy Points | MLB | general | CATEGORY_TO_MEASURE |  |
| fantasy_points_per_game_draftkings_batting | Draftkings MLB Batting Fantasy Points per Game | MLB | general | CATEGORY_TO_MEASURE |  |
| fantasy_points_per_game_draftkings_pitching | Draftkings MLB Pitching Fantasy Points per Game | MLB | general | CATEGORY_TO_MEASURE |  |
| fielders_choice | Batters, number of times hitting into fielder's choice | MLB | general | CATEGORY_TO_MEASURE |  |
| fielding_assists | Fielders, number of assists | MLB | general | CATEGORY_TO_MEASURE |  |
| fielding_caught_stealing | Number of opposing runners caught stealing | MLB | general | CATEGORY_TO_MEASURE |  |
| fielding_caught_stealing_against_percentage | Percentage of opposing runners caught stealing | MLB | general | CATEGORY_TO_MEASURE |  |
| fielding_double_plays | Fielders, number of double-plays involved in | MLB | general | CATEGORY_TO_MEASURE |  |
| fielding_games_played | Players, number of games played as a fielder | MLB | general | CATEGORY_TO_MEASURE |  |
| fielding_innings_played | Innings played in the field (calculated from outs while fielding) | MLB | general | CATEGORY_TO_MEASURE |  |
| fielding_percentage | Fielding Percentage. Putouts+Assists/Total Chances (PO+Assists+Errors) | MLB | general | CATEGORY_TO_MEASURE |  |
| fielding_putouts | Fielders, number of putouts | MLB | general | CATEGORY_TO_MEASURE |  |
| fielding_steals_against | Number of stolen bases against (opponent stolen bases) | MLB | general | CATEGORY_TO_MEASURE |  |
| fielding_steals_attempts_against | Number of stolen base attempts against (opponent's Steals + Caught Stealing) | MLB | general | CATEGORY_TO_MEASURE |  |
| fielding_stolen_base_percentage_against | Percentage of successful stolen base attempts against | MLB | general | CATEGORY_TO_MEASURE |  |
| fielding_total_chances | Players, total fielding chances (Putouts + Assists + Errors) | MLB | general | CATEGORY_TO_MEASURE |  |
| fielding_triple_plays | Fielders, number of triple-plays involved in | MLB | general | CATEGORY_TO_MEASURE |  |
| fielding_wild_pitches | Catchers, number of wild pitches thrown while at catcher | MLB | general | CATEGORY_TO_MEASURE |  |
| first_pitch_balls_seen | Batters, number of plate appearances start with a ball called | MLB | general | CATEGORY_TO_MEASURE |  |
| first_pitch_balls_thrown | Pitchers, first pitch of plate appearance called a ball | MLB | general | CATEGORY_TO_MEASURE |  |
| first_pitch_hits | Batters, hits on the first pitch of an at-bat | MLB | general | CATEGORY_TO_MEASURE |  |
| first_pitch_hits_allowed | Pitchers, hits allowed on the first pitch of an at-bat | MLB | general | CATEGORY_TO_MEASURE |  |
| first_pitch_home_run_allowed | Pitchers, home runs allowed on the first pitch of an at-bat | MLB | general | CATEGORY_TO_MEASURE |  |
| first_pitch_home_run_hit | Batters, home run on first pitch of at-bat | MLB | general | CATEGORY_TO_MEASURE |  |
| first_pitch_strikes_seen | Batters, number of strikes on the first pitch of a plate appearance | MLB | general | CATEGORY_TO_MEASURE |  |
| first_pitch_strikes_seen_percentage | Batters, percentage of plate appearances to start with a first pitch strike | MLB | general | CATEGORY_TO_MEASURE |  |
| first_pitch_strikes_thrown | Pitchers, strikes thrown on the first pitch of a plate appearance | MLB | general | CATEGORY_TO_MEASURE |  |
| first_pitch_strikes_thrown_percentage | Pitchers, percent of plate appearances to start with a strike | MLB | general | CATEGORY_TO_MEASURE |  |
| first_pitch_swing_percentage | Batters, percentage of first pitches swung at | MLB | general | CATEGORY_TO_MEASURE |  |
| first_pitch_swing_percentage_against | Pitchers, Percentage of swings against on the first pitch of a plate appearance | MLB | general | CATEGORY_TO_MEASURE |  |
| first_pitch_swings | Batters, swings on the first pitch of an at-bat | MLB | general | CATEGORY_TO_MEASURE |  |
| first_pitch_swings_agaist | Pitchers, swings against on the first pitch of an at-bat | MLB | general | CATEGORY_TO_MEASURE |  |
| fly_ball_outs | Batters, number of times out on a fly ball. (Note: does not include line drives or pop-ups) | MLB | general | CATEGORY_TO_MEASURE |  |
| fly_balls | Batters, number of balls in play that were fly balls. (Note: does not include line drives or pop-ups) | MLB | general | CATEGORY_TO_MEASURE |  |
| fly_balls_percentage | Batters, percentage of balls in play that were Fly Balls (Note: does not include line drives or pop-ups) | MLB | general | CATEGORY_TO_MEASURE |  |
| fly_pop_balls | Batters, number balls in play that were either Fly Balls or Pop-Ups. (Note: does not include Line Drives) | MLB | general | CATEGORY_TO_MEASURE |  |
| fly_pop_balls_percentage | Batters, percentage of balls in play that were either Fly Balls or Pop-Ups. (Note: does not include Line Drives) | MLB | general | CATEGORY_TO_MEASURE |  |
| game_score | Metric devised by Bill James to determine the strength of a pitcher in any particular baseball game | MLB | general | CATEGORY_TO_MEASURE |  |
| game_winning_rbi | Batters, RBI that gives a team the lead it never relinquishes | MLB | general | CATEGORY_TO_MEASURE |  |
| gamelog_batting_average | Windowed batting average based on the level of stats that are being viewed. For example, when viewing by games, this is the batting average at that point in the season. | MLB | general | CATEGORY_TO_MEASURE |  |
| gamelog_era | Windowed Earned Run Average based on the level of stats that are being viewed. For example, when viewing by games, this is the earned run average at that point in the season. | MLB | general | CATEGORY_TO_MEASURE |  |
| gamelog_era_reliever | Windowed Earned Run Average for relievers based on the level of stats that are being viewed. For example, when viewing by games, this is the earned run average at that point in the season. | MLB | general | CATEGORY_TO_MEASURE |  |
| gamelog_era_starter | Windowed Earned Run Average for starters based on the level of stats that are being viewed. For example, when viewing by games, this is the earned run average at that point in the season. | MLB | general | CATEGORY_TO_MEASURE |  |
| gamelog_era_team | Windowed Team Earned Run Average based on the level of stats that are being viewed. For example, when viewing by games, this is the earned run average at that point in the season. | MLB | general | CATEGORY_TO_MEASURE |  |
| gamelog_on_base_percentage | Windowed On-Base Percentage based on the level of stats that are being viewed. For example, when viewing by games, this is the on-base percentage at that point in the season. | MLB | general | CATEGORY_TO_MEASURE |  |
| gamelog_on_base_plus_slugging | Windowed On-Base Plus Slugging Percentage based on the level of stats that are being viewed. For example, when viewing by games, this is the on-base plus slugging percentage at that point in the season. | MLB | general | CATEGORY_TO_MEASURE |  |
| gamelog_slugging | Windowed slugging percentage based on the level of stats that are being viewed. For example, when viewing by games, this is the slugging percentage at that point in the season. | MLB | general | CATEGORY_TO_MEASURE |  |
| gamelog_team_losses | Running Total of Team Losses based on the level of stats that are being viewed. For example, when viewing by games, this is the team losses at that point in the season. | MLB | general | CATEGORY_TO_MEASURE |  |
| gamelog_team_winning_percentage | Running Total of Team Winning Percentage based on the level of stats that are being viewed. For example, when viewing by games, this is the team win pct. at that point in the season. | MLB | general | CATEGORY_TO_MEASURE |  |
| gamelog_team_wins | Running Total of Team Wins based on the level of stats that are being viewed. For example, when viewing by games, this is the team wins at that point in the season. | MLB | general | CATEGORY_TO_MEASURE |  |
| gamelog_whip | Windowed Pitcher WHIP based on the level of stats that are being viewed. For example, when viewing by games, this is the pitcher WHIP at that point in the season. | MLB | general | CATEGORY_TO_MEASURE |  |
| games_day | Number of day games played | MLB | general | CATEGORY_TO_MEASURE |  |
| games_favorite | Team Games Played as the Moneyline Favorite | MLB | general | CATEGORY_TO_MEASURE |  |
| games_finishes | Relief Pitchers, number of games finished (note: relief pitcher statistic only. Starters are not awarded games finished for complete games) | MLB | general | CATEGORY_TO_MEASURE |  |
| games_night | Number of night games played | MLB | general | CATEGORY_TO_MEASURE |  |
| games_pitched | Number of games appearing as pitcher | MLB | general | CATEGORY_TO_MEASURE |  |
| games_pitched_home | Number of games appearing as pitcher in home games | MLB | general | CATEGORY_TO_MEASURE |  |
| games_pitched_road | Number of games appearing as pitcher on the road | MLB | general | CATEGORY_TO_MEASURE |  |
| games_pitched_total | Total Games Pitched - When grouping by team, this will count every individual player game pitched as one instead of showing only one per game. | MLB | general | CATEGORY_TO_MEASURE |  |
| games_played | Games played | MLB | general | CATEGORY_TO_MEASURE |  |
| games_played_total | Total Games Played - When grouping by team, this will count every individual player game played as one instead of showing the team's games played. | MLB | general | CATEGORY_TO_MEASURE |  |
| games_started | Number of games starting at pitcher | MLB | general | CATEGORY_TO_MEASURE |  |
| go_ahead_rbi | Batters, RBI that gave their team the lead | MLB | general | CATEGORY_TO_MEASURE |  |
| ground_balls | Batters, number of balls in play that were ground balls. | MLB | general | CATEGORY_TO_MEASURE |  |
| ground_balls_percentage | Batters, percentage of balls in play that were Ground Balls | MLB | general | CATEGORY_TO_MEASURE |  |
| ground_balls_to_air_balls_differential | Batters, Ratio of balls in play, Ground Balls to Air Balls. (Note: Air Balls = balls in play classified as Fly Ball, Line Drive, or Pop-Up) | MLB | general | CATEGORY_TO_MEASURE |  |
| ground_balls_to_fly_pop_balls_differential | Batters, Ratio of balls in play, Ground Balls to Fly (F+P). (Note: Fly = balls in play classified as either Fly Ball or Pop-Up) | MLB | general | CATEGORY_TO_MEASURE |  |
| ground_outs | Batters, number of times grounding out. | MLB | general | CATEGORY_TO_MEASURE |  |
| ground_outs_percentage | Percentage of ground balls that are groundouts. | MLB | general | CATEGORY_TO_MEASURE |  |
| grounded_into_double_play | Batters, times grounding into a double-play | MLB | general | CATEGORY_TO_MEASURE |  |
| gs | Players, number of games started (either fielding, batting, or pitching) | MLB | general | CATEGORY_TO_MEASURE |  |

_Table truncated to first 200 rows for readability. Full extracted rows are available in `docs/onair/_extracted/onair_reference.extracted.json`._

## SmartStat Relevance

- Supports semantic measure dictionaries.
- Supports candidate-resolution explainability with stable measure vocabulary.
- Provides planner grammar measure inputs without runtime coupling.
