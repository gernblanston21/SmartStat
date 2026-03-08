# OnAir Measure Dictionary

## Purpose / Overview

Deterministic measure reference extracted from category-to-measure and additional-measure source sheets.

## Source workbook coverage

- Workbook(s): MLB_OnAir_v3_Stat_Syntax_v3.xlsx, NHL_OnAir_v3_Stat_Syntax_v3.xlsx
- Sheets used: CATEGORY_TO_MEASURE, CATEGORY_TO_MEASURE_PITCHER, CATEGORY_TO_MEASURE_GOALIE, Additional Measures

## Extraction notes / normalization notes

- Header normalization: `Measure Syntax` and `Measure` are normalized to `measure_name`.
- Subtype is assigned from source sheet role: `general`, `pitcher`, `goalie`, `additional`.

## Extracted reference

| measure_name | description | league | subtype | source_sheet |
| --- | --- | --- | --- | --- |
| pitcher_record |  | MLB | additional | Additional Measures |
| team_record |  | MLB | additional | Additional Measures |
| air_balls | Batters, number of balls in play that were hit in the air. (Fly Balls + Line Drives + Pop-Ups) | MLB | general | CATEGORY_TO_MEASURE |
| air_balls_percentage | Batters, percentage of balls in play that were hit in the air. (Fly Balls + Line Drives + Pop-Ups) | MLB | general | CATEGORY_TO_MEASURE |
| arsenal_changeups | Pitcher, number of changeups thrown | MLB | general | CATEGORY_TO_MEASURE |
| arsenal_changeups_percentage | Pitcher, percentage of total pitches thrown that were changeups (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |
| arsenal_curveballs | Pitcher, number of curveballs thrown | MLB | general | CATEGORY_TO_MEASURE |
| arsenal_curveballs_percentage | Pitcher, percentage of total pitches thrown that were curveballs (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |
| arsenal_cutters | Pitcher, number of cutters thrown (aka cut fastballs) | MLB | general | CATEGORY_TO_MEASURE |
| arsenal_cutters_percentage | Pitcher, percentage of total pitches thrown that were cutters (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |
| arsenal_eephuses | Pitcher, number of pitches thrown classified as Eephus | MLB | general | CATEGORY_TO_MEASURE |
| arsenal_eephuses_percentage | Pitcher, percentage of total pitches thrown that were eephus pitches (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |
| arsenal_fastballs | Pitcher, number of 4-seam fastballs thrown | MLB | general | CATEGORY_TO_MEASURE |
| arsenal_fastballs_percentage | Pitcher, percentage of total pitches thrown that were 4-seam fastballs (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |
| arsenal_forkballs | Pitcher, number of forkballs thrown | MLB | general | CATEGORY_TO_MEASURE |
| arsenal_forkballs_percentage | Pitcher, percentage of total pitches thrown that were forkballs (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |
| arsenal_knuckleballs | Pitcher, number of knuckleballs thrown | MLB | general | CATEGORY_TO_MEASURE |
| arsenal_knuckleballs_percentage | Pitcher, percentage of total pitches thrown that were knuckleballs (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |
| arsenal_knucklecurves | Pitcher, number of pitches thrown classified as knuckle curve | MLB | general | CATEGORY_TO_MEASURE |
| arsenal_knucklecurves_percentage | Pitcher, percentage of total pitches thrown that were knuckle curves (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |
| arsenal_screwballs | Pitcher, number of screwballs thrown | MLB | general | CATEGORY_TO_MEASURE |
| arsenal_screwballs_percentage | Pitcher, percentage of total pitches thrown that were screwballs (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |
| arsenal_sinkers | Pitcher, number of 2-seam fastballs thrown (aka sinkers) | MLB | general | CATEGORY_TO_MEASURE |
| arsenal_sinkers_percentage | Pitcher, percentage of total pitches thrown that were 2-seam fastballs (aka sinkers) (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |
| arsenal_sliders | Pitcher, number of sliders thrown | MLB | general | CATEGORY_TO_MEASURE |
| arsenal_sliders_percentage | Pitcher, percentage of total pitches thrown that were sliders (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |
| arsenal_slurves | Pitcher, number of slurves thrown | MLB | general | CATEGORY_TO_MEASURE |
| arsenal_slurves_percentage | Pitcher, percentage of total pitches thrown that were slurves (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |
| arsenal_splitters | Pitcher, number of splitters thrown | MLB | general | CATEGORY_TO_MEASURE |
| arsenal_splitters_percentage | Pitcher, percentage of total pitches thrown that were splitters (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |
| arsenal_sweepers | Pitcher, number of sweepers thrown | MLB | general | CATEGORY_TO_MEASURE |
| arsenal_sweepers_percentage | Pitcher, percentage of total pitches thrown that were sweepers (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |
| assists_outfield | Fielding, Outfield Assists | MLB | general | CATEGORY_TO_MEASURE |
| at_bats | Batters, number of at-bats | MLB | general | CATEGORY_TO_MEASURE |
| at_bats_bases_empty | At-Bats with no runners on base. | MLB | general | CATEGORY_TO_MEASURE |
| at_bats_bases_loaded | At-bats with the bases loaded. | MLB | general | CATEGORY_TO_MEASURE |
| at_bats_home | Batters, number of at-bats in home games | MLB | general | CATEGORY_TO_MEASURE |
| at_bats_per_home_run | Batters, Ratio of AB to HR (lower number is better, meaning a higher rate of HR hit) | MLB | general | CATEGORY_TO_MEASURE |
| at_bats_per_rbi | Batters, Ratio of AB to RBI (lower number is better, meaning a higher rate of RBI hit) | MLB | general | CATEGORY_TO_MEASURE |
| at_bats_per_strikeout | Batters, Ratio of AB to K (higher number is better, meaning a lower rate of striking out) | MLB | general | CATEGORY_TO_MEASURE |
| at_bats_road | Batters, number of at-bats in road games | MLB | general | CATEGORY_TO_MEASURE |
| at_bats_runners | At-bats with runners on base. | MLB | general | CATEGORY_TO_MEASURE |
| at_bats_runners_in_scoring_position | At-Bats with Runners in Scoring Position | MLB | general | CATEGORY_TO_MEASURE |
| at_bats_runners_in_scoring_position_two_outs | At-Bats with Runners in Scoring Position and two outs. | MLB | general | CATEGORY_TO_MEASURE |
| b_hr_multi_pct | Batters, Pct. of Home Runs hit that had men on base. | MLB | general | CATEGORY_TO_MEASURE |
| b_hr_multirun | Batters, HR hit with men on base | MLB | general | CATEGORY_TO_MEASURE |
| balks | Pitcher Balks | MLB | general | CATEGORY_TO_MEASURE |
| balls_in_field | Batters, number of balls in play (all balls in play minus HR) | MLB | general | CATEGORY_TO_MEASURE |
| balls_in_play | Batters, number of balls in play (all balls in play including HR, SacBunts, SacFlies, etc...) | MLB | general | CATEGORY_TO_MEASURE |
| balls_in_play_per_game | Batters, number of balls in play per game. | MLB | general | CATEGORY_TO_MEASURE |
| batted_balls_events | The number of balls that were put in play. | MLB | general | CATEGORY_TO_MEASURE |
| batter_chase_percentage | (also known as Out-of-Zone Swing Pct.) Batters, swing percentage on pitches out of the zone. (Zones as defined by MLB, 1->9 is in-zone, 11->14 is out-of-zone) | MLB | general | CATEGORY_TO_MEASURE |
| batter_contact_percentage | Batters, percentage of swings where contact was made (swings with contact / swings) | MLB | general | CATEGORY_TO_MEASURE |
| batter_foul_balls | Number of foul balls a batter or batting team hit | MLB | general | CATEGORY_TO_MEASURE |
| batter_home_run_to_fly_ball_percentage | Batters, Pct. of fly balls + line drives that were Home Runs | MLB | general | CATEGORY_TO_MEASURE |
| batter_inzone_contact_percentage | Batters, percentage of swings where contact was made on pitches thrown in the zone. (Zones as defined by MLB, 1->9 is in-zone, 11->14 is out-of-zone) | MLB | general | CATEGORY_TO_MEASURE |
| batter_inzone_swing_percentage | Batters, swing percentage on pitches which were thrown in the zone. (Zones as defined by MLB, 1->9 is in-zone, 11->14 is out-of-zone) | MLB | general | CATEGORY_TO_MEASURE |
| batter_inzone_whiff_percentage | Batters, whiff percentage on pitches thrown in the zone. (Zones as defined by MLB, 1->9 is in-zone, 11->14 is out-of-zone) | MLB | general | CATEGORY_TO_MEASURE |
| batter_outofzone_contact_percentage | Batters, percentage of swings where contact was made on pitches thrown out of the zone. (Zones as defined by MLB, 1->9 is in-zone, 11->14 is out-of-zone) | MLB | general | CATEGORY_TO_MEASURE |
| batter_outofzone_whiff_percentage | Batters, whiff percentage on pitches thrown out of the zone. (Zones as defined by MLB, 1->9 is in-zone, 11->14 is out-of-zone) | MLB | general | CATEGORY_TO_MEASURE |
| batter_percentage_team_runs | Players percentage of team's runs. (Player R / Team R) - NEEDS PLAYER AND TEAM GROUPING INCLUDED IN REPORTS | MLB | general | CATEGORY_TO_MEASURE |
| batter_percentage_team_walks | Batters, percentage of Teams Total Walks - NEEDS PLAYER AND TEAM GROUPING INCLUDED IN REPORTS | MLB | general | CATEGORY_TO_MEASURE |
| batter_percentage_teams_ab | Players percentage of team's at-bats. (Player AB / Team AB) - NEEDS PLAYER AND TEAM GROUPING INCLUDED IN REPORTS | MLB | general | CATEGORY_TO_MEASURE |
| batter_percentage_teams_extra_base_hits | Players percentage of team's extra base hits. (Player XBH / Team XBH) - NEEDS PLAYER AND TEAM GROUPING INCLUDED IN REPORTS | MLB | general | CATEGORY_TO_MEASURE |
| batter_percentage_teams_hits | Players percentage of team's hits. (Player H / Team H) - NEEDS PLAYER AND TEAM GROUPING INCLUDED IN REPORTS | MLB | general | CATEGORY_TO_MEASURE |
| batter_percentage_teams_home_runs | Players percentage of team's home runs. (Player HR / Team HR) - NEEDS PLAYER AND TEAM GROUPING INCLUDED IN REPORTS | MLB | general | CATEGORY_TO_MEASURE |
| batter_percentage_teams_pa | Players percentage of team's plate appearances. (Player PA / Team PA) - NEEDS PLAYER AND TEAM GROUPING INCLUDED IN REPORTS | MLB | general | CATEGORY_TO_MEASURE |
| batter_percentage_teams_rbis | Players percentage of team's RBI. (Player RBI / Team RBI) - NEEDS PLAYER AND TEAM GROUPING INCLUDED IN REPORTS | MLB | general | CATEGORY_TO_MEASURE |
| batter_percentage_teams_strikeouts | Players percentage of team's strikeouts batting. (Player SO / Team SO) - NEEDS PLAYER AND TEAM GROUPING INCLUDED IN REPORTS | MLB | general | CATEGORY_TO_MEASURE |
| batter_swing_percentage | Batters, percentage of pitches that are swung at. | MLB | general | CATEGORY_TO_MEASURE |
| batter_swings | Batters, number of times swinging at the pitch | MLB | general | CATEGORY_TO_MEASURE |
| batter_swings_contact | Batters, number of times swinging at the pitch and making contact | MLB | general | CATEGORY_TO_MEASURE |
| batter_times_chased | Batters, number of times swinging at a pitch out of the zone. (Zones as defined by MLB, 1->9 is in-zone, 11->14 is out-of-zone) | MLB | general | CATEGORY_TO_MEASURE |
| batter_whiff_percentage | Batters, percentage of swings where contact was not made (swing+miss / swings) | MLB | general | CATEGORY_TO_MEASURE |
| batter_whiffs | Batters, number of times swinging at the pitch without making contact | MLB | general | CATEGORY_TO_MEASURE |
| batters_faced | Pitchers number of batters faced | MLB | general | CATEGORY_TO_MEASURE |
| batting_average | Batting Average (hits / at-bats) | MLB | general | CATEGORY_TO_MEASURE |
| batting_average_balls_in_play | Batters, batting average on balls in play. (H - HR) / (AB - K - HR + SF) | MLB | general | CATEGORY_TO_MEASURE |
| batting_average_balls_in_play_number | Batters, number of "balls in play", as a result of SacFlies or AB only, which have a chance to be fielded (AB - HR - K + SacF) | MLB | general | CATEGORY_TO_MEASURE |
| batting_average_bases_empty | Batters, batting average with no runners on base. | MLB | general | CATEGORY_TO_MEASURE |
| batting_average_bases_loaded | Batting average when the bases are loaded. | MLB | general | CATEGORY_TO_MEASURE |
| batting_average_differential | Teams, difference between batting average and opponent batting average | MLB | general | CATEGORY_TO_MEASURE |
| batting_average_home | Batters, batting average in home games | MLB | general | CATEGORY_TO_MEASURE |
| batting_average_home_road_differential | Difference between batting average in home games vs. road games, negative indicates better on the road | MLB | general | CATEGORY_TO_MEASURE |
| batting_average_lhp | Batting average vs. left-handed pitchers | MLB | general | CATEGORY_TO_MEASURE |
| batting_average_lhp_rhp_differential | Difference between batting average against left-handed pitchers vs. right-handed pitchers, negative indicates better vs. right-handed pitchers. | MLB | general | CATEGORY_TO_MEASURE |
| batting_average_rhp | Batting average vs. right-handed pitchers | MLB | general | CATEGORY_TO_MEASURE |
| batting_average_road | Batters, batting average in road games | MLB | general | CATEGORY_TO_MEASURE |
| batting_average_runners_in_scoring_position | Batting Average (hits / at-bats) with Runners in Scoring position. | MLB | general | CATEGORY_TO_MEASURE |
| batting_average_runners_in_scoring_position_two_outs | Batting Average (hits / at-bats) with runners in scoring position and two outs. | MLB | general | CATEGORY_TO_MEASURE |
| batting_average_two_outs | Batting Average (hits / at-bats) with two outs. | MLB | general | CATEGORY_TO_MEASURE |
| batting_averages_runners_on_base | Batting average with runners on base. | MLB | general | CATEGORY_TO_MEASURE |
| batting_games_played | Players, number of games appearing as a batter | MLB | general | CATEGORY_TO_MEASURE |
| batting_summary | Doubles and Triples are spelled out | MLB | general | CATEGORY_TO_MEASURE |
| batting_summary_abb | 2B and 3B for Doubles and Triples | MLB | general | CATEGORY_TO_MEASURE |
| batting_summary_abb_noruns | 2B and 3B for Doubles and Triples (Does not include Runs) | MLB | general | CATEGORY_TO_MEASURE |
| batting_summary_noruns | Doubles and Triples are spelled out (Does not include Runs) | MLB | general | CATEGORY_TO_MEASURE |
| bet_t_w_fav | Games won where the team was favored based on the pre-game moneyline odds | MLB | general | CATEGORY_TO_MEASURE |
| bet_t_w_udog | Wins as the moneyline underdog | MLB | general | CATEGORY_TO_MEASURE |
| called_strike | Total number of called strikes plus whiffs divided by total pitches. | MLB | general | CATEGORY_TO_MEASURE |
| catcher_passed_balls | Catchers, number of passed balls allowed while catching | MLB | general | CATEGORY_TO_MEASURE |
| caught_stealing | Runners, Times caught stealing when attempting to steal a base | MLB | general | CATEGORY_TO_MEASURE |
| combined_extra_base_hits | Combined extra base hits (doubles/triples/home runs) between team and opponent | MLB | general | CATEGORY_TO_MEASURE |
| combined_extra_base_hits_per_game | Combined extra base hits (doubles/triples/home runs) between team and opponent, per game. | MLB | general | CATEGORY_TO_MEASURE |
| combined_hits | Combined Hits between Team and Opponent | MLB | general | CATEGORY_TO_MEASURE |
| combined_hits_per_game | Combined Hits between Team and Opponent, per game. | MLB | general | CATEGORY_TO_MEASURE |
| combined_home_runs | Combined total Home Runs between Team and Opponent. | MLB | general | CATEGORY_TO_MEASURE |
| combined_home_runs_per_game | Combined Home Runs between Team and Opponent, per game. | MLB | general | CATEGORY_TO_MEASURE |
| combined_runs | Combined Runs between Team and Opponent. | MLB | general | CATEGORY_TO_MEASURE |
| combined_runs_per_game | Combined Runs Scored between team and opponent, per game (include Team Grouping/Split) | MLB | general | CATEGORY_TO_MEASURE |
| combined_strikeouts | Combined strikeouts (thrown and from batters) by both teams. | MLB | general | CATEGORY_TO_MEASURE |
| combined_strikeouts_per_game | Combined strikeouts (thrown and from batters) by both teams, per game | MLB | general | CATEGORY_TO_MEASURE |
| doubles | Batters number of hits that were doubles | MLB | general | CATEGORY_TO_MEASURE |
| doubles_per_game | Doubles per game | MLB | general | CATEGORY_TO_MEASURE |
| enforced_balls_hitter | The number of times a hitter was the beneficiary of an enforced ball | MLB | general | CATEGORY_TO_MEASURE |
| enforced_balls_pitcher | The number of times a pitcher was the victim of an enforced ball | MLB | general | CATEGORY_TO_MEASURE |
| enforced_strikes_hitter | The number of times a hitter was the victim of an enforced strike | MLB | general | CATEGORY_TO_MEASURE |
| enforced_strikes_pitcher | The number of times a pitcher was the beneficiary of an enforced strike | MLB | general | CATEGORY_TO_MEASURE |
| errors_fielding | Fielders, number of errors committed classified as "fielding errors" (NOT throwing or interference errors) | MLB | general | CATEGORY_TO_MEASURE |
| errors_interference | Fielders, number of errors committed classified as "interference errors" (NOT fielding or throwing errors) | MLB | general | CATEGORY_TO_MEASURE |
| errors_throwing | Fielders, number of errors committed classified as throwing errors (NOT fielding or interference errors) | MLB | general | CATEGORY_TO_MEASURE |
| errors_total | Fielders, number of errors committed (fielding + throwing + interference) | MLB | general | CATEGORY_TO_MEASURE |
| exit_velocity_average | Batters, average exit velocity on balls in play (among pitches with recorded exit velocity only) | MLB | general | CATEGORY_TO_MEASURE |
| exit_velocity_fly_balls_line_drives_average | Batters, average exit velocity on flyballs and linedrives (among pitches with recorded exit velocity only) | MLB | general | CATEGORY_TO_MEASURE |
| exit_velocity_ground_balls_average | Batters, average exit velocity on ground balls (among pitches with recorded exit velocity only) | MLB | general | CATEGORY_TO_MEASURE |
| extra_base_hits | Batters, number of hits for extra bases (2B + 3B + HR) | MLB | general | CATEGORY_TO_MEASURE |
| extra_base_hits_differential | Teams, difference between batting XBH and pitching XBH allowed | MLB | general | CATEGORY_TO_MEASURE |
| extra_base_hits_per_game | The number of extra-base hits per game. | MLB | general | CATEGORY_TO_MEASURE |
| fantasy_points_draftkings_batting | Draftkings MLB Batting Fantasy Points | MLB | general | CATEGORY_TO_MEASURE |
| fantasy_points_draftkings_pitching | Draftkings MLB Pitching Fantasy Points | MLB | general | CATEGORY_TO_MEASURE |
| fantasy_points_per_game_draftkings_batting | Draftkings MLB Batting Fantasy Points per Game | MLB | general | CATEGORY_TO_MEASURE |
| fantasy_points_per_game_draftkings_pitching | Draftkings MLB Pitching Fantasy Points per Game | MLB | general | CATEGORY_TO_MEASURE |
| fielders_choice | Batters, number of times hitting into fielder's choice | MLB | general | CATEGORY_TO_MEASURE |
| fielding_assists | Fielders, number of assists | MLB | general | CATEGORY_TO_MEASURE |
| fielding_caught_stealing | Number of opposing runners caught stealing | MLB | general | CATEGORY_TO_MEASURE |
| fielding_caught_stealing_against_percentage | Percentage of opposing runners caught stealing | MLB | general | CATEGORY_TO_MEASURE |
| fielding_double_plays | Fielders, number of double-plays involved in | MLB | general | CATEGORY_TO_MEASURE |
| fielding_games_played | Players, number of games played as a fielder | MLB | general | CATEGORY_TO_MEASURE |
| fielding_innings_played | Innings played in the field (calculated from outs while fielding) | MLB | general | CATEGORY_TO_MEASURE |
| fielding_percentage | Fielding Percentage. Putouts+Assists/Total Chances (PO+Assists+Errors) | MLB | general | CATEGORY_TO_MEASURE |
| fielding_putouts | Fielders, number of putouts | MLB | general | CATEGORY_TO_MEASURE |
| fielding_steals_against | Number of stolen bases against (opponent stolen bases) | MLB | general | CATEGORY_TO_MEASURE |
| fielding_steals_attempts_against | Number of stolen base attempts against (opponent's Steals + Caught Stealing) | MLB | general | CATEGORY_TO_MEASURE |
| fielding_stolen_base_percentage_against | Percentage of successful stolen base attempts against | MLB | general | CATEGORY_TO_MEASURE |
| fielding_total_chances | Players, total fielding chances (Putouts + Assists + Errors) | MLB | general | CATEGORY_TO_MEASURE |
| fielding_triple_plays | Fielders, number of triple-plays involved in | MLB | general | CATEGORY_TO_MEASURE |
| fielding_wild_pitches | Catchers, number of wild pitches thrown while at catcher | MLB | general | CATEGORY_TO_MEASURE |
| first_pitch_balls_seen | Batters, number of plate appearances start with a ball called | MLB | general | CATEGORY_TO_MEASURE |
| first_pitch_balls_thrown | Pitchers, first pitch of plate appearance called a ball | MLB | general | CATEGORY_TO_MEASURE |
| first_pitch_hits | Batters, hits on the first pitch of an at-bat | MLB | general | CATEGORY_TO_MEASURE |
| first_pitch_hits_allowed | Pitchers, hits allowed on the first pitch of an at-bat | MLB | general | CATEGORY_TO_MEASURE |
| first_pitch_home_run_allowed | Pitchers, home runs allowed on the first pitch of an at-bat | MLB | general | CATEGORY_TO_MEASURE |
| first_pitch_home_run_hit | Batters, home run on first pitch of at-bat | MLB | general | CATEGORY_TO_MEASURE |
| first_pitch_strikes_seen | Batters, number of strikes on the first pitch of a plate appearance | MLB | general | CATEGORY_TO_MEASURE |
| first_pitch_strikes_seen_percentage | Batters, percentage of plate appearances to start with a first pitch strike | MLB | general | CATEGORY_TO_MEASURE |
| first_pitch_strikes_thrown | Pitchers, strikes thrown on the first pitch of a plate appearance | MLB | general | CATEGORY_TO_MEASURE |
| first_pitch_strikes_thrown_percentage | Pitchers, percent of plate appearances to start with a strike | MLB | general | CATEGORY_TO_MEASURE |
| first_pitch_swing_percentage | Batters, percentage of first pitches swung at | MLB | general | CATEGORY_TO_MEASURE |
| first_pitch_swing_percentage_against | Pitchers, Percentage of swings against on the first pitch of a plate appearance | MLB | general | CATEGORY_TO_MEASURE |
| first_pitch_swings | Batters, swings on the first pitch of an at-bat | MLB | general | CATEGORY_TO_MEASURE |
| first_pitch_swings_agaist | Pitchers, swings against on the first pitch of an at-bat | MLB | general | CATEGORY_TO_MEASURE |
| fly_ball_outs | Batters, number of times out on a fly ball. (Note: does not include line drives or pop-ups) | MLB | general | CATEGORY_TO_MEASURE |
| fly_balls | Batters, number of balls in play that were fly balls. (Note: does not include line drives or pop-ups) | MLB | general | CATEGORY_TO_MEASURE |
| fly_balls_percentage | Batters, percentage of balls in play that were Fly Balls (Note: does not include line drives or pop-ups) | MLB | general | CATEGORY_TO_MEASURE |
| fly_pop_balls | Batters, number balls in play that were either Fly Balls or Pop-Ups. (Note: does not include Line Drives) | MLB | general | CATEGORY_TO_MEASURE |
| fly_pop_balls_percentage | Batters, percentage of balls in play that were either Fly Balls or Pop-Ups. (Note: does not include Line Drives) | MLB | general | CATEGORY_TO_MEASURE |
| game_score | Metric devised by Bill James to determine the strength of a pitcher in any particular baseball game | MLB | general | CATEGORY_TO_MEASURE |
| game_winning_rbi | Batters, RBI that gives a team the lead it never relinquishes | MLB | general | CATEGORY_TO_MEASURE |
| gamelog_batting_average | Windowed batting average based on the level of stats that are being viewed. For example, when viewing by games, this is the batting average at that point in the season. | MLB | general | CATEGORY_TO_MEASURE |
| gamelog_era | Windowed Earned Run Average based on the level of stats that are being viewed. For example, when viewing by games, this is the earned run average at that point in the season. | MLB | general | CATEGORY_TO_MEASURE |
| gamelog_era_reliever | Windowed Earned Run Average for relievers based on the level of stats that are being viewed. For example, when viewing by games, this is the earned run average at that point in the season. | MLB | general | CATEGORY_TO_MEASURE |
| gamelog_era_starter | Windowed Earned Run Average for starters based on the level of stats that are being viewed. For example, when viewing by games, this is the earned run average at that point in the season. | MLB | general | CATEGORY_TO_MEASURE |
| gamelog_era_team | Windowed Team Earned Run Average based on the level of stats that are being viewed. For example, when viewing by games, this is the earned run average at that point in the season. | MLB | general | CATEGORY_TO_MEASURE |
| gamelog_on_base_percentage | Windowed On-Base Percentage based on the level of stats that are being viewed. For example, when viewing by games, this is the on-base percentage at that point in the season. | MLB | general | CATEGORY_TO_MEASURE |
| gamelog_on_base_plus_slugging | Windowed On-Base Plus Slugging Percentage based on the level of stats that are being viewed. For example, when viewing by games, this is the on-base plus slugging percentage at that point in the season. | MLB | general | CATEGORY_TO_MEASURE |
| gamelog_slugging | Windowed slugging percentage based on the level of stats that are being viewed. For example, when viewing by games, this is the slugging percentage at that point in the season. | MLB | general | CATEGORY_TO_MEASURE |
| gamelog_team_losses | Running Total of Team Losses based on the level of stats that are being viewed. For example, when viewing by games, this is the team losses at that point in the season. | MLB | general | CATEGORY_TO_MEASURE |
| gamelog_team_winning_percentage | Running Total of Team Winning Percentage based on the level of stats that are being viewed. For example, when viewing by games, this is the team win pct. at that point in the season. | MLB | general | CATEGORY_TO_MEASURE |
| gamelog_team_wins | Running Total of Team Wins based on the level of stats that are being viewed. For example, when viewing by games, this is the team wins at that point in the season. | MLB | general | CATEGORY_TO_MEASURE |
| gamelog_whip | Windowed Pitcher WHIP based on the level of stats that are being viewed. For example, when viewing by games, this is the pitcher WHIP at that point in the season. | MLB | general | CATEGORY_TO_MEASURE |
| games_day | Number of day games played | MLB | general | CATEGORY_TO_MEASURE |
| games_favorite | Team Games Played as the Moneyline Favorite | MLB | general | CATEGORY_TO_MEASURE |
| games_finishes | Relief Pitchers, number of games finished (note: relief pitcher statistic only. Starters are not awarded games finished for complete games) | MLB | general | CATEGORY_TO_MEASURE |
| games_night | Number of night games played | MLB | general | CATEGORY_TO_MEASURE |
| games_pitched | Number of games appearing as pitcher | MLB | general | CATEGORY_TO_MEASURE |
| games_pitched_home | Number of games appearing as pitcher in home games | MLB | general | CATEGORY_TO_MEASURE |
| games_pitched_road | Number of games appearing as pitcher on the road | MLB | general | CATEGORY_TO_MEASURE |
| games_pitched_total | Total Games Pitched - When grouping by team, this will count every individual player game pitched as one instead of showing only one per game. | MLB | general | CATEGORY_TO_MEASURE |
| games_played | Games played | MLB | general | CATEGORY_TO_MEASURE |
| games_played_total | Total Games Played - When grouping by team, this will count every individual player game played as one instead of showing the team's games played. | MLB | general | CATEGORY_TO_MEASURE |
| games_started | Number of games starting at pitcher | MLB | general | CATEGORY_TO_MEASURE |
| go_ahead_rbi | Batters, RBI that gave their team the lead | MLB | general | CATEGORY_TO_MEASURE |
| ground_balls | Batters, number of balls in play that were ground balls. | MLB | general | CATEGORY_TO_MEASURE |
| ground_balls_percentage | Batters, percentage of balls in play that were Ground Balls | MLB | general | CATEGORY_TO_MEASURE |
| ground_balls_to_air_balls_differential | Batters, Ratio of balls in play, Ground Balls to Air Balls. (Note: Air Balls = balls in play classified as Fly Ball, Line Drive, or Pop-Up) | MLB | general | CATEGORY_TO_MEASURE |
| ground_balls_to_fly_pop_balls_differential | Batters, Ratio of balls in play, Ground Balls to Fly (F+P). (Note: Fly = balls in play classified as either Fly Ball or Pop-Up) | MLB | general | CATEGORY_TO_MEASURE |
| ground_outs | Batters, number of times grounding out. | MLB | general | CATEGORY_TO_MEASURE |
| ground_outs_percentage | Percentage of ground balls that are groundouts. | MLB | general | CATEGORY_TO_MEASURE |
| grounded_into_double_play | Batters, times grounding into a double-play | MLB | general | CATEGORY_TO_MEASURE |
| gs | Players, number of games started (either fielding, batting, or pitching) | MLB | general | CATEGORY_TO_MEASURE |
| hard_hit_balls | Batters, number of batted balls in play with an exit velocity (launch speed) of 95 MPH or higher | MLB | general | CATEGORY_TO_MEASURE |
| hard_hit_percentage | Batters, percentage of batted balls in play with an exit velocity (launch speed) of 95 MPH or higher | MLB | general | CATEGORY_TO_MEASURE |
| hit_by_pitches | Batters, number of times hit by pitch | MLB | general | CATEGORY_TO_MEASURE |
| hits | Batters, number of hits | MLB | general | CATEGORY_TO_MEASURE |
| hits_allowed_extra_bases_percentage | Pitchers, percentage of hits allowed that are extra-base hits - (2B, 3B, HR)/Hits | MLB | general | CATEGORY_TO_MEASURE |
| hits_bases_empty | Number of hits with no runners on base. | MLB | general | CATEGORY_TO_MEASURE |
| hits_bases_loaded | Number of hits when the bases are loaded. | MLB | general | CATEGORY_TO_MEASURE |
| hits_bunts | Batters, bunt hits | MLB | general | CATEGORY_TO_MEASURE |
| hits_by_five_directions_center | Batters, hits to center field (when dividing the field into 5 equal fields, LF - LCF - CF - RCF - RF) | MLB | general | CATEGORY_TO_MEASURE |
| hits_by_five_directions_center_percentage | Batters, percentage of hits that are hits to center field (when dividing the field into 5 equal fields, LF - LCF - CF - RCF - RF) | MLB | general | CATEGORY_TO_MEASURE |
| hits_by_five_directions_left | Batters, hits to left field (when dividing the field into 5 equal fields, LF - LCF - CF - RCF - RF) | MLB | general | CATEGORY_TO_MEASURE |
| hits_by_five_directions_left_center | Batters, hits to left-center field (when dividing the field into 5 equal fields, LF - LCF - CF - RCF - RF) | MLB | general | CATEGORY_TO_MEASURE |
| hits_by_five_directions_left_center_percentage | Batters, percentage of hits that are hits to left-center field (when dividing the field into 5 equal fields, LF - LCF - CF - RCF - RF) | MLB | general | CATEGORY_TO_MEASURE |
| hits_by_five_directions_left_percentage | Batters, percentage of hits that are hits to left field (when dividing the field into 5 equal fields, LF - LCF - CF - RCF - RF) | MLB | general | CATEGORY_TO_MEASURE |
| hits_by_five_directions_right | Batters, hits to right field (when dividing the field into 5 equal fields, LF - LCF - CF - RCF - RF) | MLB | general | CATEGORY_TO_MEASURE |
| hits_by_five_directions_right_center | Batters, hits to right-center field (when dividing the field into 5 equal fields, LF - LCF - CF - RCF - RF) | MLB | general | CATEGORY_TO_MEASURE |
| hits_by_five_directions_right_center_percentage | Batters, percentage of hits that are hits to right-center field (when dividing the field into 5 equal fields, LF - LCF - CF - RCF - RF) | MLB | general | CATEGORY_TO_MEASURE |
| hits_by_five_directions_right_percentage | Batters, percentage of hits that are hits to right field (when dividing the field into 5 equal fields, LF - LCF - CF - RCF - RF) | MLB | general | CATEGORY_TO_MEASURE |
| hits_by_three_directions_center | Batters, hits to center (when dividing the field into 3 equal fields) | MLB | general | CATEGORY_TO_MEASURE |
| hits_by_three_directions_center_percentage | Batters, percentage of hits that are hits to center (when dividing the field into 3 equal fields) | MLB | general | CATEGORY_TO_MEASURE |
| hits_by_three_directions_left | Batters, hits to the left side (when dividing the field into 3 equal fields) | MLB | general | CATEGORY_TO_MEASURE |
| hits_by_three_directions_left_percentage | Batters, percentage of hits that are hits to the left side (when dividing the field into 3 equal fields) | MLB | general | CATEGORY_TO_MEASURE |
| hits_by_three_directions_right | Batters, hits to the right side (when dividing the field into 3 equal fields) | MLB | general | CATEGORY_TO_MEASURE |
| hits_by_three_directions_right_percentage | Batters, percentage of hits that are hits to the right side (when dividing the field into 3 equal fields) | MLB | general | CATEGORY_TO_MEASURE |
| hits_differential | Teams, difference between hits by offense and hits allowed by pitchers | MLB | general | CATEGORY_TO_MEASURE |
| hits_doubles_percentage | Batters, percentage of hits that are doubles | MLB | general | CATEGORY_TO_MEASURE |
| hits_extra_bases_percentage | Batters, percentage of hits that are extra-base hits - (2B, 3B, HR)/Hits | MLB | general | CATEGORY_TO_MEASURE |
| hits_home | Batters, number of hits in home games | MLB | general | CATEGORY_TO_MEASURE |
| hits_home_road_differential | Difference between hits in home games vs. road games, negative indicates more on the road | MLB | general | CATEGORY_TO_MEASURE |
| hits_home_runs_percentage | Batters, percentage of hits that are home runs | MLB | general | CATEGORY_TO_MEASURE |
| hits_infield | Battters, infield hits | MLB | general | CATEGORY_TO_MEASURE |
| hits_opposite | Amount of Hits that were hit to Opposite Field as Batter | MLB | general | CATEGORY_TO_MEASURE |
| hits_opposite_percentage | Batters, percentage of hits that are hit to the opposite field (i.e. RH batter hits to Right field, LH batter hits to Left field) | MLB | general | CATEGORY_TO_MEASURE |
| hits_per_game | Total Hits per Games Played | MLB | general | CATEGORY_TO_MEASURE |
| hits_per_game_differential | Teams, difference between hits by offense and hits allowed by pitchers per game | MLB | general | CATEGORY_TO_MEASURE |
| hits_pulled | Amount of Hits that were pulled in relation to Batter | MLB | general | CATEGORY_TO_MEASURE |
| hits_pulled_percentage | Batters, percentage of hits that are pulled (i.e. RH batter hits to Left field, LH batter hits to Right field) | MLB | general | CATEGORY_TO_MEASURE |
| hits_road | Batters, number of hits in road games | MLB | general | CATEGORY_TO_MEASURE |
| hits_runners_in_scoring_position | Hits with Runners in Scoring Position | MLB | general | CATEGORY_TO_MEASURE |
| hits_runners_in_scoring_position_two_outs | Hits with Runners in Scoring Position and two outs. | MLB | general | CATEGORY_TO_MEASURE |
| hits_runners_on_base | Hits with runners on base. | MLB | general | CATEGORY_TO_MEASURE |
| hits_singles_percentage | Batters, percentage of hits that are singles | MLB | general | CATEGORY_TO_MEASURE |
| hits_triples_percentage | Batters, percentage of hits that are triples | MLB | general | CATEGORY_TO_MEASURE |
| hitter_left_on_base | The amount of runners left on base after the hitter gets out from the Hitters perspective | MLB | general | CATEGORY_TO_MEASURE |
| home_run_distance | Batters, batted ball distance on home runs (NOTE: To see the distance for individual home runs, add the grouping "Plate Appearance Description") | MLB | general | CATEGORY_TO_MEASURE |
| home_run_distance_average | Batters, average batted ball distance of home runs (among HR with recorded distances only) | MLB | general | CATEGORY_TO_MEASURE |
| home_run_launch_angle | Batters, Launch Angle on balls for homeruns (NOTE: To see the Launch Angle for individual home runs, add the grouping "Plate Appearance Description") | MLB | general | CATEGORY_TO_MEASURE |
| home_runs | Batters, Home Runs hit | MLB | general | CATEGORY_TO_MEASURE |
| home_runs_differential | Home Run Differential. Home Runs-Home Runs allowed | MLB | general | CATEGORY_TO_MEASURE |
| home_runs_differential_home_road | Difference between home runs hit in home games vs. road games, negative indicates more on the road | MLB | general | CATEGORY_TO_MEASURE |
| home_runs_game_tying | Batters, home runs hit while their team was trailing to tie the game. | MLB | general | CATEGORY_TO_MEASURE |
| home_runs_go_ahead | Batters, home runs hit while team was tied or trailing which gave the team the lead. | MLB | general | CATEGORY_TO_MEASURE |
| home_runs_grand_slam | Grand Slam Home Runs | MLB | general | CATEGORY_TO_MEASURE |
| home_runs_home | Batters, Home Runs hit in home games | MLB | general | CATEGORY_TO_MEASURE |
| home_runs_per_game | Home Runs hit per Game | MLB | general | CATEGORY_TO_MEASURE |
| home_runs_road | Batters, Home Runs hit in road games | MLB | general | CATEGORY_TO_MEASURE |
| home_runs_solo | Batters, Solo Home Runs hit | MLB | general | CATEGORY_TO_MEASURE |
| home_runs_solo_pct | Batters, Pct. of Home Runs hit that were Solo shots. | MLB | general | CATEGORY_TO_MEASURE |
| home_runs_three_run | Batters, HR hit with 2 men on base. | MLB | general | CATEGORY_TO_MEASURE |
| home_runs_two_run | Batters, HR with one man on base. | MLB | general | CATEGORY_TO_MEASURE |
| home_wins | Wins in home games | MLB | general | CATEGORY_TO_MEASURE |
| homerun_leadoff | Batters, HR hit in the first Plate Appearance of the game for their team (can be either the top or bottom of the 1st-inning) | MLB | general | CATEGORY_TO_MEASURE |
| homerun_leadoff_inning | Batters, HR hit in the first Plate Appearance of any half-inning | MLB | general | CATEGORY_TO_MEASURE |
| homerun_leadoff_top | Batters, HR hit in the first Plate Appearance of the game (top of the 1st inning) | MLB | general | CATEGORY_TO_MEASURE |
| homeruns_by_five_directions_center | Batters, percentage of home runs that are hit to center field (when dividing the field into 5 equal fields, LF - LCF - CF - RCF - RF) | MLB | general | CATEGORY_TO_MEASURE |
| homeruns_by_five_directions_center_percentage | Batters, home runs hit to center field (when dividing the field into 5 equal fields, LF - LCF - CF - RCF - RF) | MLB | general | CATEGORY_TO_MEASURE |
| homeruns_by_five_directions_left_center | Batters, percentage of home runs that are hit to left-center field (when dividing the field into 5 equal fields, LF - LCF - CF - RCF - RF) | MLB | general | CATEGORY_TO_MEASURE |
| homeruns_by_five_directions_left_center_percentage | Batters, home runs hit to left-center field (when dividing the field into 5 equal fields, LF - LCF - CF - RCF - RF) | MLB | general | CATEGORY_TO_MEASURE |
| homeruns_by_five_directions_left_percentage | Batters, home runs hit to left field (when dividing the field into 5 equal fields, LF - LCF - CF - RCF - RF) | MLB | general | CATEGORY_TO_MEASURE |
| homeruns_by_five_directions_right_center_percentage | Batters, home runs hit to right-center field (when dividing the field into 5 equal fields, LF - LCF - CF - RCF - RF) | MLB | general | CATEGORY_TO_MEASURE |
| homeruns_by_five_directions_right_percentage | Batters, home runs hit to right field (when dividing the field into 5 equal fields, LF - LCF - CF - RCF - RF) | MLB | general | CATEGORY_TO_MEASURE |
| homeruns_by_three_directions_center | Batters, home runs hit to center (when dividing the field into 3 equal fields) | MLB | general | CATEGORY_TO_MEASURE |
| homeruns_by_three_directions_center_percentage | Batters, percentage of home runs that are hit to center (when dividing the field into 3 equal fields) | MLB | general | CATEGORY_TO_MEASURE |
| homeruns_by_three_directions_left | Batters, home runs hit to the left side (when dividing the field into 3 equal fields) | MLB | general | CATEGORY_TO_MEASURE |
| homeruns_by_three_directions_left_percentage | Batters, percentage of home runs that are hit to the left side (when dividing the field into 3 equal fields) | MLB | general | CATEGORY_TO_MEASURE |
| homeruns_by_three_directions_right | Batters, home runs hit to the right side (when dividing the field into 3 equal fields) | MLB | general | CATEGORY_TO_MEASURE |
| homeruns_by_three_directions_right_percentage | Batters, percentage of home runs that are hit to the right side (when dividing the field into 3 equal fields) | MLB | general | CATEGORY_TO_MEASURE |
| homeruns_opposite | Amount of Home Runs that were hit to Opposite Field as Batter | MLB | general | CATEGORY_TO_MEASURE |
| homeruns_opposite_percentage | Batters, percentage of home runs that are hit to the opposite field (i.e. RH batter HRs to Right field, LH batter HRs to Left field) | MLB | general | CATEGORY_TO_MEASURE |
| homeruns_pulled | Amount of Home Runs that were pulled in relation to Batter | MLB | general | CATEGORY_TO_MEASURE |
| homeruns_pulled_percentage | Batters, percentage of home runs that are pulled (i.e. RH batter HRs to Left field, LH batter HRs to Rightt field) | MLB | general | CATEGORY_TO_MEASURE |
| homes_losses | Losses in home games | MLB | general | CATEGORY_TO_MEASURE |
| hr__by__five__directions__left_pct | Batters, percentage of home runs that are hit to left field (when dividing the field into 5 equal fields, LF - LCF - CF - RCF - RF) | MLB | general | CATEGORY_TO_MEASURE |
| hr__by__five__directions__right__center_pct | Batters, percentage of home runs that are hit to right-center field (when dividing the field into 5 equal fields, LF - LCF - CF - RCF - RF) | MLB | general | CATEGORY_TO_MEASURE |
| hr__by__five__directions__right_pct | Batters, percentage of home runs that are hit to right field (when dividing the field into 5 equal fields, LF - LCF - CF - RCF - RF) | MLB | general | CATEGORY_TO_MEASURE |
| inherited_runners | Relief Pitchers, the number of runners on base when the pitcher enters the game | MLB | general | CATEGORY_TO_MEASURE |
| inherited_runners_scored | Relief Pitchers, the number of inherited runners allowed to score | MLB | general | CATEGORY_TO_MEASURE |
| inherited_runners_scored_percentage | Relief Pitchers, the percentage of inherited runners allowed to score | MLB | general | CATEGORY_TO_MEASURE |
| intentional_walks | Batters, times intentionally walked by pitchers | MLB | general | CATEGORY_TO_MEASURE |
| isolated_power | (Extra bases / At-Bats) or (SLG - AVG) | MLB | general | CATEGORY_TO_MEASURE |
| isolated_power_discipline | Batters, On-base percentage minus batting average. | MLB | general | CATEGORY_TO_MEASURE |
| launch_angle_average | Batters, average launch angle on balls in play (among pitches with recorded launch angle only) | MLB | general | CATEGORY_TO_MEASURE |
| line_drive_outs | Batters, number of times lining out. | MLB | general | CATEGORY_TO_MEASURE |
| line_drive_percentage | Batters, percentage of balls in play that were Line Drives | MLB | general | CATEGORY_TO_MEASURE |
| line_drives | Batters, number of balls in play that were line drives. | MLB | general | CATEGORY_TO_MEASURE |
| losses_day | Losses in day games | MLB | general | CATEGORY_TO_MEASURE |
| losses_lhp | Losses in games where the opposing starting pitcher was left-handed | MLB | general | CATEGORY_TO_MEASURE |
| losses_night | Losses in night games | MLB | general | CATEGORY_TO_MEASURE |
| losses_post_all_star | Losses after the All-Star break. (NOTE: first All-Star game was in 1933) | MLB | general | CATEGORY_TO_MEASURE |
| losses_pre_all_star | Losses before the All-Star break. (NOTE: first All-Star game was in 1933) | MLB | general | CATEGORY_TO_MEASURE |
| losses_rhp | Losses in games where the opposing starting pitcher was right-handed | MLB | general | CATEGORY_TO_MEASURE |
| made_playoff_season | Seasons in which the team made the playoffs. | MLB | general | CATEGORY_TO_MEASURE |
| managers | Number of managers | MLB | general | CATEGORY_TO_MEASURE |
| moneyline_favorite_losses | Losses as the pre-game moneyline favorite | MLB | general | CATEGORY_TO_MEASURE |
| moneyline_favorite_wins | The number of times a team covered the runline as a favorite. | MLB | general | CATEGORY_TO_MEASURE |
| moneyline_favorite_wins_percentage | The percentage of games as a favorite where a team covered the runline. | MLB | general | CATEGORY_TO_MEASURE |
| moneyline_underdog_games | Games where the team was the underdog based on the pre-game moneyline odds | MLB | general | CATEGORY_TO_MEASURE |
| moneyline_underdog_losses | Losses as the moneyline underdog | MLB | general | CATEGORY_TO_MEASURE |
| moneyline_underdog_wins | The number of times a team beat the spread as an underdog. | MLB | general | CATEGORY_TO_MEASURE |
| moneyline_underdog_wins_percentage | The percentage of games as an underdog where a team beat the runline. | MLB | general | CATEGORY_TO_MEASURE |
| on_base_percentage | (H + BB+ HBP) / (AB + BB + HBP + SacF) | MLB | general | CATEGORY_TO_MEASURE |
| on_base_percentage_differential | Teams, difference between batting OBP and pitching OBP allowed | MLB | general | CATEGORY_TO_MEASURE |
| on_base_percentage_home | On-Base Pct. in home games | MLB | general | CATEGORY_TO_MEASURE |
| on_base_percentage_home_road_differential | Difference between OBP in home games vs. road games, negative indicates better on the road | MLB | general | CATEGORY_TO_MEASURE |
| on_base_percentage_lhp | OBP vs. left-handed pitchers | MLB | general | CATEGORY_TO_MEASURE |
| on_base_percentage_rhp | OBP vs. right-handed pitchers | MLB | general | CATEGORY_TO_MEASURE |
| on_base_percentage_road | On-Base Pct. in road games | MLB | general | CATEGORY_TO_MEASURE |
| on_base_plus_slugging | On Base Percentage + Slugging Percentage | MLB | general | CATEGORY_TO_MEASURE |
| on_base_plus_slugging_differential | Difference between batting OPS and pitching OPS allowed. A negative value indicates a higher OPS allowed. | MLB | general | CATEGORY_TO_MEASURE |
| on_base_plus_slugging_home | OPS in home games | MLB | general | CATEGORY_TO_MEASURE |
| on_base_plus_slugging_home_road_differential | Difference between OPS in home games vs. road games, negative indicates better on the road | MLB | general | CATEGORY_TO_MEASURE |
| on_base_plus_slugging_lhp | Batters, OPS vs. left-handed pitchers | MLB | general | CATEGORY_TO_MEASURE |
| on_base_plus_slugging_lhp_rhp_differential | Difference between OPS against left-handed pitchers vs. right-handed pitchers, negative indicates better vs. right-handed pitchers. | MLB | general | CATEGORY_TO_MEASURE |
| on_base_plus_slugging_rhp | Batters, OPS vs. right-handed pitchers | MLB | general | CATEGORY_TO_MEASURE |
| on_base_plus_slugging_road | OPS in road games | MLB | general | CATEGORY_TO_MEASURE |
| p_pko | Pitchers, number of times picking off a baserunner | MLB | general | CATEGORY_TO_MEASURE |
| pickoffs | Players, times picked off as a baserunner | MLB | general | CATEGORY_TO_MEASURE |
| pitch_category_breaking_balls | Pitcher, number of breaking balls thrown (all breaking ball types) | MLB | general | CATEGORY_TO_MEASURE |
| pitch_category_breaking_balls_percentage | Pitcher, percentage of total pitches thrown that were breaking balls (all breaking ball types) (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |
| pitch_category_changeups | Pitcher, number of changeups thrown (all changeup types) -- grouped to match Baseball Savant | MLB | general | CATEGORY_TO_MEASURE |
| pitch_category_changeups_percentage | Pitcher, percentage of total pitches thrown that were changeups (among pitches with recorded pitch type only, grouped to match Savant) | MLB | general | CATEGORY_TO_MEASURE |
| pitch_category_curveballs | Pitcher, number of curveballs thrown (all curveball types) -- grouped to match Baseball Savant | MLB | general | CATEGORY_TO_MEASURE |
| pitch_category_curveballs_percentage | Pitcher, percentage of total pitches thrown that were curveballs (among pitches with recorded pitch type only, grouped to match Savant) | MLB | general | CATEGORY_TO_MEASURE |
| pitch_category_cutters | Pitcher, number of cutters thrown -- grouped to match Baseball Savant | MLB | general | CATEGORY_TO_MEASURE |
| pitch_category_cutters_percentage | Pitcher, percentage of total pitches thrown that were cutters (among pitches with recorded pitch type only, grouped to match Savant) | MLB | general | CATEGORY_TO_MEASURE |
| pitch_category_fastballs | Pitcher, number of fastballs thrown (all fastball types) | MLB | general | CATEGORY_TO_MEASURE |
| pitch_category_fastballs_percentage | Pitcher, percentage of total pitches thrown that were fastballs (all fastball types) (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |
| pitch_category_fastballs_percentage_savant | Pitcher, percentage of total pitches thrown that were fastballs (among pitches with recorded pitch type only, grouped to match Savant) | MLB | general | CATEGORY_TO_MEASURE |
| pitch_category_fastballs_savant | Pitcher, number of fastballs thrown (all fastball types) -- grouped to match Baseball Savant | MLB | general | CATEGORY_TO_MEASURE |
| pitch_category_offspeeds | Pitcher, number of offspeed pitches thrown (all offspeed types) | MLB | general | CATEGORY_TO_MEASURE |
| pitch_category_offspeeds_percentage | Pitcher, percentage of total pitches thrown that were offspeed (all offspeed types) (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |
| pitch_category_sliders | Pitcher, number of sliders thrown (all slider types) -- grouped to match Baseball Savant | MLB | general | CATEGORY_TO_MEASURE |
| pitch_category_sliders_percentage | Pitcher, percentage of total pitches thrown that were sliders (among pitches with recorded pitch type only, grouped to match Savant) | MLB | general | CATEGORY_TO_MEASURE |
| pitch_velocity_faced_average | Batters, average pitch speed faced (among pitches with recorded velocity only) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_air_balls_against | Pitchers, number of balls in play by opposing batters that were hit in the air. (Fly Balls + Line Drives + Pop-Ups) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_air_balls_against_percentage | Pitchers, percentage of balls in play by opposing batters that were hit in the air. (Fly Balls + Line Drives + Pop-Ups) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_at_bats_against | Pitchers, number of batters faced which resulted in an at-bat | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_blown_saves | Pitches, blown saves | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_breaking_balls_velocity_average | Pitchers, average breaking ball pitch type pitch speed (among pitches with recorded velocity only) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_caught_stealing | Pitchers, number of runners caught stealing | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_changeups_velocity_average | Pitchers, average changeup pitch speed (among pitches with recorded velocity only) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_chase_percentage | Pitchers, opposing batters' swing percentage on pitches out of the zone. (Zones as defined by MLB, 1->9 is in-zone, 11->14 is out-of-zone) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_complete_games | Pitcher complete games (started and finished the game) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_contact_against_percentage | Pitchers, percentage of swings by opposing batters where contact was made (swings with contact / swings) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_curveballs_velocity_average | Pitchers, average curveball pitch speed (among pitches with recorded velocity only) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_cutters_velocity_average | Pitchers, average cut fastball pitch speed (among pitches with recorded velocity only) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_decisions | Pitchers, times individually receiving a winning or losing decision | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_doubles_against | Pitchers, number of doubles hit against them | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_earned_runs_allowed | Pitchers, Earned Runs charged to the individual pitcher (NOTE: may differ from Team ER due to Section 10.18 (i) of the Scoring rules) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_earned_runs_relieiver | Pitchers, Earned Runs charged to the individual pitcher in games as a relief pitcher | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_earned_runs_starter | Pitchers, Earned Runs charged to the individual pitcher in games as the starting pitcher | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_eephuses_velocity_average | Pitchers, average eephus pitch speed (among pitches with recorded velocity only) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_era | PItchers, Earned Runs allowed per 9 IP | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_era_home | Pitchers, ERA in home games | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_era_home_road_differential | Pitchers, Difference between ERA at home vs. away | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_era_reliver | Pitchers, Earned Runs allowed per 9 IP in games as a relief pitcher | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_era_road | Pitchers, ERA in road games | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_era_starter | Pitchers, Earned Runs allowed per 9 IP in games as the starting pitcher | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_era_starter_reliever_differential | ERA Differential between Starting and Relieving | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_exit_velocity_against_average | Pitchers, average exit velocity by opposing batters on balls in play (among pitches with recorded exit velocity only) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_extra_base_hits | Pitchers, number of extra base hits allowed to opposing batters (2B + 3B + HR) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_fastballs_velocity_average | Pitchers, average fastball pitch type pitch speed (among pitches with recorded velocity only) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_fielders_choice_against | Pitchers, number of times opposing batters hit into fielder's choice | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_forkballs_velocity_average | Pitchers, average forkball pitch speed (among pitches with recorded velocity only) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_foul_balls | Number of foul balls against or pitcher or the pitching team | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_four_seam_fastballs_velocity_average | Pitchers, average 4-seam fastball pitch speed (among pitches with recorded velocity only) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_grand_slams_allowed | Pitchers, grand slam home runs allowed | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_groundball_double_plays | Pitchers ground-ball double-plays induced. | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_groundballto_fly_out_ratio | Pitchers, Ratio of balls in play, Ground Balls to Fly (F+P). (Note: Fly = balls in play classified as either Fly Ball or Pop-Up) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_h_r_to_fly_percentage | Pitchers, the percentage of fly balls + line drives by opposing batters that were home runs | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_hard_hit_balls_against | Pitchers, number of batted balls in play by opposing batters with an exit velocity (launch speed) of 95 MPH or higher | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_hard_hit_balls_against_percentage | Pitchers, percentage of batted balls in play by opposing batters with an exit velocity (launch speed) of 95 MPH or higher | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_hit_by_pitches | Pitchers, number of batters hit by pitch | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_hit_distance_against_average | Pitchers, average batted ball distance allowed to opposing batters (among pitches with recorded distances only) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_hits_allowed | Pitchers, hits allowed | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_hits_allowed_for_home_runs_percentage | Pitchers, percentage of hits allowed to opposing batters that are home runs | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_hits_allowed_per_game | Hits allowed per game played | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_hits_per_nine | Pitchers, hits allowed per 9 innings pitched | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_holds | Pitchers, holds (report MUST INCLUDE PLAYER GROUPING) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_home_run_distance_against_average | Pitchers, average batted ball distance of home runs (among HR with recorded distances only) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_home_runs_allowed | Pitchers, number of home runs hit against them | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_home_runs_allowed_per_game | Home Runs allowed per Game | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_home_runs_allowed_three_run | Pitchers, HR allowed to opposing batters with two men on base. | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_home_runs_allowed_two_run | Pitchers, HR allowed to opposing batters with one man on base. | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_home_runs_per_nine | Pitchers, home runs allowed per 9 innings pitched | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_home_runs_solo | Pitchers, Solo Home Runs allowed to opposing batters | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_home_runs_solo_percentage | Pitchers, Pct. of Home Runs allowed to opposing batters that were Solo shots. | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_in_zone_contact_against_percentage | Pitchers, opposing batters percentage of swings where contact was made on pitches thrown in the zone. (Zones as defined by MLB, 1->9 is in-zone, 11->14 is out-of-zone) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_in_zone_swing_percentage | Pitchers, opposing batters swing percentage on pitches which were thrown in the zone. (Zones as defined by MLB, 1->9 is in-zone, 11->14 is out-of-zone) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_in_zone_whiff_percentage | Pitchers, opposing batters whiff percentage on pitches thrown in the zone. (Zones as defined by MLB, 1->9 is in-zone, 11->14 is out-of-zone) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_innings_pitched | Pitchers, innings pitched | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_innings_pitched_home | Pitchers, innings pitched in home games | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_innings_pitched_r_p | Pitchers, innings pitched as a reliever | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_innings_pitched_road | Pitchers, innings pitched on the road | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_innings_pitched_sp | Pitchers, innings pitched as a starter | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_intentional_walks | Pitchers, number of times intentionally walking a batter | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_knuckleballs_velocity_average | Pitchers, average knuckleball pitch speed (among pitches with recorded velocity only) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_knucklecurves_velocity_average | Pitchers, average knuckle curve pitch speed (among pitches with recorded velocity only) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_launch_angle_against_average | Pitchers, average launch angle by opposing batters on balls in play (among pitches with recorded launch angle only) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_lob_percentage | *MUST USE PITCHER QUALIFIER FILTER* Percentage of base runners that pitchers strand on base over the course of a season, does not use actual LOB statistic - uses calculation: LOB% = (H+BB+HBP-R)/(H+BB+HBP-(1.4*HR)) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_losses | Pitchers, times individually receiving a losing decision | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_offspeeds_velocity_average | Pitchers, average offspeed pitch type pitch speed (among pitches with recorded velocity only) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_on_base_percentage_against | pitchers, OBP allowed to opponent hitters (opponent on-base pct) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_on_base_percentage_against_home | Pitchers, OBP allowed in home games | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_on_base_percentage_against_lhb | pitchers, OBP allowed to opponent left-handed hitters (batting left-handed at time of the plate appearance) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_on_base_percentage_against_rhb | pitchers, OBP allowed to opponent right-handed hitters (batting right-handed at time of the plate appearance) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_on_base_percentage_against_road | Pitchers, OBP allowed in road games | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_on_base_plus_slugging_against | pitchers, OPS allowed to opponent hitters (opponent OPS) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_on_base_plus_slugging_against_home | Pitchers, OPS allowed in home games | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_on_base_plus_slugging_against_lhb | Pitchers, OPS allowed to opponent left-handed hitters (batting left-handed at time of the plate appearance) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_on_base_plus_slugging_against_rhb | Pitchers, OPS allowed to opponent right-handed hitters (batting right-handed at time of the plate appearance) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_on_base_plus_slugging_against_road | Pitchers, OPS allowed in road games | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_opponent_at_bats_runners_in_scoring_position | Number of At-Bats a pitcher surrenders with Runners In Scoring Position | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_opponent_average_home | Pitchers, batting average allowed in home games | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_opponent_average_home_road_differential | Difference between pitcher's batting average against in home games vs. road games, negative indicates better on the road | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_opponent_average_lhb | Pitchers, batting average allowed to opposing left-handed hitters (batting left-handed at time of the plate appearance) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_opponent_average_rhb | Pitchers, batting average allowed to opposing right-handed hitters (batting right-handed at time of the plate appearance) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_opponent_average_road | Pitchers, batting average allowed in road games | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_opponent_balls_in_play | Pitchers, number of balls in play by opposing batters | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_opponent_balls_in_play_per_game | Pitchers, number of balls in play per game by opposing batters. | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_opponent_batting_average | Pitchers, batting average allowed to opposing hitters (opponent batting average) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_opponent_batting_average_lhb_rhp_differential | Difference between pitcher's batting average allowed against left-handed batters vs. right-handed batters, negative indicates worse vs. right-handed batters. | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_opponent_batting_average_runners_in_scoring_position | Pitchers, batting average allowed (hits / at-bats) with Runners in Scoring position. | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_opponent_fly_balls | Pitchers, number of balls in play that were fly balls by opposing batters. (Note: does not include line drives or pop-ups) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_opponent_fly_balls_percentage | Pitchers, percentage of balls in play that were fly balls by opposing batters. (Note: does not include line drives or pop-ups) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_opponent_fly_pop_balls | Pitchers, number of balls in play by opposing batters that were either Fly Balls or Pop-Ups. (Note: does not include Line Drives) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_opponent_fly_pop_balls_percentage | Pitchers, percentage of balls in play by opposing batters that were either Fly Balls or Pop-Ups. (Note: does not include Line Drives) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_opponent_ground_balls | Pitchers, number of balls in play that were ground balls by opposing batters | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_opponent_ground_balls_air_balls_differential | Pitchers, Ratio of balls in play, Ground Balls to Air Balls. (Note: Air Balls = balls in play classified as Fly Ball, Line Drive, or Pop-Up) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_opponent_ground_balls_out_percentage | Percentage of opponent ground balls that are groundouts. | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_opponent_ground_balls_percentage | Pitchers, percentage of balls in play that were ground balls by opposing batters | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_opponent_ground_outs | Pitchers, number of times opponent batters grounded out. | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_opponent_hits_runners_in_scoring_position | Number of Hits a pitcher surrenders with Runners In Scoring Position | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_opponent_line_drives | Pitchers, number of balls in play that were line drives by opposing batters. | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_opponent_line_drives_outs | Pitchers, number of times opponent batters lined out. | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_opponent_line_drives_percentage | Pitchers, percentage of balls in play that were line drives by opposing batters. | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_opponent_on_base_percentage_home_road_differential | Difference between pitcher's OBP in home games vs. road games, negative indicates better on the road | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_opponent_on_base_plus_slugging_home_road_differential | Difference between pitcher's OPS in home games vs. road games, negative indicates better on the road | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_opponent_on_base_plus_slugging_lhb_rhp_differential | Difference between pitcher's OPS allowed against left-handed batters vs. right-handed batters, negative indicates worse vs. right-handed batters. | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_opponent_pop_ups | Pitchers, number of balls in play that were popped up by opposing batters | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_opponent_pop_ups_outs | Pitchers, number of times opponent batters popped out. | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_opponent_pop_ups_percentage | Pitchers, percentage of balls in play that were popped up by opposing batters | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_opponent_slugging_home_road_differential | Difference between pitcher's SLG in home games vs. road games, negative indicates better on the road | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_opponent_times_reached | Pitchers, number of times allowing opponent to reach base safely (H + BB + HBP) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_out_of_zone_contact_percentage_against | Pitchers, opposing batters percentage of swings where contact was made on pitches thrown out of the zone. (Zones as defined by MLB, 1->9 is in-zone, 11->14 is out-of-zone) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_out_of_zone_whiff_percentage | Pitchers, opposing batters whiff percentage on pitches thrown out of the zone. (Zones as defined by MLB, 1->9 is in-zone, 11->14 is out-of-zone) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_out_percentage | The percentage of batters gotten out by a pitcher. The inverse of Pitcher OBP against. | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_outs | Pitchers, number of outs pitched | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_relief_hits_allowed | Pitchers, hits allowed by relief pitchers | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_relief_home_runs_allowed | Pitchers, number of home runs hit against relief pitchers | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_relief_losses | Pitchers, time individually receiving decision = Loss, in relief | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_relief_pitches_thrown | Pitcher, number of pitches thrown in relief | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_relief_strikeouts | Pitchers, number of times striking out opposing batter in relief | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_relief_walks | Pitchers, number of times walking the opposing batter in relief | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_relief_wins | Pitchers, times individually receiving decision = Win, in relief | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_reliefs_innings_pitched_per_game | Pitchers, innings pitched (as a reliever) per game. | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_roe | Pitchers, number of batters who reached base via error | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_rointerference | Pitchers, opposing batters to reach base via defensive interference (usually catcher's interference, but some instances of interference by other players) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_runs_allowed_per_nine | Pitchers, total runs (both earned and unearned) allowed per 9 innings pitched | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_runs_support | Pitchers, own team's offensive runs scored for that pitcher in the game | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_runs_support_average | Pitchers, average runs of support per 9 Innings Pitched | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_runs_support_starter | Pitchers, own team's offensive runs scored for that pitcher in the game as starter (only includes games as the starting pitcher) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_runs_support_starter_average | Pitchers, average runs of support per 9 Innings Pitched as starter (only includes games as the starting pitcher) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_sacflies_against | Pitchers, number of times batters successfully hit a sacrifice fly against them | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_sachits_against | Pitchers, number of times batters successfully completed a sacrifice hit against them (NOTE: this is sacrifice bunts for 1955 and after. Prior to 1955, bunts and flies were not separated as distinct categories) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_saves | Pitchers, saves | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_saves_opportunities | Pitchers, number of save opportunities | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_saves_percentage | Pitchers, percentage of save opportunities that end in a save) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_screwballs_velocity_average | Pitchers, average screwball pitch speed (among pitches with recorded velocity only) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_shutouts | Pitchers, number of times an individual pitcher threw a complete-game shutout | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_singles_against | Pitchers, number of singles hit against them | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_sinkers_velocity_average | Pitchers, average 2-seam fastball pitch speed (among pitches with recorded velocity only) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_sliders_velocity_average | Pitchers, average slider pitch speed (among pitches with recorded velocity only) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_slugging_against | Pitchers, slugging percentage allowed to opposing hitters (opponent SLG) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_slugging_against_home | Pitchers, SLG allowed in home games | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_slugging_against_lhb | pitchers, SLG allowed to opponent left-handed hitters (batting left-handed at time of the plate appearance) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_slugging_against_rhb | pitchers, SLG allowed to opponent right-handed hitters (batting right-handed at time of the plate appearance) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_slugging_against_road | Pitchers, SLG allowed in road games | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_spin_rate_average | Pitchers, average spin rate on pitches thrown (among pitches with recorded spin rate only) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_splitters_velocity_average | Pitchers, average splitter pitch speed (among pitches with recorded velocity only) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_starting_hits_allowed | Pitchers, hits allowed as starting pitcher | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_starting_home_runs_allowed | Pitchers, number of home runs hit against starting pitchers | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_starting_innings_pitched_per_start | Pitchers, innings pitched (in starts only) per start | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_starting_losses | Pitchers, times individually receiving decision = Loss, as starter | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_starting_no_decisions | Starting Pitchers, times receiving no decision | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_starting_pitches_thrown | Pitcher, number of pitches thrown as starter | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_starting_strikeouts | Pitchers, number of times striking out batters as starter | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_starting_walks | Pitcher, number of times walking the opposing batter as starting pitcher | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_starting_wins | Pitchers, times individually receiving decision = 0, as starter | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_steals | Pitchers, stolen bases allowed while pitching | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_steals_attempts_against | Pitchers, number of stolen base attempts against (opponent's Steals + Caught Stealing) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_steals_home | Amount of times baserunners have successfully stolen home plate against a pitcher. | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_steals_second | Amount of times baserunners have successfully stolen 2nd base against a pitcher. | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_steals_third | Amount of times baserunners have successfully stolen 3rd base against a pitcher. | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_strikeouts | Pitchers, number of times striking out opposing batters | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_strikeouts_enforced_strikes | Number of pitcher strikeouts that ended on an enforced strike. | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_strikeouts_looking | Pitchers, number of strikeouts when looking (not swinging) at the strikeout pitch. | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_strikeouts_looking_percentage | Pitchers, percent of strikeouts when looking (not swinging) at the strikeout pitch. | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_strikeouts_per_game | Pitchers, number of times striking out opposing batters per game played | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_strikeouts_per_nine | Pitchers, Strikeouts per 9 innings pitched | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_strikeouts_percentage | Pitchers - Percentage of batters faced that end in a strikeout. | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_strikeouts_swinging | Pitchers, number of strikeouts when swinging at the strikeout pitch. | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_strikeouts_swinging_percentage | Pitchers, percent of strikeouts when swinging at the strikeout pitch. | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_strikeouts_to_walk_ratio | Pitchers, ratio of strikeouts to walks | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_strikeouts_to_walk_ratio_number | Pitchers, ratio of strikeouts to walks. Does not include raw totals of strikeouts and walks. IF A PITCHER HAS ZERO WALKS, THE RATIO WILL BE 0.00 | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_strikeouts_walks_differential | The difference in a pitchers strikeout percentage (batters faced that end in strikeouts) to their walk percentage (batters faced that end in walks). | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_strikeouts_walks_percentage | Pitchers - Percentage of batters faced that end in a strikeout minus percentage of batters faced that end in a walk. | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_swings | Pitchers, number of times batter swung at the pitch | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_swings_against_percentage | Pitchers, percentage of pitches thrown where batter swung at the pitch. | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_swings_contact_against | Pitchers, number of times batter swings at the pitch and makes contact | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_team_hits_allowed_percentage | Players percentage of team's pitching hits allowed. (Must have TEAM and PLAYER Groupings) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_teams_home_runs_allowed_percentage | Players percentage of team's pitching home runs allowed. (Player oHR / Team oHR) - NEEDS PLAYER AND TEAM GROUPING INCLUDED IN REPORTS | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_teams_strikeouts_percentage | Players percentage of team's pitching strikeouts. (Player pSO / Team pSO) - NEEDS PLAYER AND TEAM GROUPING INCLUDED IN REPORTS | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_three_true_outcomes | Pitchers, number of batters faced ending in a BB, K, or HR | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_three_true_outcomes_percentage | Pitchers, percentage of batters faced ending in a BB, K, or HR | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_times_chased | Pitchers, number of times opposing batter swings at a pitch out of the zone | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_total_bases | Pitchers, total bases allowed to batters on their hits | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_triples | Pitchers, number of triples hit against them | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_unearned_runs | Pitchers, runs charged to the pitcher that were Unearned Runs | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_velocity_average | Pitchers, average pitch speed (among pitches with recorded velocity only) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_walks | Pitchers, number of times walking the opposing batter | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_walks_enforced | Number of pitcher walks that ended on an enforced ball. | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_walks_per_game | Pitchers, number of times walking the opposing batter per game played | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_walks_per_nine | Pitchers, Walks per 9 Innings Pitched | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_walks_percentage | Pitchers - Percentage of batters faced that end in a walk. | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_whiffs | Pitchers, number of times batter swings at the pitch without making contact | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_whiffs_percentage | Pitchers, percentage of swings by opposing batters where contact was not made (swing+miss / swings) | MLB | general | CATEGORY_TO_MEASURE |
| pitcher_wins | Pitchers, times individually receiving decision = Win | MLB | general | CATEGORY_TO_MEASURE |
| pitcherbatting_average_balls_in_play_against | Pitchers, (H - HR) / (AB - K - HR + SF) | MLB | general | CATEGORY_TO_MEASURE |
| pitches_faced | Batters, number of pitches seen | MLB | general | CATEGORY_TO_MEASURE |
| pitches_faced_balls | Batters, number of pitches seen that were balls | MLB | general | CATEGORY_TO_MEASURE |
| pitches_faced_breaking_balls_percentage | Batters, percentage of total pitches seen that were breaking balls (all breaking ball types) (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |
| pitches_faced_changeups_percentage | Batters, percentage of total pitches seen that were changeups (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |
| pitches_faced_curveballs_percentage | Batters, percentage of total pitches seen that were curveballs (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |
| pitches_faced_cutters_percentage | Batters, percentage of total pitches seen that were cutters (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |
| pitches_faced_eephuses_percentage | Batters, percentage of total pitches seen that was an eephus (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |
| pitches_faced_fastballs_percentage | Batter, percentage of total pitches seen that were fastballs (all fastball types) (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |
| pitches_faced_forkballs_percentage | Batters, percentage of total pitches seen that were forkballs (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |
| pitches_faced_four_seams_percentage | Batters, percentage of total pitches seen that were 4 seam fastballs. (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |
| pitches_faced_in_zone_percentage | Batters, percentage of pitches seen that were thrown in the zone. (Zones as defined by MLB, 1->9 is in-zone, 11->14 is out-of-zone) | MLB | general | CATEGORY_TO_MEASURE |
| pitches_faced_knuckleballs_percentage | Batters, percentage of total pitches seen that were knuckleballs (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |
| pitches_faced_knucklecurves_percentage | Batters, percentage of total pitches seen that were knuckle curveballs (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |
| pitches_faced_offspeeds_percentage | Batter, percentage of total pitches faced that were offspeed (all offspeed types) (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |
| pitches_faced_out_of_zone_percentage | Batters, percentage of pitches seen that were thrown out of the zone. (Zones as defined by MLB, 1->9 is in-zone, 11->14 is out-of-zone) | MLB | general | CATEGORY_TO_MEASURE |
| pitches_faced_per_game | Pitches seen per Game | MLB | general | CATEGORY_TO_MEASURE |
| pitches_faced_per_pa | Pitches seen per Plate Appearance | MLB | general | CATEGORY_TO_MEASURE |
| pitches_faced_screwballs_percentage | Batters, percentage of total pitches seen that were screwballs (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |
| pitches_faced_sinkers_percentage | Batters, percentage of total pitches seen that were 2-seam fastballs or sinkers. (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |
| pitches_faced_sliders_percentage | Batters, percentage of total pitches seen that were sliders. (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |
| pitches_faced_slurves_percentage | Batters, percentage of total pitches seen that were slurves (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |
| pitches_faced_splitters_percentage | Batters, percentage of total pitches seen that were splitters (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |
| pitches_faced_strike | Batters, number of pitches seen that were strikes | MLB | general | CATEGORY_TO_MEASURE |
| pitches_faced_strikes_called_percentage | Hitting, percent of strikes faced that were called strikes (strikes looking, batter did not swing) | MLB | general | CATEGORY_TO_MEASURE |
| pitches_faced_strikes_percentage | Hitting, percent of pitches faced that were called strikes (strikes looking, batter did not swing) | MLB | general | CATEGORY_TO_MEASURE |
| pitches_faced_sweepers_percentage | Batters, percentage of total pitches seen that were sweepers (among pitches with recorded pitch type only) | MLB | general | CATEGORY_TO_MEASURE |
| pitches_per_inning | The number of pitches thrown per inning pitched. | MLB | general | CATEGORY_TO_MEASURE |
| pitches_thrown | Pitcher, number of pitches thrown | MLB | general | CATEGORY_TO_MEASURE |
| pitches_thrown_balls | Pitcher, number of pitches thrown that were balls (NOTE: includes balls issued on automatic intentional walks) | MLB | general | CATEGORY_TO_MEASURE |
| pitches_thrown_balls_percentage | Pitcher, percentage of pitches thrown that were balls | MLB | general | CATEGORY_TO_MEASURE |
| pitches_thrown_in_zone_percentage | Pitchers, percentage of pitches thrown that were thrown in the zone. (Zones as defined by MLB, 1->9 is in-zone, 11->14 is out-of-zone) | MLB | general | CATEGORY_TO_MEASURE |
| pitches_thrown_out_zone_percentage | Pitchers, percentage of pitches thrown that were thrown out of the zone. (Zones as defined by MLB, 1->9 is in-zone, 11->14 is out-of-zone) | MLB | general | CATEGORY_TO_MEASURE |
| pitches_thrown_per_batter | Pitchers, number of pitches thrown per batter faced | MLB | general | CATEGORY_TO_MEASURE |
| pitches_thrown_per_game | Pitcher, number of pitches thrown per game | MLB | general | CATEGORY_TO_MEASURE |
| pitches_thrown_starter_per_start | Pitcher, number of pitches thrown per game started | MLB | general | CATEGORY_TO_MEASURE |
| pitches_thrown_strikes | Pitcher, number of pitches thrown that were strikes | MLB | general | CATEGORY_TO_MEASURE |
| pitches_thrown_strikes_called | Pitcher, number of pitches thrown that were called strikes (strikes looking, batter did not swing) | MLB | general | CATEGORY_TO_MEASURE |
| pitches_thrown_strikes_called_percentage | Pitching, percent of pitches thrown that were called strikes (strikes looking, batter did not swing) | MLB | general | CATEGORY_TO_MEASURE |
| pitches_thrown_strikes_called_strikes_percentage | Pitching, percent of strikes thrown that were called strikes (strikes looking, batter did not swing) | MLB | general | CATEGORY_TO_MEASURE |
| pitches_thrown_strikes_percentage | Pitcher, percentage of pitches thrown that were strikes | MLB | general | CATEGORY_TO_MEASURE |
| pitches_thrown_strikes_swinging | Pitcher, number of pitches thrown that were swinging strikes | MLB | general | CATEGORY_TO_MEASURE |
| pitching_winning_percentage | Pitchers, win percentage (individual pitcher decisions only) | MLB | general | CATEGORY_TO_MEASURE |
| plate__apperances_risp__two__outs | Plate Appearances with Runners in Scoring Position and two outs. | MLB | general | CATEGORY_TO_MEASURE |
| plate_appearances | Batters, number of plate appearances | MLB | general | CATEGORY_TO_MEASURE |
| plate_apperances_bases_empty | Plate Appearances with no runners on base. | MLB | general | CATEGORY_TO_MEASURE |
| plate_apperances_bases_loaded | Plate appearances with the bases loaded. | MLB | general | CATEGORY_TO_MEASURE |
| plate_apperances_home | Batters, number of plate appearances at home | MLB | general | CATEGORY_TO_MEASURE |
| plate_apperances_per_home_run | Batters, Ratio of PA to HR (lower number is better, meaning a higher rate of HR hit) | MLB | general | CATEGORY_TO_MEASURE |
| plate_apperances_per_rbi | Batters, Ratio of PA to RBI (lower number is better, meaning a higher rate of RBI) | MLB | general | CATEGORY_TO_MEASURE |
| plate_apperances_per_strikeout | Batters, Ratio of PA to strikeouts (lower number means a higher strikeout rate) | MLB | general | CATEGORY_TO_MEASURE |
| plate_apperances_per_walk | Batters, Ratio of PA to BB (lower number means a higher walk rate) | MLB | general | CATEGORY_TO_MEASURE |
| plate_apperances_road | Batters, number of plate appearances on the road | MLB | general | CATEGORY_TO_MEASURE |
| plate_apperances_runners_in_scoring_position_two_outs | Plate Appearances with Runners in Scoring Position. | MLB | general | CATEGORY_TO_MEASURE |
| plate_apperances_runners_on_base | Plate appearances with runners on base. | MLB | general | CATEGORY_TO_MEASURE |
| player_team_stolen_bases_percentage | Percentage of a team's stolen bases accounted by an individual player (MUST include Team Split) | MLB | general | CATEGORY_TO_MEASURE |
| players | Count distinct player_id | MLB | general | CATEGORY_TO_MEASURE |
| popups | Batters, number of balls in play that were popped up. | MLB | general | CATEGORY_TO_MEASURE |
| popups_outs | Batters, number of times popping out. | MLB | general | CATEGORY_TO_MEASURE |
| popups_percentage | Batters, percentage of balls in play that were popped up. | MLB | general | CATEGORY_TO_MEASURE |
| power_speed_number | Batters, (2 * HR * SB) / (HR + SB) | MLB | general | CATEGORY_TO_MEASURE |
| pulled_percentage | The percentage of batted ball events that were pulled. | MLB | general | CATEGORY_TO_MEASURE |
| put_away_rate | Pitchers, the rate of two-strike pitches that result in a strikeout. | MLB | general | CATEGORY_TO_MEASURE |
| quality_starts | Starting Pitchers, starts with 6.0+ IP and 3 or fewer ER allowed | MLB | general | CATEGORY_TO_MEASURE |
| quality_starts_percentage | Percentage of a Starting Pitcher's Starts that were Quality Starts (starts with 6.0+ IP and 3 or fewer ER allowed) | MLB | general | CATEGORY_TO_MEASURE |
| range_factor_per_game | Fielders, putouts and assists divided by games played | MLB | general | CATEGORY_TO_MEASURE |
| range_factor_per_nine | Fielders, putouts and assists divided by innings played | MLB | general | CATEGORY_TO_MEASURE |
| rbi_opponent | Pitching, runs batted in by the opposing team | MLB | general | CATEGORY_TO_MEASURE |
| rbis | Batters, Runs Batted In | MLB | general | CATEGORY_TO_MEASURE |
| rbis_home | Batters, runs batted in, in home games | MLB | general | CATEGORY_TO_MEASURE |
| rbis_road | Batters, runs batted in, on the road | MLB | general | CATEGORY_TO_MEASURE |
| reached_on_error | Batters, number of times reaching base via error | MLB | general | CATEGORY_TO_MEASURE |
| reached_on_interference | Batters, reached base via defensive interference (usually catcher's interference, but some instances of interference by other players) | MLB | general | CATEGORY_TO_MEASURE |
| relief_appearances | Players, number of games pitched in relief | MLB | general | CATEGORY_TO_MEASURE |
| road_losses | Losses in road games | MLB | general | CATEGORY_TO_MEASURE |
| road_wins | Wins in road games | MLB | general | CATEGORY_TO_MEASURE |
| runs | Runs scored | MLB | general | CATEGORY_TO_MEASURE |
| runs_allowed | Pitchers, runs scored that were charged to the pitcher | MLB | general | CATEGORY_TO_MEASURE |
| runs_allowed_eight_inning | Runs allowed in the 8th inning | MLB | general | CATEGORY_TO_MEASURE |
| runs_allowed_fifth_inning | Runs allowed in the 5th inning | MLB | general | CATEGORY_TO_MEASURE |
| runs_allowed_first_inning | Runs allowed in the 1st inning | MLB | general | CATEGORY_TO_MEASURE |
| runs_allowed_fourth_inning | Runs allowed in the 4th inning | MLB | general | CATEGORY_TO_MEASURE |
| runs_allowed_ninth_inning | Runs allowed in the 9th inning | MLB | general | CATEGORY_TO_MEASURE |
| runs_allowed_per_game | Runs Allowed per Game | MLB | general | CATEGORY_TO_MEASURE |
| runs_allowed_per_game_differential | Difference between runs allowed per game in home games vs away games, negative indicates a higher value on the road | MLB | general | CATEGORY_TO_MEASURE |
| runs_allowed_per_home | Runs scored by the opponent per home game | MLB | general | CATEGORY_TO_MEASURE |
| runs_allowed_per_road | Runs scored by the opponent per away game | MLB | general | CATEGORY_TO_MEASURE |
| runs_allowed_r_p | Pitchers, runs scored that were charged to relief pitcher | MLB | general | CATEGORY_TO_MEASURE |
| runs_allowed_s_p | Pitchers, runs scored that were charged to starting pitcher | MLB | general | CATEGORY_TO_MEASURE |
| runs_allowed_second_inning | Runs allowed in the 2nd inning | MLB | general | CATEGORY_TO_MEASURE |
| runs_allowed_seventh_inning | Runs allowed in the 7th inning | MLB | general | CATEGORY_TO_MEASURE |
| runs_allowed_sixth_inning | Runs allowed in the 6th inning | MLB | general | CATEGORY_TO_MEASURE |
| runs_allowed_tenth_inning | Runs allowed in the 10th inning | MLB | general | CATEGORY_TO_MEASURE |
| runs_allowed_tenth_plus_inning | Runs allowed in all innings 10th and later | MLB | general | CATEGORY_TO_MEASURE |
| runs_allowed_third_inning | Runs allowed in the 3rd inning | MLB | general | CATEGORY_TO_MEASURE |
| runs_created | Estimate of a player's offensive contribution in terms of total runs. Combining ability to get on base with ability to for extra base hits, divided by player's total opportunities. | MLB | general | CATEGORY_TO_MEASURE |
| runs_differential | Teams, run differential, scored minus allowed | MLB | general | CATEGORY_TO_MEASURE |
| runs_eight_inning | Runs scored in the 8th inning | MLB | general | CATEGORY_TO_MEASURE |
| runs_fifth_inning | Runs scored in the 5th inning | MLB | general | CATEGORY_TO_MEASURE |
| runs_first_inning | Runs scored in the 1st inning | MLB | general | CATEGORY_TO_MEASURE |
| runs_fourth_inning | Runs scored in the 4th inning | MLB | general | CATEGORY_TO_MEASURE |
| runs_home_road_per_game_differential | Difference between runs scored per game in home games vs away games, negative indicates a higher value on the road | MLB | general | CATEGORY_TO_MEASURE |
| runs_ninth_inning | Runs scored in the 9th inning | MLB | general | CATEGORY_TO_MEASURE |
| runs_per_game | Runs scored per game | MLB | general | CATEGORY_TO_MEASURE |
| runs_per_game_away | Runs scored per away game | MLB | general | CATEGORY_TO_MEASURE |
| runs_per_game_home | Runs scored per home game | MLB | general | CATEGORY_TO_MEASURE |
| runs_pergame_differential | Teams, run differential per game (scored minus allowed) | MLB | general | CATEGORY_TO_MEASURE |
| runs_produced | Batters, Runs + RBI - HR | MLB | general | CATEGORY_TO_MEASURE |
| runs_second_inning | Runs scored in the 2nd inning | MLB | general | CATEGORY_TO_MEASURE |
| runs_seventh_inning | Runs scored in the 7th inning | MLB | general | CATEGORY_TO_MEASURE |
| runs_sixth_inning | Runs scored in the 6th inning | MLB | general | CATEGORY_TO_MEASURE |
| runs_tenth_inning | Runs scored in the 10th inning | MLB | general | CATEGORY_TO_MEASURE |
| runs_tenth_plus_inning | Runs scored in all innings 10th and later | MLB | general | CATEGORY_TO_MEASURE |
| runs_third_inning | Runs scored in the 3rd inning | MLB | general | CATEGORY_TO_MEASURE |
| runs_two_outs | Runs scored with two outs in an inning. | MLB | general | CATEGORY_TO_MEASURE |
| runs_win_loss_per_game_differential | Difference between runs scored per game in won games vs lost games (Negative value implies a higher value in losses) | MLB | general | CATEGORY_TO_MEASURE |
| runs_with_one_outs_percentages | Batting, Percentage of runs scored when there are one out. | MLB | general | CATEGORY_TO_MEASURE |
| runs_with_two_outs_percentages | Batting, Percentage of runs scored when there are two outs. | MLB | general | CATEGORY_TO_MEASURE |
| runs_with_zero_outs_percentages | Batting, Percentage of runs scored when there are no outs. | MLB | general | CATEGORY_TO_MEASURE |
| sacflies | Batters, sacrifice flies hit | MLB | general | CATEGORY_TO_MEASURE |
| sachits | Batters, number of successful sacrifice hits (NOTE: this is sacrifice bunts for 1955 and after. Prior to 1955 bunts and flies were not separated as distinct categories) | MLB | general | CATEGORY_TO_MEASURE |
| seasons | Season counts | MLB | general | CATEGORY_TO_MEASURE |
| secondary_average | Batters, A measurement of hitting performance that seeks to evaluate the number of bases a player gained independent of batting average. (BB + (TB-H) + (SB-CS))/AB | MLB | general | CATEGORY_TO_MEASURE |
| singles | Batters, singles hit | MLB | general | CATEGORY_TO_MEASURE |
| singles_per_game | Singles per game | MLB | general | CATEGORY_TO_MEASURE |
| slugging | Batters, slugging percentage (total bases divided by at-bats) | MLB | general | CATEGORY_TO_MEASURE |
| slugging_day | Batters, slugging percentage in day games | MLB | general | CATEGORY_TO_MEASURE |
| slugging_day_night_differential | Difference between slugging pct. in day games vs. night games, negative indicates better at night | MLB | general | CATEGORY_TO_MEASURE |
| slugging_differential | Difference between batting SLG and pitching SLG allowed. A negative value indicates a higher SLG allowed. | MLB | general | CATEGORY_TO_MEASURE |
| slugging_home | Batters, slugging percentage in home games | MLB | general | CATEGORY_TO_MEASURE |
| slugging_home_road_differential | Difference between slugging pct. in home games vs. road games, negative indicates better on the road | MLB | general | CATEGORY_TO_MEASURE |
| slugging_lhp_rhp_differential | Difference between slugging percentage against left-handed pitchers vs. right-handed pitchers, negative indicates better vs. right-handed pitchers. | MLB | general | CATEGORY_TO_MEASURE |
| slugging_night | Batters, slugging percentage in night games | MLB | general | CATEGORY_TO_MEASURE |
| slugging_road | Batters, slugging percentage in road games | MLB | general | CATEGORY_TO_MEASURE |
| slugging_vs_lhp | Batters, slugging percentage vs. left-handed pitchers | MLB | general | CATEGORY_TO_MEASURE |
| slugging_vs_rhp | Batters, slugging percentage vs. right-handed pitchers | MLB | general | CATEGORY_TO_MEASURE |
| spin_rate_faced_average | Batters, average spin rate seen on pitches faced (among pitches with recorded spin rate only) | MLB | general | CATEGORY_TO_MEASURE |
| steals | Runners, successful stolen bases | MLB | general | CATEGORY_TO_MEASURE |
| steals_and_home_runs | Stolen bases plus home runs | MLB | general | CATEGORY_TO_MEASURE |
| steals_attempts | Runners, number of stolen base attempts (Steals + Caught Stealing) | MLB | general | CATEGORY_TO_MEASURE |
| steals_attempts_per_game | Runners, stolen base attempts per game played | MLB | general | CATEGORY_TO_MEASURE |
| steals_home | Amount of Stolen Bases player has when successfully stealing home plate. | MLB | general | CATEGORY_TO_MEASURE |
| steals_per_game | Runners, successful stolen bases per game played | MLB | general | CATEGORY_TO_MEASURE |
| steals_percentage | Runners, percentage of successful Stolen Base Attempts | MLB | general | CATEGORY_TO_MEASURE |
| steals_second | Amount of Stolen Bases player has when successfully stealing second base. | MLB | general | CATEGORY_TO_MEASURE |
| steals_third | Amount of Stolen Bases player has when successfully stealing third base. | MLB | general | CATEGORY_TO_MEASURE |
| stolen_base_differential | Teams, difference between stolen bases on offense and stolen bases allowed on defense | MLB | general | CATEGORY_TO_MEASURE |
| strikeouts | Batters, times striking out at the plate | MLB | general | CATEGORY_TO_MEASURE |
| strikeouts_differential | Teams, difference between pitching strikeouts and batting strikeouts. NOTE: A positive number means more pitching strikeouts than batting strikeouts | MLB | general | CATEGORY_TO_MEASURE |
| strikeouts_enforced_strike | A strikeout that ended on an enforced strike. | MLB | general | CATEGORY_TO_MEASURE |
| strikeouts_lookings | Batters, times strikeout out when looking (not swinging) at the strikeout pitch. | MLB | general | CATEGORY_TO_MEASURE |
| strikeouts_per_game | Batting Strikeouts per Games Played | MLB | general | CATEGORY_TO_MEASURE |
| strikeouts_per_game_differential | Teams, difference between pitching strikeouts and batting strikeouts per game. NOTE: A positive number means more pitching strikeouts than batting strikeouts | MLB | general | CATEGORY_TO_MEASURE |
| strikeouts_percentage | Batters - Percentage of plate appearances that end in a strikeout. | MLB | general | CATEGORY_TO_MEASURE |
| strikeouts_swinging | Batters, times strikeout out when swinging at the strikeout pitch. | MLB | general | CATEGORY_TO_MEASURE |
| strikeouts_to_walk_ratio | Batters, ratio of strikeouts to walks | MLB | general | CATEGORY_TO_MEASURE |
| strikeouts_walks_percentage | Batters - Percentage of plate appearances that end in a strikeout minus percentage of plate appearances that end in a walk, negative indicates the batter's walk rate is higher than their strikeout rate. | MLB | general | CATEGORY_TO_MEASURE |
| sweet_spot_balls_hit | Batters, balls in play hit with a launch angle between 8 and 32 degrees | MLB | general | CATEGORY_TO_MEASURE |
| sweet_spot_percentage | Batters, percentage of balls in play hit with a launch angle between 8 and 32 degrees | MLB | general | CATEGORY_TO_MEASURE |
| team_blown_lead_losses | Losses after leading at any point in game | MLB | general | CATEGORY_TO_MEASURE |
| team_blown_saves | Team blown saves - can be multiple per game | MLB | general | CATEGORY_TO_MEASURE |
| team_comeback_wins | Wins after trailing at any point in game | MLB | general | CATEGORY_TO_MEASURE |
| team_earned_runs_allowed | Teams, Earned Runs charged to the team (NOTE: may differ from individual Pitcher ER due to Section 10.18 (i) of the Scoring rules) | MLB | general | CATEGORY_TO_MEASURE |
| team_era | Teams, Earned Runs allowed per 9 IP | MLB | general | CATEGORY_TO_MEASURE |
| team_era_home | Teams, ERA in home games | MLB | general | CATEGORY_TO_MEASURE |
| team_era_home_road_diffential | Teams, Difference between team ERA at home vs. away | MLB | general | CATEGORY_TO_MEASURE |
| team_era_road | Teams, ERA in road games | MLB | general | CATEGORY_TO_MEASURE |
| team_holds | Team total for holds by relief pitchers | MLB | general | CATEGORY_TO_MEASURE |
| team_left_on_base | Teams, number of runners left on base | MLB | general | CATEGORY_TO_MEASURE |
| team_left_on_base_per_game | Teams, number of runners left on base per game | MLB | general | CATEGORY_TO_MEASURE |
| team_losses | Teams, games lost | MLB | general | CATEGORY_TO_MEASURE |
| team_runs_allowed_for_home_runs_percentage | Teams, percentage of runs allowed to opponents which came via RBI from allowing a home run. | MLB | general | CATEGORY_TO_MEASURE |
| team_runs_for_home_runs | Teams, total number of runs scored which came via RBI from hitting a home run. | MLB | general | CATEGORY_TO_MEASURE |
| team_runs_for_home_runs_percentage | Teams, percentage of runs scored which came via RBI from hitting a home run. | MLB | general | CATEGORY_TO_MEASURE |
| team_saves | Teams, saves | MLB | general | CATEGORY_TO_MEASURE |
| team_saves_opportunities | Teams, number of save opportunities | MLB | general | CATEGORY_TO_MEASURE |
| team_saves_percentage | Pitchers, percentage of save opportunities that end in a save | MLB | general | CATEGORY_TO_MEASURE |
| team_shutouts | Teams, number of games allowing zero runs to opponent | MLB | general | CATEGORY_TO_MEASURE |
| team_ties | Teams, number of games tied | MLB | general | CATEGORY_TO_MEASURE |
| team_walkoffs_losses | Number of games where the team lost via walkoff (opponent team won by walkoff) | MLB | general | CATEGORY_TO_MEASURE |
| team_walkoffs_wins | Number of games where the team won via walkoff | MLB | general | CATEGORY_TO_MEASURE |
| team_win_percentage | Teams, winning percentage (W / (W+L) ). NOTE: ties are not included in any way in winning pct. calculations, but ties do count toward games played and all statistics accumulated in ties are official | MLB | general | CATEGORY_TO_MEASURE |
| team_wins | Teams, games won | MLB | general | CATEGORY_TO_MEASURE |
| teams | Count of the distinct teams included in the given groupings | MLB | general | CATEGORY_TO_MEASURE |
| teams_runs_allowed_for_home_runs | Teams, total number of runs allowed to opponents which came via RBI from allowing a home run. | MLB | general | CATEGORY_TO_MEASURE |
| teams_times_shutout | Teams, number of games when scoring 0 runs | MLB | general | CATEGORY_TO_MEASURE |
| three_true_outcomes | Batters, number of PA that end in a BB, K, or HR | MLB | general | CATEGORY_TO_MEASURE |
| three_true_outcomes_percentage | Batters, percentage of PA that end in a BB, K, or HR | MLB | general | CATEGORY_TO_MEASURE |
| times_reached | Batters, number of times reaching base safely (H + BB + HBP) | MLB | general | CATEGORY_TO_MEASURE |
| total_bases | Batters, total bases accrued on hits (singles + (2*doubles) + (3*triples) + (4*HR)) | MLB | general | CATEGORY_TO_MEASURE |
| total_distance_average | Batters, average batted ball distance (among pitches with recorded distances only) | MLB | general | CATEGORY_TO_MEASURE |
| total_over_attempts | Percentage of total Team Games that hit the over/under line | MLB | general | CATEGORY_TO_MEASURE |
| total_over_games | Total Team Games where the Over/Under Line was hit | MLB | general | CATEGORY_TO_MEASURE |
| total_under_attempts | Percentage of total Team Games played that did not meet the over/under line | MLB | general | CATEGORY_TO_MEASURE |
| total_under_games | Number of total Team Games played where the over/under line was not hit | MLB | general | CATEGORY_TO_MEASURE |
| triples | Batters, number of triples hit | MLB | general | CATEGORY_TO_MEASURE |
| triples_per_game | Triples per game | MLB | general | CATEGORY_TO_MEASURE |
| venue | Number of different ballparks in this grouping. | MLB | general | CATEGORY_TO_MEASURE |
| walkoff_hits | Batters, walkoff hits (single, double, triple, or home run) | MLB | general | CATEGORY_TO_MEASURE |
| walkoff_home_runs | Batters, walkoff home runs | MLB | general | CATEGORY_TO_MEASURE |
| walkoff_rbis | Batters, walkoff RBI (any walkoff result in which the run that won the game was an RBI for the batter) | MLB | general | CATEGORY_TO_MEASURE |
| walks | Batters, number of walks | MLB | general | CATEGORY_TO_MEASURE |
| walks_differential | Teams, difference between batting walks drawn and pitching walks issued | MLB | general | CATEGORY_TO_MEASURE |
| walks_enforced_ball | Number of hitter walks that ended on an enforced ball. | MLB | general | CATEGORY_TO_MEASURE |
| walks_per_game | Batters, number of walks per game played | MLB | general | CATEGORY_TO_MEASURE |
| walks_percentage | Batters - Percentage of plate appearances that end in a walk. | MLB | general | CATEGORY_TO_MEASURE |
| war | Season-level statistic measuring a players wins above replacement as a player (hitting+pitching) | MLB | general | CATEGORY_TO_MEASURE |
| war_baserunning | Season-level statistic measuring a players wins above replacement as a baserunner. | MLB | general | CATEGORY_TO_MEASURE |
| war_batter | Season-level statistic measuring a players wins above replacement as an offensive player. | MLB | general | CATEGORY_TO_MEASURE |
| war_batter_per_full_season | A batter's WAR in a given season projected over 162 games. ONLY INCLUDES BATTING WAR - DOES NOT INCLUDE PITCHING WAR | MLB | general | CATEGORY_TO_MEASURE |
| war_batting | Season-level statistic measuring a players wins above replacement as a batter (hitting only) | MLB | general | CATEGORY_TO_MEASURE |
| war_fielding | Season-level statistic measuring a players wins above replacement as a fielder. | MLB | general | CATEGORY_TO_MEASURE |
| war_pitching | Season-level statistic measuring a players wins above replacement as a pitcher. | MLB | general | CATEGORY_TO_MEASURE |
| whip | Pitchers, Walks + Hits allowed per Inning Pitched | MLB | general | CATEGORY_TO_MEASURE |
| whip_home | Pitchers, WHIP in home games | MLB | general | CATEGORY_TO_MEASURE |
| whip_home_road_differential | Difference between WHIP in home games vs. road games, negative indicates better at home | MLB | general | CATEGORY_TO_MEASURE |
| whip_road | Pitchers, WHIP in road games | MLB | general | CATEGORY_TO_MEASURE |
| whip_three_decimals | Pitchers, Walks + Hits allowed per Inning Pitched (formatted to include the thousandth place) | MLB | general | CATEGORY_TO_MEASURE |
| wild_pitches | Pitchers, number of wild pitches thrown. | MLB | general | CATEGORY_TO_MEASURE |
| winning_percentage_day | Winning percent in day games | MLB | general | CATEGORY_TO_MEASURE |
| winning_percentage_day_night_differential | The difference between a team's winning percentage in day games vs. night games, negative indicates better at night | MLB | general | CATEGORY_TO_MEASURE |
| winning_percentage_favorite | Team Winning percentage as the Moneyline Favorite | MLB | general | CATEGORY_TO_MEASURE |
| winning_percentage_home | Winning percent in home games | MLB | general | CATEGORY_TO_MEASURE |
| winning_percentage_home_road_differential | The difference between a team's winning percentage in home games vs. road games, negative indicates better on the road | MLB | general | CATEGORY_TO_MEASURE |
| winning_percentage_lhp | Winning percent in games where the opposing starting pitcher was left-handed | MLB | general | CATEGORY_TO_MEASURE |
| winning_percentage_night | Winning percent in night games | MLB | general | CATEGORY_TO_MEASURE |
| winning_percentage_post_all_star | Winning percent after the All-Star break. (NOTE: first All-Star game was in 1933) | MLB | general | CATEGORY_TO_MEASURE |
| winning_percentage_pre_all_star | Winning percent before the All-Star break. (NOTE: first All-Star game was in 1933) | MLB | general | CATEGORY_TO_MEASURE |
| winning_percentage_pre_post_all_star_differential | The difference between a team's winning percentage before vs. after the All-Star break. A positive number indicates better after the break, negative indicates better before the break | MLB | general | CATEGORY_TO_MEASURE |
| winning_percentage_rhp | Winning percent in games where the opposing starting pitcher was right-handed | MLB | general | CATEGORY_TO_MEASURE |
| winning_percentage_rhp_lhp_differential | The difference between a team's winning percentage in games where the opposing starting pitcher was right-handed vs. left-handed, negative indicates better vs. left-handers | MLB | general | CATEGORY_TO_MEASURE |
| winning_percentage_road | Winning percent in road games | MLB | general | CATEGORY_TO_MEASURE |
| winning_percentage_underdog | Team Winning Percentage in games in which the team was the moneyline underdog | MLB | general | CATEGORY_TO_MEASURE |
| wins_day | Wins in day games | MLB | general | CATEGORY_TO_MEASURE |
| wins_lhp | Wins in games where the opposing starting pitcher was left-handed | MLB | general | CATEGORY_TO_MEASURE |
| wins_night | Wins in night games | MLB | general | CATEGORY_TO_MEASURE |
| wins_post_all_star | Wins after the All-Star break. (NOTE: first All-Star game was in 1933) | MLB | general | CATEGORY_TO_MEASURE |
| wins_pre_all_star | Wins before the All-Star break. (NOTE: first All-Star game was in 1933) | MLB | general | CATEGORY_TO_MEASURE |
| wins_rhp | Wins in games where the opposing starting pitcher was right-handed | MLB | general | CATEGORY_TO_MEASURE |
| balks | Pitcher Balks | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| batters_faced | Pitchers number of batters faced | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| called_strike | Total number of called strikes plus whiffs divided by total pitches. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| enforced_balls_pitcher | The number of times a pitcher was the victim of an enforced ball | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| enforced_strikes_pitcher | The number of times a pitcher was the beneficiary of an enforced strike | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| first_pitch_balls_thrown | Pitchers, first pitch of plate appearance called a ball | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| first_pitch_hits_allowed | Pitchers, hits allowed on the first pitch of an at-bat | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| first_pitch_home_run_allowed | Pitchers, home runs allowed on the first pitch of an at-bat | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| first_pitch_strikes_thrown | Pitchers, strikes thrown on the first pitch of a plate appearance | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| first_pitch_strikes_thrown_percentage | Pitchers, percent of plate appearances to start with a strike | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| first_pitch_swing_percentage_against | Pitchers, Percentage of swings against on the first pitch of a plate appearance | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| first_pitch_swings_agaist | Pitchers, swings against on the first pitch of an at-bat | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| game_score | Metric devised by Bill James to determine the strength of a pitcher in any particular baseball game | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| gamelog_era | Windowed Earned Run Average based on the level of stats that are being viewed. For example, when viewing by games, this is the earned run average at that point in the season. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| gamelog_era_team | Windowed Team Earned Run Average based on the level of stats that are being viewed. For example, when viewing by games, this is the earned run average at that point in the season. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| gamelog_whip | Windowed Pitcher WHIP based on the level of stats that are being viewed. For example, when viewing by games, this is the pitcher WHIP at that point in the season. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| games_pitched | Number of games appearing as pitcher | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| games_pitched_total | Total Games Pitched - When grouping by team, this will count every individual player game pitched as one instead of showing only one per game. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| games_started | Number of games starting at pitcher | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| hits_allowed_extra_bases_percentage | Pitchers, percentage of hits allowed that are extra-base hits - (2B, 3B, HR)/Hits | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| inherited_runners | Relief Pitchers, the number of runners on base when the pitcher enters the game | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| inherited_runners_scored | Relief Pitchers, the number of inherited runners allowed to score | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| inherited_runners_scored_percentage | Relief Pitchers, the percentage of inherited runners allowed to score | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| p_pko | Pitchers, number of times picking off a baserunner | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitch_velocity_faced_average | Batters, average pitch speed faced (among pitches with recorded velocity only) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_air_balls_against | Pitchers, number of balls in play by opposing batters that were hit in the air. (Fly Balls + Line Drives + Pop-Ups) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_air_balls_against_percentage | Pitchers, percentage of balls in play by opposing batters that were hit in the air. (Fly Balls + Line Drives + Pop-Ups) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_at_bats_against | Pitchers, number of batters faced which resulted in an at-bat | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_blown_saves | Pitches, blown saves | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_breaking_balls_velocity_average | Pitchers, average breaking ball pitch type pitch speed (among pitches with recorded velocity only) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_caught_stealing | Pitchers, number of runners caught stealing | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_changeups_velocity_average | Pitchers, average changeup pitch speed (among pitches with recorded velocity only) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_chase_percentage | Pitchers, opposing batters' swing percentage on pitches out of the zone. (Zones as defined by MLB, 1->9 is in-zone, 11->14 is out-of-zone) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_complete_games | Pitcher complete games (started and finished the game) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_contact_against_percentage | Pitchers, percentage of swings by opposing batters where contact was made (swings with contact / swings) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_curveballs_velocity_average | Pitchers, average curveball pitch speed (among pitches with recorded velocity only) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_cutters_velocity_average | Pitchers, average cut fastball pitch speed (among pitches with recorded velocity only) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_decisions | Pitchers, times individually receiving a winning or losing decision | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_doubles_against | Pitchers, number of doubles hit against them | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_earned_runs_allowed | Pitchers, Earned Runs charged to the individual pitcher (NOTE: may differ from Team ER due to Section 10.18 (i) of the Scoring rules) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_earned_runs_relieiver | Pitchers, Earned Runs charged to the individual pitcher in games as a relief pitcher | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_earned_runs_starter | Pitchers, Earned Runs charged to the individual pitcher in games as the starting pitcher | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_eephuses_velocity_average | Pitchers, average eephus pitch speed (among pitches with recorded velocity only) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_era | PItchers, Earned Runs allowed per 9 IP | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_era_home_road_differential | Pitchers, Difference between ERA at home vs. away | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_era_reliver | Pitchers, Earned Runs allowed per 9 IP in games as a relief pitcher | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_era_starter | Pitchers, Earned Runs allowed per 9 IP in games as the starting pitcher | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_era_starter_reliever_differential | ERA Differential between Starting and Relieving | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_exit_velocity_against_average | Pitchers, average exit velocity by opposing batters on balls in play (among pitches with recorded exit velocity only) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_extra_base_hits | Pitchers, number of extra base hits allowed to opposing batters (2B + 3B + HR) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_fastballs_velocity_average | Pitchers, average fastball pitch type pitch speed (among pitches with recorded velocity only) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_forkballs_velocity_average | Pitchers, average forkball pitch speed (among pitches with recorded velocity only) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_foul_balls | Number of foul balls against or pitcher or the pitching team | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_four_seam_fastballs_velocity_average | Pitchers, average 4-seam fastball pitch speed (among pitches with recorded velocity only) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_grand_slams_allowed | Pitchers, grand slam home runs allowed | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_groundball_double_plays | Pitchers ground-ball double-plays induced. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_groundballto_fly_out_ratio | Pitchers, Ratio of balls in play, Ground Balls to Fly (F+P). (Note: Fly = balls in play classified as either Fly Ball or Pop-Up) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_h_r_to_fly_percentage | Pitchers, the percentage of fly balls + line drives by opposing batters that were home runs | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_hard_hit_balls_against | Pitchers, number of batted balls in play by opposing batters with an exit velocity (launch speed) of 95 MPH or higher | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_hard_hit_balls_against_percentage | Pitchers, percentage of batted balls in play by opposing batters with an exit velocity (launch speed) of 95 MPH or higher | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_hit_by_pitches | Pitchers, number of batters hit by pitch | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_hit_distance_against_average | Pitchers, average batted ball distance allowed to opposing batters (among pitches with recorded distances only) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_hits_allowed | Pitchers, hits allowed | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_hits_allowed_for_home_runs_percentage | Pitchers, percentage of hits allowed to opposing batters that are home runs | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_hits_allowed_per_game | Hits allowed per game played | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_hits_per_nine | Pitchers, hits allowed per 9 innings pitched | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_holds | Pitchers, holds (report MUST INCLUDE PLAYER GROUPING) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_home_run_distance_against_average | Pitchers, average batted ball distance of home runs (among HR with recorded distances only) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_home_runs_allowed | Pitchers, number of home runs hit against them | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_home_runs_allowed_three_run | Pitchers, HR allowed to opposing batters with two men on base. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_home_runs_allowed_two_run | Pitchers, HR allowed to opposing batters with one man on base. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_home_runs_per_nine | Pitchers, home runs allowed per 9 innings pitched | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_home_runs_solo | Pitchers, Solo Home Runs allowed to opposing batters | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_home_runs_solo_percentage | Pitchers, Pct. of Home Runs allowed to opposing batters that were Solo shots. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_in_zone_contact_against_percentage | Pitchers, opposing batters percentage of swings where contact was made on pitches thrown in the zone. (Zones as defined by MLB, 1->9 is in-zone, 11->14 is out-of-zone) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_in_zone_swing_percentage | Pitchers, opposing batters swing percentage on pitches which were thrown in the zone. (Zones as defined by MLB, 1->9 is in-zone, 11->14 is out-of-zone) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_in_zone_whiff_percentage | Pitchers, opposing batters whiff percentage on pitches thrown in the zone. (Zones as defined by MLB, 1->9 is in-zone, 11->14 is out-of-zone) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_innings_pitched | Pitchers, innings pitched | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_innings_pitched_r_p | Pitchers, innings pitched as a reliever | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_innings_pitched_sp | Pitchers, innings pitched as a starter | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_intentional_walks | Pitchers, number of times intentionally walking a batter | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_knuckleballs_velocity_average | Pitchers, average knuckleball pitch speed (among pitches with recorded velocity only) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_knucklecurves_velocity_average | Pitchers, average knuckle curve pitch speed (among pitches with recorded velocity only) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_launch_angle_against_average | Pitchers, average launch angle by opposing batters on balls in play (among pitches with recorded launch angle only) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_lob_percentage | *MUST USE PITCHER QUALIFIER FILTER* Percentage of base runners that pitchers strand on base over the course of a season, does not use actual LOB statistic - uses calculation: LOB% = (H+BB+HBP-R)/(H+BB+HBP-(1.4*HR)) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_losses | Pitchers, times individually receiving a losing decision | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_offspeeds_velocity_average | Pitchers, average offspeed pitch type pitch speed (among pitches with recorded velocity only) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_on_base_percentage_against | pitchers, OBP allowed to opponent hitters (opponent on-base pct) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_on_base_percentage_against_home | Pitchers, OBP allowed in home games | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_on_base_percentage_against_lhb | pitchers, OBP allowed to opponent left-handed hitters (batting left-handed at time of the plate appearance) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_on_base_percentage_against_rhb | pitchers, OBP allowed to opponent right-handed hitters (batting right-handed at time of the plate appearance) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_on_base_percentage_against_road | Pitchers, OBP allowed in road games | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_on_base_plus_slugging_against | pitchers, OPS allowed to opponent hitters (opponent OPS) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_on_base_plus_slugging_against_home | Pitchers, OPS allowed in home games | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_on_base_plus_slugging_against_lhb | Pitchers, OPS allowed to opponent left-handed hitters (batting left-handed at time of the plate appearance) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_on_base_plus_slugging_against_rhb | Pitchers, OPS allowed to opponent right-handed hitters (batting right-handed at time of the plate appearance) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_on_base_plus_slugging_against_road | Pitchers, OPS allowed in road games | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_opponent_at_bats_runners_in_scoring_position | Number of opponent At-Bats a pitcher records with Runners In Scoring Position | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_opponent_average_home | Pitchers, batting average allowed in home games | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_opponent_average_home_road_differential | Difference between pitcher's batting average against in home games vs. road games, negative indicates better on the road | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_opponent_average_lhb | Pitchers, batting average allowed to opposing left-handed hitters (batting left-handed at time of the plate appearance) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_opponent_average_rhb | Pitchers, batting average allowed to opposing right-handed hitters (batting right-handed at time of the plate appearance) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_opponent_average_road | Pitchers, batting average allowed in road games | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_opponent_balls_in_play | Pitchers, number of balls in play by opposing batters | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_opponent_balls_in_play_per_game | Pitchers, number of balls in play per game by opposing batters. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_opponent_batting_average | Pitchers, batting average allowed to opposing hitters (opponent batting average) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_opponent_batting_average_runners_in_scoring_position | Pitchers, batting average allowed (hits / at-bats) with Runners in Scoring position. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_opponent_fly_balls | Pitchers, number of balls in play that were fly balls by opposing batters. (Note: does not include line drives or pop-ups) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_opponent_fly_balls_percentage | Pitchers, percentage of balls in play that were fly balls by opposing batters. (Note: does not include line drives or pop-ups) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_opponent_fly_pop_balls | Pitchers, number of balls in play by opposing batters that were either Fly Balls or Pop-Ups. (Note: does not include Line Drives) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_opponent_fly_pop_balls_percentage | Pitchers, percentage of balls in play by opposing batters that were either Fly Balls or Pop-Ups. (Note: does not include Line Drives) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_opponent_ground_balls | Pitchers, number of balls in play that were ground balls by opposing batters | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_opponent_ground_balls_out_percentage | Percentage of opponent ground balls that are groundouts. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_opponent_ground_balls_percentage | Pitchers, percentage of balls in play that were ground balls by opposing batters | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_opponent_ground_outs | Pitchers, number of times opponent batters grounded out. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_opponent_hits_runners_in_scoring_position | Number of Hits a pitcher surrenders with Runners In Scoring Position | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_opponent_line_drives | Pitchers, number of balls in play that were line drives by opposing batters. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_opponent_line_drives_outs | Pitchers, number of times opponent batters lined out. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_opponent_line_drives_percentage | Pitchers, percentage of balls in play that were line drives by opposing batters. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_opponent_on_base_percentage_home_road_differential | Difference between pitcher's OBP in home games vs. road games, negative indicates better on the road | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_opponent_on_base_plus_slugging_home_road_differential | Difference between pitcher's OPS in home games vs. road games, negative indicates better on the road | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_opponent_pop_ups | Pitchers, number of balls in play that were popped up by opposing batters | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_opponent_pop_ups_outs | Pitchers, number of times opponent batters popped out. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_opponent_pop_ups_percentage | Pitchers, percentage of balls in play that were popped up by opposing batters | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_opponent_slugging_home_road_differential | Difference between pitcher's SLG in home games vs. road games, negative indicates better on the road | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_opponent_times_reached | Pitchers, number of times allowing opponent to reach base safely (H + BB + HBP) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_out_of_zone_contact_percentage_against | Pitchers, opposing batters percentage of swings where contact was made on pitches thrown out of the zone. (Zones as defined by MLB, 1->9 is in-zone, 11->14 is out-of-zone) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_out_of_zone_whiff_percentage | Pitchers, opposing batters whiff percentage on pitches thrown out of the zone. (Zones as defined by MLB, 1->9 is in-zone, 11->14 is out-of-zone) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_out_percentage | The percentage of batters gotten out by a pitcher. The inverse of Pitcher OBP against. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_outs | Pitchers, number of outs pitched | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_relief_hits_allowed | Pitchers, hits allowed by relief pitchers | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_relief_home_runs_allowed | Pitchers, number of home runs hit against relief pitchers | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_relief_losses | Pitchers, time individually receiving decision = Loss, in relief | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_relief_pitches_thrown | Pitcher, number of pitches thrown in relief | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_relief_strikeouts | Pitchers, number of times striking out opposing batter in relief | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_relief_walks | Pitchers, number of times walking the opposing batter in relief | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_relief_wins | Pitchers, times individually receiving decision = Win, in relief | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_reliefs_innings_pitched_per_game | Pitchers, innings pitched (as a reliever) per game. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_rointerference | Pitchers, opposing batters to reach base via defensive interference (usually catcher's interference, but some instances of interference by other players) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_runs_allowed_per_nine | Pitchers, total runs (both earned and unearned) allowed per 9 innings pitched | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_sacflies_against | Pitchers, number of times batters successfully hit a sacrifice fly against them | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_sachits_against | Pitchers, number of times batters successfully completed a sacrifice hit against them (NOTE: this is sacrifice bunts for 1955 and after. Prior to 1955, bunts and flies were not separated as distinct categories) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_saves | Pitchers, saves | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_saves_opportunities | Pitchers, number of save opportunities | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_saves_percentage | Pitchers, percentage of save opportunities that end in a save) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_screwballs_velocity_average | Pitchers, average screwball pitch speed (among pitches with recorded velocity only) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_shutouts | Pitchers, number of times an individual pitcher threw a complete-game shutout | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_singles_against | Pitchers, number of singles hit against them | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_sinkers_velocity_average | Pitchers, average 2-seam fastball pitch speed (among pitches with recorded velocity only) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_sliders_velocity_average | Pitchers, average slider pitch speed (among pitches with recorded velocity only) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_slugging_against | Pitchers, slugging percentage allowed to opposing hitters (opponent SLG) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_slugging_against_home | Pitchers, SLG allowed in home games | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_slugging_against_lhb | pitchers, SLG allowed to opponent left-handed hitters (batting left-handed at time of the plate appearance) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_slugging_against_rhb | pitchers, SLG allowed to opponent right-handed hitters (batting right-handed at time of the plate appearance) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_slugging_against_road | Pitchers, SLG allowed in road games | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_spin_rate_average | Pitchers, average spin rate on pitches thrown (among pitches with recorded spin rate only) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_splitters_velocity_average | Pitchers, average splitter pitch speed (among pitches with recorded velocity only) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_starting_hits_allowed | Pitchers, hits allowed as starting pitcher | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_starting_home_runs_allowed | Pitchers, number of home runs hit against starting pitchers | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_starting_innings_pitched_per_start | Pitchers, innings pitched (in starts only) per start | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_starting_losses | Pitchers, times individually receiving decision = Loss, as starter | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_starting_no_decisions | Starting Pitchers, times receiving no decision | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_starting_pitches_thrown | Pitcher, number of pitches thrown as starter | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_starting_strikeouts | Pitchers, number of times striking out batters as starter | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_starting_walks | Pitcher, number of times walking the opposing batter as starting pitcher | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_starting_wins | Pitchers, times individually receiving decision = 0, as starter | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_steals | Pitchers, stolen bases allowed while pitching | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_steals_attempts_against | Pitchers, number of stolen base attempts against (opponent's Steals + Caught Stealing) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_steals_home | Amount of times baserunners have successfully stolen home plate against a pitcher. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_steals_second | Amount of times baserunners have successfully stolen 2nd base against a pitcher. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_steals_third | Amount of times baserunners have successfully stolen 3rd base against a pitcher. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_strikeouts | Pitchers, number of times striking out opposing batters | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_strikeouts_enforced_strikes | Number of pitcher strikeouts that ended on an enforced strike. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_strikeouts_looking | Pitchers, number of strikeouts when looking (not swinging) at the strikeout pitch. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_strikeouts_looking_percentage | Pitchers, percent of strikeouts when looking (not swinging) at the strikeout pitch. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_strikeouts_per_game | Pitchers, number of times striking out opposing batters per game played | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_strikeouts_per_nine | Pitchers, Strikeouts per 9 innings pitched | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_strikeouts_percentage | Pitchers - Percentage of batters faced that end in a strikeout. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_strikeouts_swinging | Pitchers, number of strikeouts when swinging at the strikeout pitch. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_strikeouts_swinging_percentage | Pitchers, percent of strikeouts when swinging at the strikeout pitch. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_strikeouts_to_walk_ratio | Pitchers, ratio of strikeouts to walks | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_strikeouts_to_walk_ratio_number | Pitchers, ratio of strikeouts to walks. Does not include raw totals of strikeouts and walks. IF A PITCHER HAS ZERO WALKS, THE RATIO WILL BE 0.00 | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_strikeouts_walks_differential | The difference in a pitchers strikeout percentage (batters faced that end in strikeouts) to their walk percentage (batters faced that end in walks). | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_strikeouts_walks_percentage | Pitchers - Percentage of batters faced that end in a strikeout minus percentage of batters faced that end in a walk. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_swings | Pitchers, number of times batter swung at the pitch | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_swings_against_percentage | Pitchers, percentage of pitches thrown where batter swung at the pitch. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_swings_contact_against | Pitchers, number of times batter swings at the pitch and makes contact | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_times_chased | Pitchers, number of times opposing batter swings at a pitch out of the zone | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_total_bases | Pitchers, total bases allowed to batters on their hits | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_triples | Pitchers, number of triples hit against them | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_unearned_runs | Pitchers, runs charged to the pitcher that were Unearned Runs | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_velocity_average | Pitchers, average pitch speed (among pitches with recorded velocity only) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_walks | Pitchers, number of times walking the opposing batter | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_walks_enforced | Number of pitcher walks that ended on an enforced ball. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_walks_per_game | Pitchers, number of times walking the opposing batter per game played | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_walks_per_nine | Pitchers, Walks per 9 Innings Pitched | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_walks_percentage | Pitchers - Percentage of batters faced that end in a walk. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_whiffs | Pitchers, number of times batter swings at the pitch without making contact | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_whiffs_percentage | Pitchers, percentage of swings by opposing batters where contact was not made (swing+miss / swings) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcher_wins | Pitchers, times individually receiving decision = Win | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitcherbatting_average_balls_in_play_against | Pitchers, (H - HR) / (AB - K - HR + SF) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitches_per_inning | The number of pitches thrown per inning pitched. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitches_thrown | Pitcher, number of pitches thrown | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitches_thrown_balls | Pitcher, number of pitches thrown that were balls (NOTE: includes balls issued on automatic intentional walks) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitches_thrown_balls_percentage | Pitcher, percentage of pitches thrown that were balls | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitches_thrown_in_zone_percentage | Pitchers, percentage of pitches thrown that were thrown in the zone. (Zones as defined by MLB, 1->9 is in-zone, 11->14 is out-of-zone) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitches_thrown_out_zone_percentage | Pitchers, percentage of pitches thrown that were thrown out of the zone. (Zones as defined by MLB, 1->9 is in-zone, 11->14 is out-of-zone) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitches_thrown_per_batter | Pitchers, number of pitches thrown per batter faced | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitches_thrown_per_game | Pitcher, number of pitches thrown per game | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitches_thrown_starter_per_start | Pitcher, number of pitches thrown per game started | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitches_thrown_strikes | Pitcher, number of pitches thrown that were strikes | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitches_thrown_strikes_called | Pitcher, number of pitches thrown that were called strikes (strikes looking, batter did not swing) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitches_thrown_strikes_called_percentage | Pitching, percent of pitches thrown that were called strikes (strikes looking, batter did not swing) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitches_thrown_strikes_called_strikes_percentage | Pitching, percent of strikes thrown that were called strikes (strikes looking, batter did not swing) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitches_thrown_strikes_percentage | Pitcher, percentage of pitches thrown that were strikes | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitches_thrown_strikes_swinging | Pitcher, number of pitches thrown that were swinging strikes | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| pitching_winning_percentage | Pitchers, win percentage (individual pitcher decisions only) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| put_away_rate | Pitchers, the rate of two-strike pitches that result in a strikeout. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| quality_starts | Starting Pitchers, starts with 6.0+ IP and 3 or fewer ER allowed | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| quality_starts_percentage | Percentage of a Starting Pitcher's Starts that were Quality Starts (starts with 6.0+ IP and 3 or fewer ER allowed) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| rbi_opponent | Pitching, runs batted in by the opposing team | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| runs_allowed | Pitchers, runs scored that were charged to the pitcher | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| runs_allowed_per_game | Runs Allowed per Game | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| runs_allowed_per_game_differential | Difference between runs allowed per game in home games vs away games, negative indicates a higher value on the road | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| runs_allowed_per_home | Runs scored by the opponent per home game | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| runs_allowed_per_road | Runs scored by the opponent per away game | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| runs_allowed_r_p | Pitchers, runs scored that were charged to relief pitcher | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| runs_allowed_s_p | Pitchers, runs scored that were charged to starting pitcher | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| team_earned_runs_allowed | Teams, Earned Runs charged to the team (NOTE: may differ from individual Pitcher ER due to Section 10.18 (i) of the Scoring rules) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| team_era | Teams, Earned Runs allowed per 9 IP | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| team_era_home_road_diffential | Teams, Difference between team ERA at home vs. away | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| team_holds | Team total for holds by relief pitchers | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| team_runs_allowed_for_home_runs_percentage | Teams, percentage of runs allowed to opponents which came via RBI from allowing a home run. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| team_saves | Teams, saves | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| team_saves_opportunities | Teams, number of save opportunities | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| team_saves_percentage | Pitchers, percentage of save opportunities that end in a save | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| team_shutouts | Teams, number of games allowing zero runs to opponent | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| teams_runs_allowed_for_home_runs | Teams, total number of runs allowed to opponents which came via RBI from allowing a home run. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| war | Season-level statistic measuring a players wins above replacement as a player (hitting+pitching) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| war_pitching | Season-level statistic measuring a players wins above replacement as a pitcher. | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| whip | Pitchers, Walks + Hits allowed per Inning Pitched | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| whip_three_decimals | Pitchers, Walks + Hits allowed per Inning Pitched (formatted to include the thousandth place) | MLB | pitcher | CATEGORY_TO_MEASURE_PITCHER |
| player_record |  | NHL | additional | Additional Measures |
| team_record |  | NHL | additional | Additional Measures |
| assists |  | NHL | general | CATEGORY_TO_MEASURE |
| assists_per_game | The number of assists per game. | NHL | general | CATEGORY_TO_MEASURE |
| awarded_own_goal | Goals were a player/team is credited with a goal after their opponent puts it in their own net | NHL | general | CATEGORY_TO_MEASURE |
| blocked_shots | Blocked shots | NHL | general | CATEGORY_TO_MEASURE |
| blocked_shots_diff | Teams, difference between blocked shots and opponent blocked shots | NHL | general | CATEGORY_TO_MEASURE |
| blocked_shots_per_game | The number of shots blocked per game. | NHL | general | CATEGORY_TO_MEASURE |
| blocked_shots_plus_hits | The sum of blocks and hits | NHL | general | CATEGORY_TO_MEASURE |
| coach | Number of coaches | NHL | general | CATEGORY_TO_MEASURE |
| comb_goals_no_shootout | The total number of goals scored and allowed (NOT including the goals added for winning a shootout) MUST USE TEAM GROUPING when looking at a game level. | NHL | general | CATEGORY_TO_MEASURE |
| comb_goals_per_game | Combined goals for and goals allowed per game. MUST USE TEAM GROUPING when looking at a game level. | NHL | general | CATEGORY_TO_MEASURE |
| comb_goals_with_shootout | The total number of goals scored and allowed (including the goals added for winning a shootout) MUST USE TEAM GROUPING when looking at a game level. | NHL | general | CATEGORY_TO_MEASURE |
| comb_shots | Total number of shots on goal, own plus opponent. MUST USE TEAM GROUPING when looking at a game level. | NHL | general | CATEGORY_TO_MEASURE |
| count_of_players | Number of Players | NHL | general | CATEGORY_TO_MEASURE |
| count_of_seasons | Count of seasons | NHL | general | CATEGORY_TO_MEASURE |
| count_of_shifts | Number of shifts. | NHL | general | CATEGORY_TO_MEASURE |
| count_of_teams | The number of teams. | NHL | general | CATEGORY_TO_MEASURE |
| emptynet_goals |  | NHL | general | CATEGORY_TO_MEASURE |
| evenstrength_assist_pct | Percent of assists on even strength. | NHL | general | CATEGORY_TO_MEASURE |
| evenstrength_assists | Number of assists recorded while even strength | NHL | general | CATEGORY_TO_MEASURE |
| evenstrength_goal_diff | Even strength goals scored minus opponent even strength goals scored (team) | NHL | general | CATEGORY_TO_MEASURE |
| evenstrength_goal_pct | Percent of goals scored when team is at even strength. | NHL | general | CATEGORY_TO_MEASURE |
| evenstrength_goals | Goals scored while on even strength (5v5) (4v4) etc. | NHL | general | CATEGORY_TO_MEASURE |
| evenstrength_goals_per_game | Goals scored while on even strength (5v5) (4v4) etc. per game | NHL | general | CATEGORY_TO_MEASURE |
| evenstrength_point_pct | Percent of a player's points scored on even strength. | NHL | general | CATEGORY_TO_MEASURE |
| evenstrength_points | Points (goals + assists) scored while even strength | NHL | general | CATEGORY_TO_MEASURE |
| evenstrength_shots |  | NHL | general | CATEGORY_TO_MEASURE |
| faceoffs_lost |  | NHL | general | CATEGORY_TO_MEASURE |
| faceoffs_total |  | NHL | general | CATEGORY_TO_MEASURE |
| faceoffs_win_pct | The percentage of face offs won. | NHL | general | CATEGORY_TO_MEASURE |
| faceoffs_won |  | NHL | general | CATEGORY_TO_MEASURE |
| fav_losses | Games lost where the team was favored based on the pre-game moneyline odds | NHL | general | CATEGORY_TO_MEASURE |
| fav_wins | Games won where the team was favored based on the pre-game moneyline odds | NHL | general | CATEGORY_TO_MEASURE |
| favorite_winning_pct | Team Win Percentage in games where the team was favored based on the pre-game moneyline odds | NHL | general | CATEGORY_TO_MEASURE |
| first_period_goal_diff | 1st Period goals scored minus 1st period goals allowed | NHL | general | CATEGORY_TO_MEASURE |
| first_period_goal_pct | Percent of total goals scored in first period. | NHL | general | CATEGORY_TO_MEASURE |
| first_period_goals | Goals scored in the first period | NHL | general | CATEGORY_TO_MEASURE |
| first_period_shot_diff | Shots on goal differential in the first period | NHL | general | CATEGORY_TO_MEASURE |
| first_period_shots | Shots on goal in the first period | NHL | general | CATEGORY_TO_MEASURE |
| game_tying_goal | Number of game-tying goals | NHL | general | CATEGORY_TO_MEASURE |
| game_tying_goals_allowed |  | NHL | general | CATEGORY_TO_MEASURE |
| game_winning_goal_pct | Game winning goals divided by total goals scored. | NHL | general | CATEGORY_TO_MEASURE |
| game_winning_goals | The number of Game Winning goals | NHL | general | CATEGORY_TO_MEASURE |
| games_as_underdog | Games where the team was the underdog based on the pre-game moneyline odds | NHL | general | CATEGORY_TO_MEASURE |
| games_favored | Games where the team was favored based on the pre-game moneyline odds | NHL | general | CATEGORY_TO_MEASURE |
| games_over_total | Team Games where the total score was over the pre-game Over/Under point total | NHL | general | CATEGORY_TO_MEASURE |
| games_played | The number of games played. | NHL | general | CATEGORY_TO_MEASURE |
| games_pushed | Team Games where the total score was equal to the pre-game Over/Under point total. | NHL | general | CATEGORY_TO_MEASURE |
| games_under_total | Team Games where the total score was under the pre-game Over/Under point total | NHL | general | CATEGORY_TO_MEASURE |
| giveaways | Giveaways | NHL | general | CATEGORY_TO_MEASURE |
| giveaways_per_game | The number of Giveaways per game. | NHL | general | CATEGORY_TO_MEASURE |
| goal_diff | The differential between own and opponent goals in regulation and OT (does not include goals added for winning shootout) | NHL | general | CATEGORY_TO_MEASURE |
| goal_diff_per_game | Goals for minus goals against. | NHL | general | CATEGORY_TO_MEASURE |
| goal_diff_second_period | Goals scored minus goals allowed through the end of the 2nd period | NHL | general | CATEGORY_TO_MEASURE |
| goal_diff_third_period | Goals scored minus goals allowed through the end of the 3rd period | NHL | general | CATEGORY_TO_MEASURE |
| goals | The number of own goals scored in regulation and OT (does not include goal added for winning shootout) | NHL | general | CATEGORY_TO_MEASURE |
| goals_against | The number of opponent goals scored in regulation and OT (does not include goal added for winning shootout) | NHL | general | CATEGORY_TO_MEASURE |
| goals_against_per_game | The number of opponent goals scored per game in regulation and OT (does not include goal added for winning shootout) | NHL | general | CATEGORY_TO_MEASURE |
| goals_against_thru_second_period | Opponent goals in the game through the end of the 2nd period | NHL | general | CATEGORY_TO_MEASURE |
| goals_against_thru_third_period | Opponent goals in the game through the end of the 3rd period | NHL | general | CATEGORY_TO_MEASURE |
| goals_per_game | The number of own goals scored per game in regulation and OT (does not include goal added for winning shootout) | NHL | general | CATEGORY_TO_MEASURE |
| goals_thru_second_period | Goals scored in the game through the end of the 2nd period | NHL | general | CATEGORY_TO_MEASURE |
| goals_thru_third_period | Goals scored in the game through the end of the 3rd period. | NHL | general | CATEGORY_TO_MEASURE |
| hat_trick_games | Number of games with three or more goals scored by a player | NHL | general | CATEGORY_TO_MEASURE |
| hit_diff | Team Hit Differential | NHL | general | CATEGORY_TO_MEASURE |
| hit_diff_per_game | Team Hit Differential per Game | NHL | general | CATEGORY_TO_MEASURE |
| hit_net_pct | The percent of total shot attempts that are shots on goal. | NHL | general | CATEGORY_TO_MEASURE |
| hits |  | NHL | general | CATEGORY_TO_MEASURE |
| hits_per_game | The number of hits per game played. | NHL | general | CATEGORY_TO_MEASURE |
| hits_suffered | Number of hits suffered by a skater | NHL | general | CATEGORY_TO_MEASURE |
| home_losses | Losses at home | NHL | general | CATEGORY_TO_MEASURE |
| home_penalty_killing_pct | Success rate of preventing a goal while shorthanded at home. | NHL | general | CATEGORY_TO_MEASURE |
| home_wins | Wins in home games. | NHL | general | CATEGORY_TO_MEASURE |
| losses | The number of team losses (0 points) NOTE: does not include OTL (1 point). | NHL | general | CATEGORY_TO_MEASURE |
| missed_shots | Missed shots | NHL | general | CATEGORY_TO_MEASURE |
| missed_shots_per_game | The number of missed shots per game. | NHL | general | CATEGORY_TO_MEASURE |
| multi_assist_games | Games where a player recorded more than 1 assist | NHL | general | CATEGORY_TO_MEASURE |
| multi_goal_games | Games where a player scored more than 1 goal | NHL | general | CATEGORY_TO_MEASURE |
| multi_point_games | Games where a player scored more than 1 point | NHL | general | CATEGORY_TO_MEASURE |
| net_penalties | Penalties drawn minus penalties taken | NHL | general | CATEGORY_TO_MEASURE |
| net_takeaways | Takeaways minus giveaways | NHL | general | CATEGORY_TO_MEASURE |
| opp_blocked_shots | The number of shots the opponent had blocked. | NHL | general | CATEGORY_TO_MEASURE |
| opp_emptynet_goals | empty net goals allowed | NHL | general | CATEGORY_TO_MEASURE |
| opp_evenstrength_goals |  | NHL | general | CATEGORY_TO_MEASURE |
| opp_evenstrength_goals_per_game | Opponent goals scored while on even strength (5v5) (4v4) etc. per game | NHL | general | CATEGORY_TO_MEASURE |
| opp_evenstrength_shots |  | NHL | general | CATEGORY_TO_MEASURE |
| opp_first_period_goals | Opponent 1st period goals | NHL | general | CATEGORY_TO_MEASURE |
| opp_first_period_shots | Shots against in the first period | NHL | general | CATEGORY_TO_MEASURE |
| opp_hits | The number of opponent hits. | NHL | general | CATEGORY_TO_MEASURE |
| opp_hits_per_game | The number of opponent hits per game. | NHL | general | CATEGORY_TO_MEASURE |
| opp_missed_shots | The number of opponent missed shots. | NHL | general | CATEGORY_TO_MEASURE |
| opp_missed_shots_per_game | The number of opponent missed shots per game. | NHL | general | CATEGORY_TO_MEASURE |
| opp_penalties | The number of opponent penalties. | NHL | general | CATEGORY_TO_MEASURE |
| opp_penalties_per_game | The number of opponent penalties per game. | NHL | general | CATEGORY_TO_MEASURE |
| opp_penalty_minutes | The number of opponent penalty minutes. | NHL | general | CATEGORY_TO_MEASURE |
| opp_penalty_minutes_per_game | The number of opponent penalty minutes per game. | NHL | general | CATEGORY_TO_MEASURE |
| opp_powerplay_goals | The number of opponent power play goals. | NHL | general | CATEGORY_TO_MEASURE |
| opp_powerplay_goals_per_game | The number of opponent power play goals per game. | NHL | general | CATEGORY_TO_MEASURE |
| opp_powerplay_opportunities | The number of opponent power play opportunities. | NHL | general | CATEGORY_TO_MEASURE |
| opp_powerplay_opportunities_per_game | The number of opponent power play opportunities per game. | NHL | general | CATEGORY_TO_MEASURE |
| opp_powerplay_shots | The number of opponent power play shots on goal. | NHL | general | CATEGORY_TO_MEASURE |
| opp_powerplay_shots_per_game | The number of opponent power plays shots on goal per game. | NHL | general | CATEGORY_TO_MEASURE |
| opp_powerplay_shots_per_powerplay | Opposing team's power play shots on goal per power play opportunity. | NHL | general | CATEGORY_TO_MEASURE |
| opp_second_period_goals | Opponent 2nd period goals | NHL | general | CATEGORY_TO_MEASURE |
| opp_second_period_shots | Shots against in the second period | NHL | general | CATEGORY_TO_MEASURE |
| opp_shootout_attempts | The number of opponent shootout attempts. | NHL | general | CATEGORY_TO_MEASURE |
| opp_shootout_goals | The number of opponent shootout goals. | NHL | general | CATEGORY_TO_MEASURE |
| opp_shorthanded_goals | The number of shorthanded goals scored by the opponent. | NHL | general | CATEGORY_TO_MEASURE |
| opp_shorthanded_goals_per_game | The number of opponent shorthanded goals per game. | NHL | general | CATEGORY_TO_MEASURE |
| opp_shorthanded_shots | The number of opponent short handed shots on goal. | NHL | general | CATEGORY_TO_MEASURE |
| opp_shorthanded_shots_per_game | The number of shorthanded shots on goal per game by the opponent. | NHL | general | CATEGORY_TO_MEASURE |
| opp_shot_attempts | Total Opponent Shot Attempts (opp shots on goal + opp missed shots + blocks) | NHL | general | CATEGORY_TO_MEASURE |
| opp_shot_attempts_per_game | Total Opponent Shot Attempts per game (opp shots on goal + opp missed shots + blocks) | NHL | general | CATEGORY_TO_MEASURE |
| opp_shots_per_game | Opponent Shots On Goal Per Game. | NHL | general | CATEGORY_TO_MEASURE |
| opp_third_period_goals | Opponent 3rd period goals | NHL | general | CATEGORY_TO_MEASURE |
| opp_third_period_shots | Shots against in the third period | NHL | general | CATEGORY_TO_MEASURE |
| over_pct | Percentage of Team Games where the total score was over the pre-game Over/Under goal total | NHL | general | CATEGORY_TO_MEASURE |
| overtime_goal_diff | Overtime goals scored minus overtime goals allowed (not including Shootouts) | NHL | general | CATEGORY_TO_MEASURE |
| overtime_goals_against | Opponent Overtime goals | NHL | general | CATEGORY_TO_MEASURE |
| overtime_goals_game |  | NHL | general | CATEGORY_TO_MEASURE |
| overtime_goals_period | Goals scored in overtime | NHL | general | CATEGORY_TO_MEASURE |
| overtime_losses | The number of team overtime losses (1 point). NOTE: includes Shootout Losses (1 point). | NHL | general | CATEGORY_TO_MEASURE |
| overtime_wins | Team wins in overtime - does not include regulation or shootout wins. | NHL | general | CATEGORY_TO_MEASURE |
| p_pts_diff_home_road | Difference between player points (goals plus assists) in home games vs. road games, negative indicates more on the road. | NHL | general | CATEGORY_TO_MEASURE |
| penalties | Penalties taken. | NHL | general | CATEGORY_TO_MEASURE |
| penalties_drawn | Number of penalties drawn by a player | NHL | general | CATEGORY_TO_MEASURE |
| penalties_per_game | Team Penalties Committed Per Game | NHL | general | CATEGORY_TO_MEASURE |
| penalties_served | Number of penalties served by a player | NHL | general | CATEGORY_TO_MEASURE |
| penalty_killing_pct | Success rate of preventing a goal while shorthanded | NHL | general | CATEGORY_TO_MEASURE |
| penalty_minutes |  | NHL | general | CATEGORY_TO_MEASURE |
| penalty_minutes_per_game | The number of penalty minutes per game. | NHL | general | CATEGORY_TO_MEASURE |
| penalty_shot_attempts | Number of penalty shot attempts. | NHL | general | CATEGORY_TO_MEASURE |
| penalty_shot_goals | Number of penalty shot goals | NHL | general | CATEGORY_TO_MEASURE |
| penalty_shot_missed | Number of penalty shots that missed the net completely | NHL | general | CATEGORY_TO_MEASURE |
| penalty_shot_pct | Percentage of scoring on penalty shot attempts (penalty shot goals / penalty shot attempts). | NHL | general | CATEGORY_TO_MEASURE |
| penalty_shots_on_goal | The number of skater penalty shots that were on goal (excludes shots that missed the net) | NHL | general | CATEGORY_TO_MEASURE |
| player_assist_pct_team_goals | Player, percent of team goals where a player assisted on the goal. MUST USE TEAM GROUPING | NHL | general | CATEGORY_TO_MEASURE |
| player_pct_team_assists | Players, percentage of team's assists. MUST USE TEAM GROUPING | NHL | general | CATEGORY_TO_MEASURE |
| player_pct_team_evenstrength_assists | Players, percentage of team's even strength assists. MUST USE TEAM GROUPING | NHL | general | CATEGORY_TO_MEASURE |
| player_pct_team_evenstrength_goals | Players, percent of team even strength goals (does not include goals awarded for Shootout Wins in calculation) where a player scored the goal. MUST USE TEAM GROUPING | NHL | general | CATEGORY_TO_MEASURE |
| player_pct_team_evenstrength_points | Players, percentage of Team's Goals scored (does not include goals awarded for Shootout Wins in calculation) + team's assists when a team is even strength. MUST USE TEAM GROUPING | NHL | general | CATEGORY_TO_MEASURE |
| player_pct_team_face_offs | Players, percentage of Team's Face Offs Taken | NHL | general | CATEGORY_TO_MEASURE |
| player_pct_team_goals | Players, percentage of Team's Goals scored (does not include goals awarded for Shootout Wins in calculation) MUST USE TEAM GROUPING. | NHL | general | CATEGORY_TO_MEASURE |
| player_pct_team_points | Players, percentage of Team's Goals scored (does not include goals awarded for Shootout Wins in calculation) + team's assists. MUST USE TEAM GROUPING | NHL | general | CATEGORY_TO_MEASURE |
| player_pct_team_powerplay_assists | Players, percentage of team's power play assists. MUST USE TEAM GROUPING | NHL | general | CATEGORY_TO_MEASURE |
| player_pct_team_powerplay_goals | Player, percent of team power play goals (does not include goals awarded for Shootout Wins in calculation) where a player scored the goal. MUST USE TEAM GROUPING | NHL | general | CATEGORY_TO_MEASURE |
| player_pct_team_powerplay_points | Players, percentage of Team's Goals scored (does not include goals awarded for Shootout Wins in calculation) + team's assists on the power play. MUST USE TEAM GROUPING | NHL | general | CATEGORY_TO_MEASURE |
| player_pct_team_shots | Players, percentage of Team's Shots on Goal | NHL | general | CATEGORY_TO_MEASURE |
| player_point_assist_pct | The percent of player points that are assists. | NHL | general | CATEGORY_TO_MEASURE |
| player_point_goal_pct | The percent of player points that are goals. | NHL | general | CATEGORY_TO_MEASURE |
| player_point_pct_team_goals | Player, percent of team goals where a player either scored or assisted on the goal. MUST USE TEAM GROUPING | NHL | general | CATEGORY_TO_MEASURE |
| player_points | Goals plus assists equals player points. | NHL | general | CATEGORY_TO_MEASURE |
| player_points_per_game | Points per games played. | NHL | general | CATEGORY_TO_MEASURE |
| player_powerplay_points | The number of power play points for a player. | NHL | general | CATEGORY_TO_MEASURE |
| playoff_season | Seasons in which the team made the playoffs. | NHL | general | CATEGORY_TO_MEASURE |
| plus_minus |  | NHL | general | CATEGORY_TO_MEASURE |
| point_pct_powerplay | Percent of player points that come on the power play. | NHL | general | CATEGORY_TO_MEASURE |
| powerplay_assist_pct | Percent of assists on the power play | NHL | general | CATEGORY_TO_MEASURE |
| powerplay_assists |  | NHL | general | CATEGORY_TO_MEASURE |
| powerplay_assists_per_game | The number of power play assists per game. | NHL | general | CATEGORY_TO_MEASURE |
| powerplay_goal_diff | Power play goals scored minus opponent power play goals scored (team). | NHL | general | CATEGORY_TO_MEASURE |
| powerplay_goal_pct | Percent of goals scored when team is on power play. | NHL | general | CATEGORY_TO_MEASURE |
| powerplay_goals |  | NHL | general | CATEGORY_TO_MEASURE |
| powerplay_goals_per_game | The number of power play goals per game. | NHL | general | CATEGORY_TO_MEASURE |
| powerplay_opportunities |  | NHL | general | CATEGORY_TO_MEASURE |
| powerplay_opportunities_per_game | The number of power play opportunities per game. | NHL | general | CATEGORY_TO_MEASURE |
| powerplay_pct | Power play goals / power play opportunities. | NHL | general | CATEGORY_TO_MEASURE |
| powerplay_pct_at_home | Power play goals / power play opportunities for games played at Home | NHL | general | CATEGORY_TO_MEASURE |
| powerplay_pct_on_road | Power play goals / power play opportunities for games played on the Road | NHL | general | CATEGORY_TO_MEASURE |
| powerplay_pct_plus_penalty_kill_pct | The sum of a teams' power play percentage and their penalty-kill percentage. | NHL | general | CATEGORY_TO_MEASURE |
| powerplay_points_per_game | The number of power play points per game. | NHL | general | CATEGORY_TO_MEASURE |
| powerplay_shooting_pct | Shooting percentage on the power play (team or player) | NHL | general | CATEGORY_TO_MEASURE |
| powerplay_shots |  | NHL | general | CATEGORY_TO_MEASURE |
| powerplay_shots_per_game | The number of power play shots per game. | NHL | general | CATEGORY_TO_MEASURE |
| powerplay_shots_per_powerplay | The number of power play shots on goal per power play opportunity. | NHL | general | CATEGORY_TO_MEASURE |
| road_losses | Losses in road games. | NHL | general | CATEGORY_TO_MEASURE |
| road_penalty_killing_pct | Success rate of preventing a goal while shorthanded on the road. | NHL | general | CATEGORY_TO_MEASURE |
| road_wins | Wins in road games. | NHL | general | CATEGORY_TO_MEASURE |
| saves_per_game | Saves per game played. | NHL | general | CATEGORY_TO_MEASURE |
| second_period_goal_diff | 2nd Period goals scored minus 2nd period goals allowed | NHL | general | CATEGORY_TO_MEASURE |
| second_period_goal_pct | Percent of total goals scored in second period. | NHL | general | CATEGORY_TO_MEASURE |
| second_period_goals | Goals scored in second period | NHL | general | CATEGORY_TO_MEASURE |
| second_period_shot_diff | Shots on goal differential in the second period | NHL | general | CATEGORY_TO_MEASURE |
| second_period_shots | Shots on goal in the second period | NHL | general | CATEGORY_TO_MEASURE |
| shifts_per_games_played | The number of shifts per games played. | NHL | general | CATEGORY_TO_MEASURE |
| shooting_pct | The number percentage of goals to shots on goal. | NHL | general | CATEGORY_TO_MEASURE |
| shootout_attempts | Shootout attempts | NHL | general | CATEGORY_TO_MEASURE |
| shootout_goals | Shootout goals | NHL | general | CATEGORY_TO_MEASURE |
| shootout_pct | Percentage of scoring on shootout attempts (shootout goals / shootout attempts). | NHL | general | CATEGORY_TO_MEASURE |
| shorthanded_assist_pct | Percent of assists while short-handed. | NHL | general | CATEGORY_TO_MEASURE |
| shorthanded_assists | The assists that occur when the opposing team has an advantage in players. | NHL | general | CATEGORY_TO_MEASURE |
| shorthanded_goal_diff | Short handed goals scored minus opponent short handed goals scored (team). | NHL | general | CATEGORY_TO_MEASURE |
| shorthanded_goal_pct | Percentage of goals scored while shorthanded (Short handed goals divided by all goals). | NHL | general | CATEGORY_TO_MEASURE |
| shorthanded_goals | Number of goals scored while the opposing team has a player advantage. | NHL | general | CATEGORY_TO_MEASURE |
| shorthanded_goals_per_game | Number of goals scored while the opposing team has a player advantage per game. | NHL | general | CATEGORY_TO_MEASURE |
| shorthanded_player_points | The number of shorthanded player points. | NHL | general | CATEGORY_TO_MEASURE |
| shorthanded_shots | Shorthanded shots on goal. | NHL | general | CATEGORY_TO_MEASURE |
| shot_attempt_diff | The difference between Total Shot Attempts (shots on goal + missed shots + attempts blocked by opponent) and Opponent Total Shot Attempts. | NHL | general | CATEGORY_TO_MEASURE |
| shot_attempts | Total Shot Attempts (shots on goal + missed shots + attempts blocked by opponent). | NHL | general | CATEGORY_TO_MEASURE |
| shot_attempts_blocked | Shot Attempts Blocked by Opponent Skater | NHL | general | CATEGORY_TO_MEASURE |
| shot_attempts_blocked_per_game | Number of Shot Attempts by a Skater that are blocked by the opponent, per game | NHL | general | CATEGORY_TO_MEASURE |
| shot_attempts_per_game | Shot Attempts per Game (shots on goal + missed shots + attempts blocked by opponent). | NHL | general | CATEGORY_TO_MEASURE |
| shots_against | The number of opponent shots on goal (includes empty net goals against). | NHL | general | CATEGORY_TO_MEASURE |
| shots_against_per_game | The number of opponent shots on goal per game (includes empty net goals against). | NHL | general | CATEGORY_TO_MEASURE |
| shots_diff_per_game | The average difference in shots on goal per game. | NHL | general | CATEGORY_TO_MEASURE |
| shots_on_goal | Shots on goal (shots that are either saved by the opposing goalie or result in a goal scored). | NHL | general | CATEGORY_TO_MEASURE |
| shots_on_goal_diff | Shots on Goal - Opponent Shots on Goal. | NHL | general | CATEGORY_TO_MEASURE |
| shots_per_game | The number of shots on goal per game. | NHL | general | CATEGORY_TO_MEASURE |
| skater_evenstrength_time_on_ice | Time on ice for skaters while team is even strength. | NHL | general | CATEGORY_TO_MEASURE |
| skater_evenstrength_time_on_ice_per_game | Time on ice per game for skaters while team is even strength. | NHL | general | CATEGORY_TO_MEASURE |
| skater_minutes | The total number of minutes on ice. | NHL | general | CATEGORY_TO_MEASURE |
| skater_powerplay_time_on_ice | Time on ice for skaters while team is on power play. | NHL | general | CATEGORY_TO_MEASURE |
| skater_powerplay_time_on_ice_per_game | Time on Ice per game for skaters while team is on power play. | NHL | general | CATEGORY_TO_MEASURE |
| skater_shorthanded_time_on_ice | Time on ice for skaters while team is shorthanded. | NHL | general | CATEGORY_TO_MEASURE |
| skater_shorthanded_time_on_ice_per_game | Time on Ice per game for skaters while team is shorthanded. | NHL | general | CATEGORY_TO_MEASURE |
| skater_time_on_ice | MINUTES:SECONDS | NHL | general | CATEGORY_TO_MEASURE |
| skater_time_on_ice_per_game | The amount of time on ice per game. | NHL | general | CATEGORY_TO_MEASURE |
| special_teams_goal_diff | Special team goals scored minus opponent special team goals scored (team). | NHL | general | CATEGORY_TO_MEASURE |
| special_teams_goals | Goals scored while player's team is on the power play or player's team is shorthanded. | NHL | general | CATEGORY_TO_MEASURE |
| special_teams_goals_against | Opponent goals scored while player's team is on the power play or player's team is shorthanded. | NHL | general | CATEGORY_TO_MEASURE |
| takeaways |  | NHL | general | CATEGORY_TO_MEASURE |
| takeaways_per_game | The number of takeaways per game. | NHL | general | CATEGORY_TO_MEASURE |
| team_blown_lead_losses | Losses after leading at any point in game (includes Overtime Losses). | NHL | general | CATEGORY_TO_MEASURE |
| team_comeback_wins | Wins after trailing at any point in game. | NHL | general | CATEGORY_TO_MEASURE |
| team_evenstrength_time_on_ice | Time on ice for teams while even strength. | NHL | general | CATEGORY_TO_MEASURE |
| team_evenstrength_time_on_ice_per_game | The average time on ice per game for teams while even strength | NHL | general | CATEGORY_TO_MEASURE |
| team_goal_diff_including_shooutout | The differential between own and opponent score (including the goals added for winning a shootout). | NHL | general | CATEGORY_TO_MEASURE |
| team_goals_against_avg_including_empty_net | Team Goals Against Average (including empty net goals). | NHL | general | CATEGORY_TO_MEASURE |
| team_goals_against_including_shootout | The opponent score (including the goal added for winning a shootout). | NHL | general | CATEGORY_TO_MEASURE |
| team_goals_including_shootout | The team score (including the goal added for winning a shootout). | NHL | general | CATEGORY_TO_MEASURE |
| team_point_pct | The teams point percentage (NHL's version of win percentage). | NHL | general | CATEGORY_TO_MEASURE |
| team_point_pct_at_home | The teams point percentage (NHL's version of win percentage) at home. | NHL | general | CATEGORY_TO_MEASURE |
| team_point_pct_diff_home_road | The difference between a team's point percentage (NHL's version of win percentage) in home games vs. road games, negative indicates better on the road. | NHL | general | CATEGORY_TO_MEASURE |
| team_point_pct_on_road | The teams point percentage (NHL's version of win percentage) in road games | NHL | general | CATEGORY_TO_MEASURE |
| team_points_earned | The number of team points earned based on the result of the game (W=2, L=0, T=1, OTL=1). | NHL | general | CATEGORY_TO_MEASURE |
| team_points_earned_at_home | Team points in home games. (W=2, L=0, T=1, OTL=1). | NHL | general | CATEGORY_TO_MEASURE |
| team_points_earned_on_road | Team points in road games. (W=2, L=0, T=1, OTL=1). | NHL | general | CATEGORY_TO_MEASURE |
| team_powerplay_time_on_ice | Time on ice for teams while on the power play. | NHL | general | CATEGORY_TO_MEASURE |
| team_powerplay_time_on_ice_per_game | The average time on ice per game for teams while on the power play. | NHL | general | CATEGORY_TO_MEASURE |
| team_regulation_and_overtime_wins | Team wins in regulation and overtime - does not include shootout wins. | NHL | general | CATEGORY_TO_MEASURE |
| team_regulation_wins | Team wins in regulation - does not include overtime or shootout wins. | NHL | general | CATEGORY_TO_MEASURE |
| team_save_pct_incl_empty_net_goals | The percentage of shots on goal stopped by a team (including empty net goals). | NHL | general | CATEGORY_TO_MEASURE |
| team_shootout_losses | The number of team shootout losses | NHL | general | CATEGORY_TO_MEASURE |
| team_shootout_wins | The number of team shootout wins | NHL | general | CATEGORY_TO_MEASURE |
| team_shorthanded_time_on_ice | Time on ice for teams while shorthanded | NHL | general | CATEGORY_TO_MEASURE |
| team_shorthanded_time_on_ice_per_game | The average time on ice per game for teams while shorthanded | NHL | general | CATEGORY_TO_MEASURE |
| team_ties | The number of team ties (1 point). | NHL | general | CATEGORY_TO_MEASURE |
| team_time_on_ice | The total time the team was on the ice (including OT). | NHL | general | CATEGORY_TO_MEASURE |
| team_time_on_ice_per_game | The average time on ice for teams per game | NHL | general | CATEGORY_TO_MEASURE |
| team_win_pct | Total wins (Regulation/Overtime/Shootout) plus Team ties divided by 2, divided by total games played | NHL | general | CATEGORY_TO_MEASURE |
| team_wins | The number of team wins (2 points). | NHL | general | CATEGORY_TO_MEASURE |
| third_period_goal_diff | 3rd Period goals scored minus 3rd period goals allowed | NHL | general | CATEGORY_TO_MEASURE |
| third_period_goal_pct | Percent of total goals scored in third period. | NHL | general | CATEGORY_TO_MEASURE |
| third_period_goals | Goals scored in the third period | NHL | general | CATEGORY_TO_MEASURE |
| third_period_shot_diff | Shots on goal differential in the third period | NHL | general | CATEGORY_TO_MEASURE |
| third_period_shots | Shots on goal in the third period | NHL | general | CATEGORY_TO_MEASURE |
| under_pct | Percentage of Team Games where the total score was under the pre-game Over/Under goal total | NHL | general | CATEGORY_TO_MEASURE |
| underdog_losses | Games lost where the team was the underdog based on the pre-game moneyline odds | NHL | general | CATEGORY_TO_MEASURE |
| underdog_win_pct | Team Win Percentage in games where the team was the underdog based on the pre-game moneyline odds | NHL | general | CATEGORY_TO_MEASURE |
| underdog_wins | Games won where the team was the underdog based on the pre-game moneyline odds | NHL | general | CATEGORY_TO_MEASURE |
| goalie_avg_points | The average amount of possible points earned per game, per goalie appearance | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_evenstrength_save_pct | Save percentage while teams are even strength | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_evenstrength_saves | The number of saves by goalies while teams are at even strength. | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_games_played | Games played by goalies | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_games_started | Games started in goal | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_goals_against | Goalie Goals Against | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_goals_against_avg | Goalie Goals Against Average | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_opp_evenstrength_shots | The number of shots on goal against goalies while teams are even strength. | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_overtime_losses_no_shootout | Goalie losses in games that went to Overtime but *not* to a shootout | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_penalty_goals_against | Number of penalty shot goals against | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_penalty_shot_attempts_against | The number of penalty shot attempts a goalie faced (includes shots that missed the goal completely) | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_penalty_shot_save_pct | Percentage of penalty shot attempts that do not result in a goal. | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_penalty_shot_saves | Number of penalty shot saves | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_powerplay_goals_against | The number of goals scored against goalies while the opposing team is on the power play (goalie’s team is shorthanded) | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_powerplay_save_pct | Save percentage while the opposing team is on the power play (goalie’s team is shorthanded) | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_powerplay_saves | The number of saves by goalies while the opposing team is on the power play (goalie’s team is shorthanded) | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_powerplay_shots_against | The number of shots on goal against goalies while the opposing team is on the power play (goalie’s team is shorthanded) | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_save_pct | The percentage of shots on goal that a goalie stops | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_saves | Goalie Saves | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_shootout_attempts_against | Goalkeeper shootout attempts against | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_shootout_goals_against | Goalkeeper shootout goals against | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_shootout_losses | Goalie losses in games that go to a shootout | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_shootout_save_pct | Percentage of shootout attempts that do not result in a goal. | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_shootout_wins | Goalie wins in games that go to a shootout | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_shorthanded_goals_against | The number of goals scored against goalies while the opposing team is shorthanded (goalie’s team is on power play) *excludes empty-net goals* | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_shorthanded_save_pct | Save percentage while the opposing team is shorthanded (goalie’s team is on power play) | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_shorthanded_saves | The number of saves by goalies while the opposing team is shorthanded (goalie’s team is on power play) | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_shorthanded_shots_against | The number of shots on goal against goalies while the opposing team is shorthanded (goalie’s team is on power play) | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_shots_against | The total number of shots on goal against Goalie(s). | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_shots_against_per_game | The number of shots against a goalie(s). * Does not include empty net goals, that is Team SA/G | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_special_teams_goals_against | Goals against while the opposing team is on the power play (goalie’s team is shorthanded) or opposing team is shorthanded (goalie’s team is on power play) | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_special_teams_save_pct | Save percentage while the opposing team is on the power play (goalie’s team is shorthanded) or opposing team is shorthanded (goalie’s team is on power play) | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_special_teams_saves | Saves while the opposing team is on the power play (goalie’s team is shorthanded) or opposing team is shorthanded (goalie’s team is on power play) | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_special_teams_shot_attempts_against | Shots against while the opposing team is on the power play (goalie’s team is shorthanded) or opposing team is shorthanded (goalie’s team is on power play) | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_time_on_ice | Total time on ice for goalies | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_time_on_ice_per_game | Time on ice per game for goalies | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_total_decisions | The number of decisions (W/L/T/OTL) attributed to that goalie. This is used to determine Goalie Win Percentage. | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_total_losses | Goalie received a Loss that game. (L: 0 team points) | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_total_overtime_losses | *Pre-2005, Teams were not awarded a point for an OTL | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_total_shutouts | Goaltender total shutouts | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_total_ties | Goaltender total ties | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_total_wins | Goaltender total wins | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |
| goalie_win_pct | The goalie's point percentage in games that they received a decision. | NHL | goalie | CATEGORY_TO_MEASURE_GOALIE |

## SmartStat Relevance

- Provides stable measure vocabulary for semantic dictionaries.
- Improves deterministic candidate-resolution context for stat terms.
- Supplies planner grammar measure slots without runtime coupling.
