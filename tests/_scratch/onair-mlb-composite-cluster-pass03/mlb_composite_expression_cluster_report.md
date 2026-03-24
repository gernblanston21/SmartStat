# MLB Composite Expression Cluster Report (PASS_03)

## Scope

- Pass: `MLB_ONAIR_COMPOSITE_EXPRESSION_CLUSTER_DEFINITION_PASS_03`
- Mode: read-only composite-cluster analysis.
- Inputs: PASS_02 lane-definition artifacts and PASS_01 classification artifacts.
- This pass did not modify runtime, INI files, or composite implementation logic.

## Composite Lane Recap

- Composite candidate count: **633**.
- PASS_01/PASS_02 already established this lane as non-runtime-authoritative and not safe for direct one-line INI promotion.
- Interpretation: composite value is substantial and structurally diverse, requiring multiple conceptual handling classes.

## Composite Cluster Model

- Defined cluster families:
- `ratio_or_slash_cluster`
- `text_plus_token_cluster`
- `multi_measure_sequence_cluster`
- `split_or_comparison_cluster`
- `header_value_pair_cluster`
- `contextual_template_cluster`
- `unresolved_composite_cluster`
- Boundary rule: every composite candidate is assigned to exactly one best-fit cluster, fail-closed to `unresolved_composite_cluster` when uncertain.

## Cluster Findings

### ratio_or_slash_cluster
- Count: **161** (25.43%)
- Structural signature: Contains slash-delimited or per-form ratio expressions (non-split context).
- Why not simple INI mapping: Requires ratio/computed composition semantics beyond direct one-key map.
- Capability needs: simple join logic, ordered multi-measure composition
- Representative examples:
  - at_bats_per_home_run
  - at_bats_per_strikeout
  - at_bats_per_rbi
  - walks_per_game
  - singles_per_game
### text_plus_token_cluster
- Count: **177** (27.96%)
- Structural signature: Contains literal text outside placeholder tokens.
- Why not simple INI mapping: Requires mixed literal-text and dynamic-token assembly not represented by simple mapping keys.
- Capability needs: ordered multi-measure composition, tooling/operator assist only
- Representative examples:
  - {{info.time.season}} ({{stats.player.season.games_played}} GAMES)
  - CAREER ({{info.player.experience | ordinal}} SEASON)
  - {{info.time.month}} ({{stats.player.season.month.at_bats}} AB)
  - 1st HALF ({{stats.player.season.all_star_break(before).games_played}} GAMES)
  - 2nd HALF ({{stats.player.season.all_star_break(after).games_played}} GAMES)
### multi_measure_sequence_cluster
- Count: **23** (3.63%)
- Structural signature: Ordered package of multiple measures (3+ components or explicit sequence delimiters).
- Why not simple INI mapping: Requires deterministic ordering/composition of multiple values, not a single lookup.
- Capability needs: ordered multi-measure composition, contextual template awareness
- Representative examples:
  - {{info.time.season}}: {{stats.player.season.batting_average}}, {{stats.player.season.home_runs}} {{custom.header_hr}}, {{stats.player.season.rbis}} {{custom.header_rbi}}
  - WITH {{info.team.name}}: {{stats.player.season.on_team.batting_average}}, {{stats.player.season.on_team.home_runs}} {{custom.header_hr}}, {{stats.player.season.on_team.rbis}} {{custom.header_rbi}}
  - [REGULAR SEASON: {{stats.player.season.batting_average}}, {{stats.player.season.home_runs}} {{custom.header_hr}}, {{stats.player.season.rbis}} {{custom.header_rbi}}
  - {{stats.player.runners(loaded).batting_average}} AVG, {{stats.player.runners(loaded).home_runs}} GRAND SLAMS
  - {{stats.player.season.count(3-2).batting_average}} AVG,  {{stats.player.season.count(3-2).home_runs}} HR,  {{stats.player.season.count(3-2).rbis}} RBI
### split_or_comparison_cluster
- Count: **115** (18.17%)
- Structural signature: Contains split/comparison indicators (vs, home/away, division/conference/time splits).
- Why not simple INI mapping: Requires qualifier/split context binding and comparative interpretation.
- Capability needs: future semantic resolver grammar, contextual template awareness
- Representative examples:
  - batting_average_home_road_differential
  - hits_home_road_differential
  - home_runs_differential_home_road
  - on_base_percentage_home_road_differential
  - on_base_plus_slugging_home_road_differential
### header_value_pair_cluster
- Count: **9** (1.42%)
- Structural signature: Custom header placeholders coupled to adjacent value tokens/stat fields.
- Why not simple INI mapping: Requires coupling of display-header semantics with composed value payloads.
- Capability needs: contextual template awareness, tooling/operator assist only
- Representative examples:
  - {{custom.header_avg}} VS PITCH TYPE
  - WITH {{info.team.name}}: {{stats.player.season.on_team.games_started}} {{custom.header_gstart}}, {{stats.player.season.on_team.pitcher_record}}, {{stats.player.season.on_team.pitcher_era}} {{custom.header_era}}
  - WITH {{info.team.name}}: {{stats.player.season.on_team.games_pitched}} {{custom.header_gpitch}}, {{stats.player.season.on_team.pitcher_record}}, {{stats.player.season.on_team.pitcher_era}} ERA
  - {{custom.leaders_avg}} vs LHP
  - {{custom.leaders_avg}} vs RHP
### contextual_template_cluster
- Count: **10** (1.58%)
- Structural signature: Composite meaning depends on playbook/template context rather than expression text alone.
- Why not simple INI mapping: Needs surrounding template metadata and context to resolve intent safely.
- Capability needs: contextual template awareness, future semantic resolver grammar
- Representative examples:
  - {{info.team.alias}}
  - {{info.time.season}}: {{stats.player.season.runners(loaded).batting_summary_abb}}
  - {{stats.coach.career.team_record}} ({{stats.coach.career. team_win_percentage}})
  - {{previous.team.minimum(runs, 10).game_date|flex_short}} {{previous.team.minimum(runs, 10).game_vs}} {{previous.team.minimum(runs, 10).opp_name}} ({{previous.team.minimum(runs, 10).runs}})
  - {{previous.team_player.minimum(home_runs, 2).full_name}} {{previous.team_player.minimum(home_runs, 2).game_vs}} {{previous.team_player.minimum(home_runs, 2).opp_name}}
### unresolved_composite_cluster
- Count: **138** (21.8%)
- Structural signature: Composite candidate signals present but no safe deterministic best-fit family.
- Why not simple INI mapping: Insufficient structural confidence for safe direct categorization or runtime mapping.
- Capability needs: future semantic resolver grammar, tooling/operator assist only
- Representative examples:
  - called_strike
  - batter_home_run_to_fly_ball_percentage
  - batting_average_balls_in_play
  - batting_average_differential
  - batting_average_runners_in_scoring_position

## High-Value Composite Opportunities

- Highest-frequency clusters:
  - text_plus_token_cluster: 177
  - ratio_or_slash_cluster: 161
  - unresolved_composite_cluster: 138
  - split_or_comparison_cluster: 115
  - multi_measure_sequence_cluster: 23
  - contextual_template_cluster: 10
  - header_value_pair_cluster: 9
- Most operator-relevant clusters:
  - text_plus_token_cluster: 177
  - multi_measure_sequence_cluster: 23
  - split_or_comparison_cluster: 115
  - ratio_or_slash_cluster: 161
- Lowest-complexity/highest-value clusters:
  - ratio_or_slash_cluster (161): Most deterministic structural shape for bounded follow-on analysis.
  - multi_measure_sequence_cluster (23): Clear ordered package semantics, manageable bounded composition rules.
- Highest-complexity defer clusters:
  - contextual_template_cluster (10): High dependence on playbook context; higher resolver coupling risk.
  - unresolved_composite_cluster (138): No safe deterministic best-fit cluster; fail-closed by design.
  - header_value_pair_cluster (9): Display/header coupling semantics require careful template-context modeling.

## Real-World Broadcast Value

- Operators can surface richer composite stat callouts without manually stitching every expression under show pressure.
- Ratio and ordered stat packages can become more consistent and faster to prepare when their structure is modeled explicitly.
- Text+token and split/comparison composites can improve contextual storytelling without forcing unsafe direct runtime mappings.
- Keeping unresolved composites fail-closed protects on-air safety while preserving candidate value for controlled future review.

## Recommended Next Pass

- `MLB_ONAIR_COMPOSITE_RATIO_SEQUENCE_CONTRACT_PASS_04`
- Scope: read-only contract-definition pass for `ratio_or_slash_cluster` and `multi_measure_sequence_cluster` only.
- Boundary: no runtime mutation, no INI mutation, no composite/render implementation.
