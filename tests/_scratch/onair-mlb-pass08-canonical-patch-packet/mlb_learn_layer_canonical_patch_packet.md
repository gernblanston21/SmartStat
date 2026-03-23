# MLB Learn-Layer Canonical Patch Packet (PASS_08)

## A. Patch-ready additions

- Total patch-ready lines: **8**
- Target file (future controlled mutation pass): `SmartStat_Mappings.learn.ini`

### Section: `CATEGORY_TO_MEASURE`

- `ASSISTS=fielding_assists`
  - dependency: `none`
  - confidence: `medium`
  - source: `measure_214`
  - rationale: Fielding-prefixed measure with supporting user phrase; category canonical destination is viable.
- `DOUBLE PLAYS=fielding_double_plays`
  - dependency: `none`
  - confidence: `medium`
  - source: `measure_217`
  - rationale: Fielding-prefixed measure with supporting user phrase; category canonical destination is viable.
- `GAME WINNING RBI=game_winning_rbi`
  - dependency: `none`
  - confidence: `high`
  - source: `measure_762`
  - rationale: Direct MLB batting outcome concept; canonical category destination is clear.
- `GO AHEAD RBI=go_ahead_rbi`
  - dependency: `none`
  - confidence: `high`
  - source: `measure_769`
  - rationale: Direct MLB batting outcome concept; canonical category destination is clear.
- `PUTOUTS=fielding_putouts`
  - dependency: `none`
  - confidence: `medium`
  - source: `measure_226`
  - rationale: Fielding-prefixed measure with supporting user phrase; category canonical destination is viable.
- `SINGLES=singles`
  - dependency: `none`
  - confidence: `high`
  - source: `measure_62`
  - rationale: Direct MLB measure concept; aligns with CATEGORY_TO_MEASURE key->measure_value direction.

### Section: `CATEGORY_TO_MEASURE_PITCHER`

- `BLOWN SAVES=pitcher_blown_saves`
  - dependency: `none`
  - confidence: `high`
  - source: `measure_449`
  - rationale: Pitcher-prefixed measure clearly belongs in pitcher canonical section.

### Section: `CATEGORY_TO_MEASURE_ALIASES`

- `GO-AHEAD RBI=GO AHEAD RBI`
  - dependency: `after CATEGORY_TO_MEASURE:GO AHEAD RBI`
  - confidence: `high`
  - source: `user_107`
  - rationale: User phrase is a variant that can safely map to proposed canonical key.

## B. Blocked / excluded candidates

- manual_review_required: **8**
- reject: **11**
- composite_defer (PASS_07 source pool): **0**
- composite_defer (upstream PASS_05 context, intentionally out of patch body): **238**
- dependency-ambiguous dropped from patch-ready set: **0**

Compact excluded list (PASS_07 source pool):
- `measure_140` -> `manual_review_required`: Concept role ambiguity (runner vs pitcher context) prevents deterministic canonical placement.
- `measure_216` -> `manual_review_required`: Concept overlaps with caught_stealing and pitcher_caught_stealing; canonical destination requires manual decision.
- `measure_437` -> `manual_review_required`: Measure value lacks pitcher prefix while likely pitcher-domain; canonical destination needs manual confirmation.
- `measure_451` -> `manual_review_required`: Concept overlaps with fielding/non-pitcher caught stealing; canonical destination requires manual decision.
- `measure_49` -> `manual_review_required`: Concept collides across runner/fielding/pitcher semantics; canonical destination is not singular.
- `user_123` -> `manual_review_required`: User phrase maps to multiple candidate measure concepts; canonical dependency is ambiguous.
- `user_94` -> `manual_review_required`: Dependent measure concept is manual-review gated; alias cannot be safely promoted.
- `user_95` -> `manual_review_required`: Dependent measure concept is manual-review gated; alias cannot be safely promoted.
- `measure_156` -> `reject`: Alias key already exists in alias sections; no new semantic addition required.
- `measure_277` -> `reject`: Canonical key already exists; alias self-reference would be structurally invalid.
- `user_119` -> `reject`: User phrase equals proposed canonical key; alias entry would be redundant.
- `user_138` -> `reject`: User phrase equals proposed canonical key; alias entry would be redundant.
- `user_183` -> `reject`: Duplicate user phrase evidence; first occurrence retained for modeling.
- `user_2` -> `reject`: User phrase equals proposed canonical key; alias entry would be redundant.
- `user_4` -> `reject`: User phrase equals proposed canonical key; alias entry would be redundant.
- `user_40` -> `reject`: Duplicate user phrase evidence; first occurrence retained for modeling.
- `user_41` -> `reject`: Duplicate user phrase evidence; first occurrence retained for modeling.
- `user_6` -> `reject`: User phrase equals proposed canonical key; alias entry would be redundant.
- `user_69` -> `reject`: User phrase equals proposed canonical key; alias entry would be redundant.

## C. Ordering notes

- Canonical section entries must be applied before alias section entries in any future real mutation pass.
- In this packet, `CATEGORY_TO_MEASURE: GO AHEAD RBI=go_ahead_rbi` must precede `CATEGORY_TO_MEASURE_ALIASES: GO-AHEAD RBI=GO AHEAD RBI`.
- Aliases are valid only when canonical target key is existing or included earlier in the same approved patch set.
