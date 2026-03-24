# MLB Composite Ratio/Sequence Contract Report (PASS_04)

## Scope

- Pass: `MLB_ONAIR_COMPOSITE_RATIO_SEQUENCE_CONTRACT_PASS_04`
- Mode: read-only contract-definition analysis.
- In-scope clusters only: `ratio_or_slash_cluster`, `multi_measure_sequence_cluster`.
- No runtime changes, no INI changes, no composite implementation changes were made.

## PASS_03 Recap

- `ratio_or_slash_cluster`: **161**
- `multi_measure_sequence_cluster`: **23**
- Combined in-scope population: **184**
- Interpretation: these are the two highest-priority composite families for bounded contract definition, but still non-authoritative for runtime mutation.

## Ratio / Slash Contract

### Recognition Rules
- Accepted separators: slash pair (`A/B`) or per-form (`A per B`, `A_per_B`).
- Accepted component count: exactly 2 components only (bounded safe scope).
- Ordering is mandatory (`A/B` != `B/A`).
- Repeated components are not allowed.
- Custom placeholders are not allowed in bounded pass contract.
- Free text wrappers are only acceptable when they do not change two-component ratio semantics; otherwise defer.

### Component Rules
- Each component must resolve to canonical measure or deterministic alias-to-canonical target.
- Missing or conflicting component interpretation fails closed.
- Cross-domain component mixes defer/reject when scope is ambiguous.

### Assembly Rules
- Conceptual assembly is ordered two-component join with separator metadata.
- Labels between values are out of bounded contract unless ratio remains unambiguous.
- Contract is semantic-only; no rendering behavior is implemented.

### Fail-Closed Rules
- Reject path-like/file expressions (example: `E:/.../HEADSHOTS/...png`).
- Reject slash forms not equal to exactly two components.
- Reject unresolved components or duplicate/conflicting bindings.
- Reject split/comparison semantics entering ratio family.
- Defer tokenized/contextual wrappers requiring template interpretation.

### Representative Pass / Fail / Defer Examples
- Pass: `K/BB`, `H/AB`, `R/ER`, `at_bats_per_home_run`
- Fail: `E:/EDRIVE/MLB/HEADSHOTS/{{info.team.alias}}/...png`, `batting_average_home_road_differential`, `{{custom.header_avg}} VS PITCH TYPE`
- Defer: `PITCHES/PA - {{info.time.season}}`, `{{stats.player.season.hits}}/{{stats.player.season.at_bats}}`
- PASS_03 membership disposition under this contract: pass=130, defer=30, reject=1

## Multi-Measure Sequence Contract

### Recognition Rules
- Accepted shapes:
  - comma-delimited sequence with 2-3 stat components
  - slash+parenthetical bundle with 2-3 stat components
- Ordering is mandatory.
- Sequence length outside 2-3 is out-of-contract.
- Custom header placeholders are boundary signals; defer to later cluster handling.

### Component Rules
- All components must resolve to canonical measures or deterministic aliases.
- Any unresolved component fails closed.
- Mixed context/domain components defer/reject when unbounded.

### Assembly Rules
- Conceptual assembly is ordered component list + separator metadata + optional label metadata.
- Labels may appear before/between components, but render policy is out-of-scope.
- Contract remains semantic-only (no runtime rendering logic).

### Fail-Closed Rules
- Reject path-like/file-template expressions.
- Reject unsupported sequence shapes or component-count mismatch.
- Reject unresolved/conflicting component interpretations.
- Defer custom header/value couplings and strong template-context dependencies.

### Representative Pass / Fail / Defer Examples
- Pass: `{{stats.player.runners(loaded).batting_average}} AVG, {{stats.player.runners(loaded).home_runs}} GRAND SLAMS`,
  `After 0-1: {{stats.player.season.count_after(0-1).pitcher_opponent_batting_average}} AVG, {{stats.player.season.count_after(0-1).pitcher_home_runs_allowed}} HR`,
  `{{stats.player.career.vs_league(inter).pitcher_saves}}/{{stats.player.career.vs_league(inter).pitcher_saves_opportunities}} ({{stats.player.career.vs_league(inter).pitcher_saves_percentage}})`
- Fail: `E:/EDRIVE/MLB/HEADSHOTS/{{info.team.alias}}/...png`, `K/BB`, `{{info.time.season}} ({{stats.player.season.games_played}} GAMES)`
- Defer: `{{info.time.season}}: {{stats.player.season.batting_average}}, {{stats.player.season.home_runs}} {{custom.header_hr}}, {{stats.player.season.rbis}} {{custom.header_rbi}}`
- PASS_03 membership disposition under this contract: pass=16, defer=5, reject=2

## Contract Boundaries

- Out of scope in PASS_04: `text_plus_token_cluster`, `split_or_comparison_cluster`, `header_value_pair_cluster`, `contextual_template_cluster`, `unresolved_composite_cluster`.
- These remain outside this pass because they require additional grammar/context/render coupling not safe to merge into the two-cluster contract scope.

## Advisory-First Envelope

- Safe advisory-only first step:
  - recognize in-scope shape
  - extract ordered components
  - attempt read-only canonical binding
  - emit PASS/DEFER/REJECT preview with deterministic reason code
- Must remain forbidden:
  - runtime mutation/apply/take/cue
  - INI mutation
  - tabfield/socket execution
  - payload/output semantic mutation

## Why This Is Not An INI Problem

- INI category mappings are static one-key lookups; these two families require ordered multi-component interpretation and fail-closed composition checks.
- PASS_03 evidence includes contextual/tokenized/path-like variants that cannot be represented safely as direct INI mapping lines.
- Contract-first modeling prevents unsafe promotion while preserving operator-relevant composite value.

## Recommended Next Pass

- `MLB_ONAIR_COMPOSITE_ADVISORY_RECOGNITION_SPEC_PASS_05`
- Scope: define read-only advisory PASS/DEFER/REJECT reason-code schema for ratio/sequence clusters using this PASS_04 contract.
- Boundary: no runtime mutation, no INI mutation, no composite implementation.
