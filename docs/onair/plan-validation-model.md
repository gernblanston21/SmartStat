# OnAir Plan Validation Model (Phase 0 Architecture Reference)

> WP-17 note: plan capture artifact shape is now enforceable via
> `docs/onair/plan-capture-contract.md` and `docs/onair/plan-capture.schema.json`.
> This document remains the architecture model for WP-18 validation policy.

## 1 Purpose / Overview

This document defines an architecture-level validation model for captured OnAir semantic plans.

Validation sits between plan capture and future execution planning. Its role is to reject semantically incompatible captured plans before any execution-oriented planning stage.

This is a reference model only. It does not define runtime validation code.

## 2 Source Context

Validation is derived from:

- [onair_semantic_grammar.md](e:/EDRIVE/UNIVERSAL/SmartStat/docs/onair/onair_semantic_grammar.md)
- semantic dictionaries (`entity`, `filter`, `measure`, `attribute`, `formatter`, aliases)
- [slot-resolution-model.md](e:/EDRIVE/UNIVERSAL/SmartStat/docs/onair/slot-resolution-model.md)
- [plan-capture-shape.md](e:/EDRIVE/UNIVERSAL/SmartStat/docs/onair/plan-capture-shape.md)
- [plan-capture-contract.md](e:/EDRIVE/UNIVERSAL/SmartStat/docs/onair/plan-capture-contract.md)

## 3 Validation Layer in Architecture

Conceptual pipeline:

OnAir syntax  
→ semantic grammar  
→ slot resolution  
→ candidate resolution  
→ plan capture  
→ plan validation  
→ execution planning

Validation confirms that captured slots are compatible in sequence and semantics.

## 4 Slot Compatibility Rules

| slot type | allowed predecessors |
| --- | --- |
| `entity` | `family` or `operator` |
| `scope` | `entity` |
| `filter` | `entity` or `scope` or `filter` |
| `terminal measure` | `entity` or `scope` or `filter` |
| `terminal attribute` | `entity` or `scope` or `filter` |
| `formatter` | `terminal measure` or `terminal attribute` |

Interpretation note: this table is a conservative architecture baseline derived from workbook-observed chains.

## 5 Entity Compatibility

Entity context constrains valid terminal candidates:

- `player` entity should resolve to player-compatible measures/attributes
- `team` entity should resolve to team-compatible measures/attributes
- `coach` entity should resolve to coach-compatible attributes and compatible stat contexts

Invalid shape example:

- `team` entity + player-only terminal attribute

Validation should reject incompatible entity-terminal pairings.

## 6 Terminal Constraints

Terminal type must align with family/operator context.

Baseline interpretation:

- `stats` family generally requires a terminal measure
- `info` family generally requires a terminal attribute
- operator-style chains may constrain terminal type (for example leader/rank flows usually depend on measure context, while projected output can be attribute in some operator forms)

Validation rejects terminal kinds that conflict with resolved context.

## 7 Operator Constraints

Operator slots can impose additional constraints beyond base family rules.

Examples (architecture-level):

- `leader(...)` expects measure-driven context for ranking/selection
- `rank(...)` expects a rankable numeric measure context
- `previous` expects a valid condition/filter context before terminal output
- `calendar(...)` expects calendar-compatible entity/time contexts

These are interpretation constraints, not runtime algorithms.

## 8 Formatter Constraints

Formatter constraints:

- formatter appears only after a resolved terminal slot
- formatter should be compatible with terminal output type (for example date/time formatters vs numeric/text formatters)

Validation rejects formatter placement before terminal resolution.

## 9 Example Validated Plans

Example valid:

- `{{stats.player.month(april).hits}}`
  - family `stats` + entity `player` + filter `month(april)` + terminal measure `hits`

Example valid:

- `{{leader(1, home_runs).player.season(2023).full_name}}`
  - operator `leader(...)` with measure context and entity/filter chain; terminal attribute allowed as projected output in this operator form

## 10 Example Invalid Plans

Example invalid:

- `{{stats.team.full_name}}`
  - invalid because `stats` context with attribute terminal violates baseline terminal compatibility

Example invalid:

- `{{info.player.hits}}`
  - invalid because `info` context with measure terminal violates baseline terminal compatibility

Example invalid:

- `{{stats.player.hits | long_year}}`
  - invalid if `long_year` formatter requires date-like output and terminal measure is numeric

## 11 Deferred / Non-Goals

This document does not define:

- runtime validation implementation
- planner algorithm implementation
- final error-code contract
- runtime integration behavior

It is a Phase 0 architecture artifact for semantic-plan validation intent.
