# OnAir Slot Resolution Model (Phase 0 Architecture Reference)

## 1. Purpose / Overview

This document defines a slot-oriented interpretation model for workbook-observed OnAir syntax.

In this context, slot resolution means interpreting an expression as an ordered sequence of typed semantic slots, then resolving candidates within the currently expected slot class.

This artifact exists to bridge:
- grammar interpretation in [onair_semantic_grammar.md](e:/EDRIVE/UNIVERSAL/SmartStat/docs/onair/onair_semantic_grammar.md)
- Phase 9 candidate-resolution scaffolding in `tools/semantic-source-view`
- future deterministic plan capture / Plan Engine work

This is an architecture/reference artifact only. It does not implement runtime resolution, planner execution, or SmartStat runtime behavior.

## 2. Source Context

This model is derived from:
- workbook-derived grammar and dictionaries under `docs/onair/`
- workbook query examples in [query-skeletons.md](e:/EDRIVE/UNIVERSAL/SmartStat/docs/onair/query-skeletons.md)
- operator-layer interpretation documented in [onair_semantic_grammar.md](e:/EDRIVE/UNIVERSAL/SmartStat/docs/onair/onair_semantic_grammar.md)
- Phase 9 candidate-resolution scaffolding direction (read-only architectural linkage)

No direct implementation claim is made in this document.

## 3. Slot Classes

| slot_class | description | example_tokens | interior_or_terminal | behavior notes |
| --- | --- | --- | --- | --- |
| family slot | Base query family context | `info`, `stats` | interior | Establishes high-level query shape and expected downstream slot types. |
| operator slot | Operator-style wrapper/category token | `leader(1, home_runs)`, `rank(asc)`, `previous`, `streak(runs, >=3)`, `calendar(1, sun)` | interior | Distinct from base family; may constrain expected terminal slot class and filter behavior. |
| entity slot | Primary subject/scope entity | `player`, `team`, `coach`, `time`, `pitcher`, `goalie` | interior | Drives available attribute/measure/filter candidate pools. |
| scope slot | Broad scope selector | `career`, `season`, `postseason`, `all_time` | interior | Narrows context before or alongside filters; workbook-observed as scope-like modifiers. |
| filter slot | Parameterized qualifier/filter token | `month(april)`, `innings(7-9)`, `location(away)`, `vs(DET)` | interior | Can repeat; resolved against canonical filter dictionary and alias map. |
| terminal measure slot | Final stats target token | `hits`, `home_runs`, `team_wins` | terminal | Terminates stat-oriented chains. |
| terminal attribute slot | Final info/display target token | `full_name`, `alias`, `altcity` | terminal | Terminates attribute-oriented chains. |
| formatter slot | Output transform/post-processing token | `| ordinal`, `| long_year`, `| lowercase` | terminal (post-expression) | Applies after main semantic chain; does not replace semantic terminal token. |

## 4. Example Slot Walkthroughs

### Example A

`{{stats.player.career.month(april).innings(7-9).hits}}`

- family slot = `stats`
- entity slot = `player`
- scope slot = `career`
- filter slot = `month(april)`
- filter slot = `innings(7-9)`
- terminal measure slot = `hits`

### Example B

`{{leader(1, home_runs).player.season(2023).location(away).full_name}}`

- operator slot = `leader(1, home_runs)`
- entity slot = `player`
- filter slot = `season(2023)`
- filter slot = `location(away)`
- terminal attribute slot = `full_name`

### Example C (additional workbook-observed operator form)

`{{rank(asc).player(TB, 56).season(2023).is_qualified_hitting.strikeouts | ordinal}}`

- operator slot = `rank(asc)`
- entity slot = `player(TB, 56)`
- filter slot = `season(2023)`
- filter slot = `is_qualified_hitting`
- terminal measure slot = `strikeouts`
- formatter slot = `| ordinal`

## 5. Slot-to-Candidate Resolution Mapping

Candidate resolution can be interpreted as slot-local lookup against the relevant dictionary pool.

| slot_class | candidate pool shape | primary reference source | phase-9 linkage note |
| --- | --- | --- | --- |
| family slot | base family tokens | [query-skeletons.md](e:/EDRIVE/UNIVERSAL/SmartStat/docs/onair/query-skeletons.md), [onair_semantic_grammar.md](e:/EDRIVE/UNIVERSAL/SmartStat/docs/onair/onair_semantic_grammar.md) | Candidate ranking can treat family as early high-confidence gating context. |
| operator slot | operator-style tokens | [query-skeletons.md](e:/EDRIVE/UNIVERSAL/SmartStat/docs/onair/query-skeletons.md), [onair_semantic_grammar.md](e:/EDRIVE/UNIVERSAL/SmartStat/docs/onair/onair_semantic_grammar.md) | Generic candidate scaffolding can represent operator candidates before planner formalization. |
| entity slot | entity names and parameter forms | [entity-dictionary.md](e:/EDRIVE/UNIVERSAL/SmartStat/docs/onair/entity-dictionary.md) | Candidate model can hold preferred/alternate entity candidates per slot. |
| scope slot | scope-like filters/selectors | [filter-grammar-dictionary.md](e:/EDRIVE/UNIVERSAL/SmartStat/docs/onair/filter-grammar-dictionary.md) | Scope candidates can be resolved with deterministic ordering before terminal slots. |
| filter slot | canonical filters + aliases | [filter-grammar-dictionary.md](e:/EDRIVE/UNIVERSAL/SmartStat/docs/onair/filter-grammar-dictionary.md), [alias-map.qualifiers.md](e:/EDRIVE/UNIVERSAL/SmartStat/docs/onair/alias-map.qualifiers.md) | Phase 9-style candidate rationale can annotate inferred vs direct filter matches. |
| terminal measure slot | measure tokens | [measure-dictionary.md](e:/EDRIVE/UNIVERSAL/SmartStat/docs/onair/measure-dictionary.md) | Terminal measure candidates can be selected only when chain context supports stats-style output. |
| terminal attribute slot | attribute tokens | [attribute-dictionary.md](e:/EDRIVE/UNIVERSAL/SmartStat/docs/onair/attribute-dictionary.md) | Terminal attribute candidates can be selected when entity and family/operator context allows info-style output. |
| formatter slot | formatter tokens | [formatter-dictionary.md](e:/EDRIVE/UNIVERSAL/SmartStat/docs/onair/formatter-dictionary.md) | Formatter candidates can remain optional post-chain slot candidates. |

Phase 9 candidate-resolution scaffolding can be viewed as an early generic prototype for this slot-local candidate resolution behavior.

## 6. Ambiguity by Slot Type

Ambiguity shape is not uniform across slots:

- operator ambiguity: category/operator label overlap (`game high` vs `game_high`, macro-style wrappers vs base families)
- entity ambiguity: multiple entities may be superficially valid before context narrowing (`team` vs `player`)
- filter ambiguity: alias-driven and parameter-form ambiguity (`last#`, `prev#`, before/after semantics)
- terminal measure ambiguity: many measure candidates can share lexical similarity
- terminal attribute ambiguity: attribute names can overlap across entity contexts
- formatter ambiguity: formatter-like tokens can appear syntactically but remain optional

Future deterministic planner behavior will likely need slot-class-aware ambiguity policies rather than one global ambiguity rule.

## 7. Slot Constraints and Ordering

Workbook-evidenced ordering interpretation:

- family slot or operator slot tends to appear first
- entity selection generally precedes terminal measure/attribute selection
- scope/filter slots usually appear before terminal slot selection
- formatter slots appear after the main semantic expression chain
- terminal measure vs terminal attribute selection is mutually constraining by query shape

These are observed constraints, not finalized runtime grammar rules.

## 8. Relationship to Plan Capture

Slot-aware resolution provides a deterministic bridge toward future plan capture:

- grammar interpretation identifies expected next slot class
- candidate resolution resolves against that slot-local candidate pool
- future plan capture can record the ordered resolved slot sequence with provenance

This is an architecture bridge toward WP-17 plan capture direction, not plan execution behavior.

This slot-resolution model leads directly to the WP-17 capture contract in [plan-capture-contract.md](e:/EDRIVE/UNIVERSAL/SmartStat/docs/onair/plan-capture-contract.md), with conceptual background in [plan-capture-shape.md](e:/EDRIVE/UNIVERSAL/SmartStat/docs/onair/plan-capture-shape.md).

## 9. Deferred / Non-Goals

This document does not define:

- runtime resolution implementation
- execution semantics
- planner implementation details
- final serialization format
- runtime bridge behavior

It is a pre-implementation architecture artifact intended to reduce ambiguity between grammar interpretation, candidate resolution, and future deterministic plan capture.
