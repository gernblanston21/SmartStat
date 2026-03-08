# OnAir Plan Capture Shape (Phase 0 Architecture Reference)

> WP-17 note: the enforceable capture contract is now defined in
> `docs/onair/plan-capture-contract.md` and `docs/onair/plan-capture.schema.json`.
> This document remains conceptual background and rationale.

## 1. Purpose / Overview

This document defines a conceptual deterministic plan-capture shape for OnAir expressions after slot-aware semantic resolution.

In SmartStat architecture terms, plan capture means recording the resolved semantic structure in a stable, ordered artifact that can later support deterministic execution planning.

This artifact connects:

- OnAir semantic grammar interpretation
- slot-resolution behavior
- Phase 9 candidate-resolution scaffolding
- future execution-planning direction

This document is architecture-only:

- it does not define runtime execution
- it does not implement planning logic

## 2. Source Context

The plan-capture shape is derived from:

- [onair_semantic_grammar.md](e:/EDRIVE/UNIVERSAL/SmartStat/docs/onair/onair_semantic_grammar.md)
- [slot-resolution-model.md](e:/EDRIVE/UNIVERSAL/SmartStat/docs/onair/slot-resolution-model.md)
- workbook-derived semantic dictionaries (`entity`, `filter`, `measure`, `attribute`, `formatter`, alias references)
- Phase 9 candidate-resolution scaffolding (architecture linkage only)

Plan capture is interpreted as occurring after semantic slot resolution has produced selected slot candidates.

## 3. High-Level Plan Object Shape

Conceptual JSON-like shape:

```json
{
  "query_family": "stats",
  "operator": "leader(1, home_runs)",
  "entity": "player",
  "filters": [
    "season(2023)",
    "location(away)"
  ],
  "terminal": {
    "type": "attribute",
    "token": "full_name"
  },
  "formatter": "ordinal",
  "slot_sequence": [
    "... ordered slot captures ..."
  ]
}
```

Field meaning (conceptual):

| field | meaning |
| --- | --- |
| `query_family` | resolved base family token when present (`info`, `stats`) |
| `operator` | resolved operator-style token when present (`leader(...)`, `rank(...)`, `previous`, etc.) |
| `entity` | resolved entity context token |
| `filters` | ordered resolved scope/filter-like tokens |
| `terminal` | resolved terminal target (`measure` or `attribute`) |
| `formatter` | resolved output formatter token when present |
| `slot_sequence` | deterministic ordered slot capture list with provenance metadata |

This shape is a reference model only, not a final implementation schema.

## 4. Slot Capture Structure

Each resolved slot can be captured as a deterministic plan node.

Conceptual node shape:

```json
{
  "slot_type": "filter",
  "token": "month",
  "parameters": ["april"],
  "source_dictionary": "filter-grammar-dictionary",
  "candidate_status": "preferred"
}
```

Interpretation notes:

- slots are captured in deterministic expression order
- each slot node carries semantic provenance
- candidate status reflects selection context (for example preferred vs alternate) without defining runtime algorithms

Final serialization keys may differ in future implementation.

## 5. Example Captured Plans

### Example A

Source query:
`{{stats.player.career.month(april).innings(7-9).hits}}`

Conceptual capture (ordered):

- family slot: `stats`
- entity slot: `player`
- scope slot: `career`
- filter slot: `month(april)`
- filter slot: `innings(7-9)`
- terminal measure slot: `hits`

### Example B

Source query:
`{{leader(1, home_runs).player.season(2023).location(away).full_name}}`

Conceptual capture (ordered):

- operator slot: `leader(1, home_runs)`
- entity slot: `player`
- filter slot: `season(2023)`
- filter slot: `location(away)`
- terminal attribute slot: `full_name`

### Example C

Source query:
`{{rank(asc).player(TB, 56).season(2023).is_qualified_hitting.strikeouts | ordinal}}`

Conceptual capture (ordered):

- operator slot: `rank(asc)`
- entity slot: `player(TB, 56)`
- filter slot: `season(2023)`
- filter slot: `is_qualified_hitting`
- terminal measure slot: `strikeouts`
- formatter slot: `ordinal`

## 6. Slot Provenance

Plan capture can preserve per-slot provenance for explainability and debugging.

Provenance dimensions (conceptual):

- dictionary source (entity/filter/measure/attribute/formatter)
- alias normalization evidence (when alias mapping is used)
- candidate-resolution selection context (preferred/alternate/rejected summaries)
- workbook-derived reference lineage

This provenance layer supports downstream deterministic diagnostics without implying runtime execution behavior.

## 7. Deterministic Ordering

Deterministic slot ordering is a core plan-capture requirement.

Observed ordering model:

1. family/operator
2. entity
3. scope
4. filters
5. terminal (measure or attribute)
6. formatter

Preserving this order reduces ambiguity and supports deterministic planning behavior.

## 8. Relationship to Candidate Resolution

Relationship summary:

- candidate resolution selects the token for the expected slot class
- plan capture records that selected token in ordered slot context

Phase 9 candidate-resolution scaffolding can be interpreted as the pre-plan selection stage; plan capture is the ordered recording stage.

## 9. Relationship to Future Plan Execution

A deterministic captured plan shape can later support execution-planning layers such as:

- SmartStat runtime query construction from resolved slots
- semantic compatibility checks before execution steps
- deterministic stat-computation pathway planning

This section describes architecture intent only; it does not define execution algorithms.

## 10. Deferred / Non-Goals

This artifact does not define:

- runtime query execution behavior
- SmartStat runtime integration details
- planner algorithm implementation
- final serialization schema
- external API contract

This is a Phase 0 architecture clarification document only.
