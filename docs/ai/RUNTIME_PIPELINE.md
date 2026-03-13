# SmartStat Runtime Pipeline

## Purpose

This file gives AI sessions a stable way to reason about SmartStat runtime and
runtime-adjacent lanes without flattening frozen, active, deferred, and planned
states into one blurred model.

---

## Important Distinction

This file intentionally separates:

- runtime core execution reality
- semantic and contract modeling lanes
- governance and readiness lanes
- future implementation lanes

---

## Layer 1 - Protected Runtime Core

Protected runtime/core baseline:
- `SmartStat_v4.0.0_beta.vbs`

Baseline config family:
- `SmartStat_Mappings.ini`
- `SmartStat_Mappings.learn.ini`
- `SmartStat_MappingsNBA.ini`
- `SmartStat_MappingsNBA.learn.ini`
- `SmartStat_MappingsNHL.ini`
- `SmartStat_MappingsNHL.learn.ini`
- `SmartStat_StaticOverrides.ini`
- `SmartStat_TemplateConfig.ini`

Core expectations:
- deterministic behavior
- fail-closed safety
- transactional safety
- operator-safe outputs
- no undocumented Viz Trio assumptions

---

## Layer 2 - High-Level Processing Model

Use this conceptual runtime model unless a narrower task requires otherwise:

1. input/page/tabfield state acquisition
2. config and mapping interpretation
3. candidate or semantic resolution
4. ambiguity gating
5. plan/validation shaping where applicable
6. terminal/output selection
7. output synthesis
8. apply/write safety checks
9. diagnostics, evidence, and logging

---

## Layer 3 - Repo-Truth Lane Status

The repo state currently separates:

- frozen runtime/core baselines
- active semantic architecture/tooling lane
- closed read-only contract and validation packages
- closed WP-20 governance packaging
- runtime implementation that is still separately gated

Do not overwrite repo-truth casually.

---

## Layer 4 - Session / Workstream Override Model

A conversation or coding session may discuss:

- WP-20 planning
- runtime bridge governance
- runtime slice rehearsal
- approval/evidence packets
- bridge-boundary proposals
- read-only runtime-lane scaffolds

When that happens, treat it as session-truth unless and until the repo itself
reflects the same state.

---

## Layer 5 - OnAir / Semantic Lane References

The semantic and contract side of the repo may include:

- plan capture
- plan validation
- slot resolution
- viewer contracts
- query skeletons
- formatter logic
- explainability contracts
- bridge-governance artifacts

These are architecture and contract aids unless explicitly merged into the
runtime core under approved rules.

---

## AI Interpretation Rules

When using this pipeline:

- never imply that a modeled lane is already live runtime behavior
- never imply that a governance packet equals implementation
- never imply that a fixture or validator equals merged runtime support
- always specify whether a statement is:
  - runtime-core truth
  - semantic-model truth
  - governance/readiness truth
  - proposed future-state

---

## Short Lane Map

Use this compact map in AI reasoning:

- protected runtime core
- semantic/contracts lane
- governance/readiness lane
- separately authorized runtime-expansion lane

---

## Runtime Safety Reminder

No proposal is acceptable if it:

- bypasses fail-closed behavior
- weakens transaction safety
- assumes undocumented Trio behavior
- implies apply/mutation authorization from read-only artifacts
- hides compatibility or regression risk
