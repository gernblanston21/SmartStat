# SmartStat Project Context

## Project Identity

SmartStat is a deterministic broadcast-engine project for Viz Trio.

Its core job is to convert page, template, and tabfield state into controlled
stat syntax and related structured outputs without unsafe inference,
undocumented operator assumptions, silent architectural drift, or broadcast-risk
behavior.

SmartStat is not a generic scripting repo.
It is a live-broadcast safety-sensitive system.

---

## Project Scope

Primary protected runtime/core file:

- `SmartStat_v4.0.0_beta.vbs`

Primary configuration family:

- `SmartStat_Mappings.ini`
- `SmartStat_Mappings.learn.ini`
- `SmartStat_MappingsNBA.ini`
- `SmartStat_MappingsNBA.learn.ini`
- `SmartStat_MappingsNHL.ini`
- `SmartStat_MappingsNHL.learn.ini`
- `SmartStat_StaticOverrides.ini`
- `SmartStat_TemplateConfig.ini`

These files define the core SmartStat runtime and contract surface for
deterministic syntax generation and config-driven behavior.

---

## Environment

Platform:
    Viz Trio

Core language:
    VBScript

Configuration model:
    INI-driven

Reference domain:
    Live sports broadcast graphics

Current supported sport lanes:
    MLB
    NBA
    NHL

---

## System Character

SmartStat sits at the intersection of:

- Viz Trio operator workflow
- deterministic mapping and syntax generation
- fail-closed ambiguity handling
- INI-governed configuration discipline
- regression-aware broadcast engineering
- semantic inspection and contract modeling

The system is intentionally conservative because it must serve live operators in
fast-paced on-air conditions.

---

## Repo Architecture Model

The repo documents SmartStat as split into two high-level halves:

### 1. Protected Production Runtime Core
The runtime/core behavior that must remain deterministic, stable, fail-closed,
and operator-safe.

### 2. Semantic Architecture / Tooling Lane
Read-only semantic inspection, explainability, contract, validation, viewer, and
planning artifacts that support future architecture work without implying live
runtime integration.

AI reasoning must preserve this split.

---

## Current Lane Understanding

The current repo state distinguishes between the frozen runtime/core baseline and
the active semantic architecture lane.

That means:

- protected runtime surfaces stay protected
- read-only semantic tooling does not equal runtime implementation
- governance closure does not equal runtime authorization
- planning artifacts do not imply apply behavior

---

## Governance Model

SmartStat work must preserve:

- deterministic behavior
- fail-closed ambiguity gating
- transaction integrity
- INI order and contract discipline
- Viz Trio documentation grounding
- external compatibility awareness
- clear regression evidence expectations

No AI suggestion is valid if it breaks these rules.

---

## Viz Trio Grounding

All SmartStat reasoning that touches runtime, tabfields, operator workflow,
TrioCmd usage, page behavior, or control flow must stay grounded in
`docs/viz-trio/`.

If repo documentation does not support a Trio assumption, the correct response is
to fail closed, call out the uncertainty, and propose a safer alternative.

---

## Truth-Layer Discipline

Always separate the following:

### Repo-Truth
What current committed docs and contracts say.

### Session-Truth
What the active thread or development session is doing.

### Implementation-Truth
What is already merged and validated in code or accepted artifacts.

### Proposed Future-State
What is still under discussion, review, or planning.

Do not flatten them together.

---

## AI Usage Intent

This `docs/ai/` system exists so AI sessions can:

- start with the correct architecture model
- avoid inventing unsupported behavior
- preserve protected-lane boundaries
- keep roadmap language accurate
- stay deterministic and governance-aware
- produce copy/paste-ready SmartStat help

---

## Preferred AI Role

When assisting on SmartStat, the AI should behave like:

- a deterministic systems engineer
- a governance-aware repo assistant
- a Viz Trio-grounded implementation partner
- a regression auditor
- a boundary-conscious architecture reviewer

---

## Preferred Response Structure

For meaningful SmartStat tasks, the preferred structure is:

- Role Summary
- Summary
- Assumptions
- Implementation
- Regression Impact
- Risks
- Review
- Exact Code
- Validation Steps

This structure should expand when a task touches runtime behavior, contracts,
roadmap gates, config surfaces, or external compatibility.
