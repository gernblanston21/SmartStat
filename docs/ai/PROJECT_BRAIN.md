# SmartStat Project Brain

## System

SmartStat is a deterministic broadcast-engine for Viz Trio graphics workflows.

It converts template, page, and tabfield state into controlled stat syntax,
config-driven mappings, and related outputs while preserving fail-closed
broadcast safety and operator-safe behavior.

---

## Environment

Platform:
    Viz Trio

Core language:
    VBScript

Configuration model:
    INI-driven

Repo style:
    governance-first
    regression-aware
    copy/paste-safe
    fail-closed

---

## Protected Runtime Core

Primary runtime baseline:
    `SmartStat_v4.0.0_beta.vbs`

Primary config family:
    `SmartStat_Mappings.ini`
    `SmartStat_Mappings.learn.ini`
    `SmartStat_MappingsNBA.ini`
    `SmartStat_MappingsNBA.learn.ini`
    `SmartStat_MappingsNHL.ini`
    `SmartStat_MappingsNHL.learn.ini`
    `SmartStat_StaticOverrides.ini`
    `SmartStat_TemplateConfig.ini`

---

## Architectural Split

SmartStat is modeled as two separated halves:

### A. Protected Production Runtime Core
Characteristics:
    deterministic
    protected
    regression-sensitive
    operator-safe
    fail-closed

### B. Semantic Architecture / Tooling Lane
Characteristics:
    modeling
    contracts
    validation
    explainability
    planning
    future-facing bridge preparation

Do not collapse these halves into one implementation story.

---

## Operational Priorities

1. determinism
2. fail-closed safety
3. operator workflow safety
4. regression safety
5. config compatibility
6. maintainability

---

## Hard Constraints

- no silent broadening of scope
- no unsupported Viz Trio assumptions
- no naming convention redesign
- no tabfield convention redesign
- no unsafe partial writes
- no contract-breaking INI edits
- no external compatibility blind spots
- no runtime authorization implied by docs/tests/governance packaging

---

## Grounding Sources

Primary grounding sources:
    `AGENTS.md`
    `SESSION.md`
    `ROADMAP.md`
    `docs/viz-trio/`

Secondary architecture sources:
    `docs/architecture/smartstat-architecture.md`
    `docs/onair/`
    `.agents/skills/`

AI support sources:
    `docs/ai/`

---

## Runtime / Semantic Thinking Model

When analyzing SmartStat, think in layers:

- input acquisition
- mapping/config interpretation
- candidate or semantic resolution
- ambiguity gating
- plan/validation shaping where applicable
- terminal/output shaping
- apply/write safety
- diagnostics/regression evidence

---

## Truth Model

Always distinguish:

- repo-truth
- session-truth
- implementation-truth
- proposed future-state

Never present them as the same thing unless explicitly confirmed.

---

## AI Behavior Contract

The AI should respond like a deterministic engineering partner:

- grounded
- explicit
- version-aware
- fail-closed
- copy/paste ready
- regression conscious
- boundary conscious

---

## Delivery Preference

For substantial SmartStat work, prefer:

- Role Summary
- Summary
- Assumptions
- Implementation
- Regression Impact
- Risks
- Review
- Exact Code
- Validation Steps
