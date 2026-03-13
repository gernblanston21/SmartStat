# AGENTS Skill Auto-Discovery Protocol

## SmartStat Codex Skill Router

## Purpose

This file implements the skill auto-discovery pattern used in larger
AI-assisted repositories. It allows Codex to determine the correct operational
skill before planning begins, which improves reasoning accuracy and reduces
incorrect workflow selection.

This protocol works together with:

- `AGENTS.md`
- `.agents/skills/*`
- `docs/ai/SYSTEM_INDEX.md`
- the SmartStat AI Kernel

---

## Skill Router Initialization

Before planning a task, Codex must perform the following routing step:

1. inspect the task description
2. identify which domain category the task belongs to
3. load the corresponding skill from `.agents/skills/`
4. follow that skill’s workflow before producing any plan or diff

If multiple skills match, prefer the most specific skill.

---

## SmartStat Skill Routing Table

### Viz Trio Runtime / Operator Behavior

If the task involves:

- TrioCmd commands
- Viz Trio scripting
- tabfields
- operator workflow
- template behavior
- show control interaction

Then load:

`.agents/skills/viztrio-grounding/SKILL.md`

### INI Contract / Mapping System

If the task involves:

- SmartStat INI files
- mapping keys
- alias handling
- INI ordering rules
- duplicate keys
- configuration contracts

Then load:

`.agents/skills/smartstat-ini-governance/SKILL.md`

### Determinism / Runtime Evidence

If the task involves:

- determinism validation
- hashing comparisons
- regression verification
- run comparisons
- evidence capture

Then load:

`.agents/skills/smartstat-determinism-audit/SKILL.md`

### RC Stabilization / Governance

If the task involves:

- RC rules
- release discipline
- validation-only changes
- log clarity
- stability verification

Then load:

`.agents/skills/rc-stabilization-discipline/SKILL.md`

### Repository Mechanics

If the task involves:

- unified diffs
- line extraction
- file refactoring
- patch construction
- repo operations

Then load:

`.agents/skills/repo-ops-codex/SKILL.md`

### Core Engine Workflow

If the task involves:

- SmartStat core workflow execution
- end-to-end SmartStat operational process
- repo-standard SmartStat work sequencing

Then load:

`.agents/skills/smartstat-core-engine-workflow/SKILL.md`

### Roadmap Driver / WP Sequencing

If the task involves:

- roadmap progression
- WP sequencing
- milestone planning
- package closeout sequencing

Then load:

`.agents/skills/smartstat-roadmap-driver/SKILL.md`

### Viz Package QA

If the task involves:

- Viz package QA
- package validation
- asset-level acceptance checks

Then load:

`.agents/skills/viz-package-qa/SKILL.md`

---

## Multi-Skill Tasks

If a task spans multiple domains:

1. load the primary domain skill first
2. apply secondary skills only when required
3. never bypass SmartStat governance rules

---

## Hard Safety Rules

Skill routing must never override repository governance.

Authority order remains:

1. `AGENTS.md`
2. `SESSION.md`
3. `ROADMAP.md`
4. `docs/viz-trio/`
5. `docs/architecture/smartstat-architecture.md`
6. `docs/onair/`
7. `docs/ai/*`
8. `.agents/skills/*`

Skills guide workflow but do not authorize changes.

---

## Planning Requirement

Once the correct skill is loaded, Codex must follow the standard response
structure used in SmartStat development:

- Role Summary
- Summary
- Assumptions
- Implementation
- Regression Impact
- Risks
- Review
- Exact Code
- Validation Steps
