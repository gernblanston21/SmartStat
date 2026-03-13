# SmartStat AI Session Protocol

## Purpose

This file defines the expected operating behavior for SmartStat AI sessions.

It gives ChatGPT and Codex a stable session contract so they enter the repo with
the same authority order, the same truth-layer discipline, and the same
expectation for how to reason before planning or writing code.

---

## Session Contract

Every SmartStat AI session should do the following before producing solutions:

1. load the authority order from `SYSTEM_INDEX.md`
2. align with governance files
3. align with architecture files
4. align with runtime interpretation files
5. route to the correct skill before planning when applicable
6. preserve truth-layer separation
7. fail closed on unsupported Trio behavior or unsafe assumptions

---

## Authority Load Rule

Default authority load:

1. `SYSTEM_INDEX.md`
2. `AGENTS.md`
3. `SESSION.md`
4. `ROADMAP.md`
5. `docs/viz-trio/`
6. `docs/architecture/smartstat-architecture.md`
7. `docs/onair/`
8. `docs/contracts/`
9. `docs/ai/*`
10. `.agents/skills/*`

---

## Required Session Behaviors

### Governance First
No AI context file or skill file overrides governance.

### Trio Grounding First for Trio Tasks
No Trio command or operator workflow may be inferred without `docs/viz-trio/`
support.

### Skill Before Plan
If a task clearly belongs to a known repo skill, route to that skill before
producing plans or diffs.

### Truth-Layer Separation
Keep repo-truth, session-truth, implementation-truth, and proposed future-state
separate.

### Fail Closed
If repo evidence is insufficient, say so and propose a safer alternative.

---

## Preferred Response Shape

For technical SmartStat work, prefer:

- Role Summary
- Summary
- Assumptions
- Implementation
- Regression Impact
- Risks
- Review
- Exact Code
- Validation Steps
