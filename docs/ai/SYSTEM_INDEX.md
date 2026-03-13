# SmartStat AI System Index

## Purpose

This file is the master table of contents for the SmartStat AI context system.

All AI sessions should load this file first when possible.
It defines the authoritative context hierarchy, the role of each `docs/ai/*`
file, and the architectural model that all later reasoning must follow.

This file is designed to reduce:

- architecture drift
- inconsistent chat startup behavior
- runtime vs semantic-lane confusion
- over-reading scratch/evidence artifacts as implementation truth
- duplicated or conflicting AI context rules

---

## Authority Order

When reasoning about SmartStat, use this source hierarchy:

1. `AGENTS.md`
2. `SESSION.md`
3. `ROADMAP.md`
4. `docs/viz-trio/`
5. `docs/architecture/smartstat-architecture.md`
6. `docs/onair/`
7. `docs/contracts/`
8. `docs/ai/*`
9. `.agents/skills/*`

If a lower item conflicts with a higher item, the higher item wins.

---

## SmartStat Architecture Summary

SmartStat is a deterministic Viz Trio broadcast-engine project with two
intentionally separated layers:

### A. Protected Production Runtime Core

Primary properties:

- deterministic
- fail-closed
- transaction-safe
- operator-safe
- regression-sensitive
- broadcast-safe

Primary runtime/config surface:

- `SmartStat_v4.0.0_beta.vbs`
- `SmartStat_Mappings*.ini`
- `SmartStat_StaticOverrides.ini`
- `SmartStat_TemplateConfig.ini`

### B. Semantic Architecture / Tooling Layer

Primary properties:

- read-only by default
- contract-oriented
- validation-oriented
- explainability-oriented
- planning-oriented
- non-authorizing

Primary examples:

- semantic source view
- plan capture contracts
- plan validation layer
- viewer contracts
- runtime bridge governance artifacts

The semantic layer does **not** imply runtime integration.

---

## Truth Layers

Always distinguish:

### Repo-Truth
What committed repo sources currently say.

### Session-Truth
What the active chat, branch lane, or planning thread is doing.

### Implementation-Truth
What is actually implemented and validated.

### Proposed Future-State
What is planned, drafted, under review, or not yet authorized.

Do not silently flatten these layers into one story.

---

## File Map

### Core AI Context Files

- `docs/ai/README.md`
- `docs/ai/project-context.md`
- `docs/ai/context_seed.md`
- `docs/ai/PROJECT_BRAIN.md`
- `docs/ai/ARCHITECTURE_ANCHOR.md`
- `docs/ai/DEVELOPMENT_RULES.md`
- `docs/ai/RUNTIME_PIPELINE.md`
- `docs/ai/SMARTSTAT_AI_BOOTSTRAP.md`
- `docs/ai/chat-starters.md`

### Advanced AI Context Files

- `docs/ai/SYSTEM_INDEX.md`
- `docs/ai/SMARTSTAT_RUNTIME_MAP.md`
- `docs/ai/AI_KERNEL.md`
- `docs/ai/AGENTS_SKILL_ROUTER.md`
- `docs/ai/AI_SESSION_PROTOCOL.md`
- `docs/ai/KERNEL_CHANGELOG.md`

---

## Standard Load Order

For a full SmartStat AI session, use this order:

1. `docs/ai/SYSTEM_INDEX.md`
2. `AGENTS.md`
3. `SESSION.md`
4. `ROADMAP.md`
5. `docs/ai/PROJECT_BRAIN.md`
6. `docs/ai/ARCHITECTURE_ANCHOR.md`
7. `docs/ai/DEVELOPMENT_RULES.md`
8. `docs/ai/RUNTIME_PIPELINE.md`
9. `docs/ai/SMARTSTAT_RUNTIME_MAP.md`
10. `docs/viz-trio/`
11. `docs/architecture/smartstat-architecture.md`
12. `docs/onair/`
13. `docs/ai/AI_KERNEL.md`
14. `docs/ai/SMARTSTAT_AI_BOOTSTRAP.md`

Use `project-context.md` when deeper orientation is needed.

Use `context_seed.md` for lightweight ChatGPT continuity.

---

## AI Kernel Pattern

The AI Kernel pattern turns the repo into a stable AI operating environment.

It consists of:

- one authority index (`SYSTEM_INDEX.md`)
- one system model (`PROJECT_BRAIN.md`)
- one architecture lock (`ARCHITECTURE_ANCHOR.md`)
- one rule layer (`DEVELOPMENT_RULES.md`)
- one runtime reasoning layer (`RUNTIME_PIPELINE.md`)
- one visual/runtime topology map (`SMARTSTAT_RUNTIME_MAP.md`)
- one startup seed (`context_seed.md`)
- one session bootstrap (`SMARTSTAT_AI_BOOTSTRAP.md`)
- one skill router (`AGENTS_SKILL_ROUTER.md`)
- one operating model (`AI_KERNEL.md`)
- one session contract (`AI_SESSION_PROTOCOL.md`)

This pattern makes new AI sessions reconstruct the same SmartStat mental model
reliably across chats.

---

## Hard Rule

No `docs/ai/*` file authorizes code or config changes by itself.

All implementation proposals must still satisfy:

- `AGENTS.md`
- `SESSION.md`
- `ROADMAP.md`
- `docs/viz-trio/`
- relevant contracts and evidence requirements
