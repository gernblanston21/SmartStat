# SmartStat AI Context System

This directory contains the repo-local AI grounding system for SmartStat.

It exists to make ChatGPT, Codex, and other AI tools start from the correct
project model before suggesting architecture, runtime, config, or workflow
changes.

These files do not replace repo governance. They are a structured context layer
designed to reduce:

- architecture drift
- hallucinated runtime assumptions
- unsafe Viz Trio guesses
- roadmap confusion
- repo-truth vs session-truth flattening
- skill-routing inconsistency across AI entry points

---

## Authoritative Source Hierarchy

When an AI session reasons about SmartStat, use this hierarchy:

1. `AGENTS.md`
2. `SESSION.md`
3. `ROADMAP.md`
4. `docs/viz-trio/`
5. `docs/architecture/smartstat-architecture.md`
6. `docs/onair/`
7. `docs/contracts/`
8. `docs/ai/*`
9. `.agents/skills/*`

If content in a lower-priority source conflicts with a higher-priority source,
the higher-priority source wins.

---

## v1.0 Context Loading Order

For a full SmartStat session boot, load context in this order:

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

Use `project-context.md` when a deeper orientation pass is needed.

Use `context_seed.md` when you need a fast startup in a new chat.

Use `chat-starters.md` for copy/paste launch blocks.

---

## File Roles

### Core Context Files

- `SYSTEM_INDEX.md`
- `PROJECT_BRAIN.md`
- `ARCHITECTURE_ANCHOR.md`
- `DEVELOPMENT_RULES.md`
- `RUNTIME_PIPELINE.md`
- `SMARTSTAT_RUNTIME_MAP.md`

### Session Boot Files

- `context_seed.md`
- `SMARTSTAT_AI_BOOTSTRAP.md`
- `chat-starters.md`
- `AI_SESSION_PROTOCOL.md`

### Operating Model Files

- `AI_KERNEL.md`
- `AGENTS_SKILL_ROUTER.md`
- `KERNEL_CHANGELOG.md`

### Supporting Orientation Files

- `project-context.md`

---

## Design Principles

This context system is built around the following principles:

- governance first
- runtime safety first
- deterministic reasoning
- explicit truth-layer separation
- no unsupported Viz Trio assumptions
- no silent scope broadening
- skill-aware task routing
- convergent startup behavior across ChatGPT and Codex

---

## AI Kernel Pattern

The SmartStat AI context system supports the AI Kernel pattern.

The AI Kernel adds:

- `SYSTEM_INDEX.md` as the master authority index
- `SMARTSTAT_RUNTIME_MAP.md` as the visual/runtime topology layer
- `AI_KERNEL.md` as the operating model for multi-chat consistency
- `AGENTS_SKILL_ROUTER.md` as the pre-planning skill router
- `AI_SESSION_PROTOCOL.md` as the session operating contract

This pattern makes SmartStat sessions reconstruct the same architecture more
reliably across chats and tools.

---

## Maintenance Rules

Update these files with different frequencies:

### Rarely Update

- `SYSTEM_INDEX.md`
- `ARCHITECTURE_ANCHOR.md`

### Update When Architecture or Runtime Interpretation Changes

- `PROJECT_BRAIN.md`
- `RUNTIME_PIPELINE.md`
- `SMARTSTAT_RUNTIME_MAP.md`
- `project-context.md`
- `SMARTSTAT_AI_BOOTSTRAP.md`

### Update When Governance Changes

- `DEVELOPMENT_RULES.md`
- `AI_KERNEL.md`
- `AI_SESSION_PROTOCOL.md`
- `AGENTS_SKILL_ROUTER.md`

### Update When Startup Prompts Improve

- `context_seed.md`
- `chat-starters.md`

---

## Hard Rule

These files help AI reason correctly, but they do not authorize changes.

No runtime or config change is authorized unless it is also valid under:

- `AGENTS.md`
- `SESSION.md`
- `ROADMAP.md`
- `docs/viz-trio/`
- accepted roadmap and evidence discipline
