# SmartStat AI Kernel

## Purpose

This file defines the SmartStat AI Kernel pattern.

The AI Kernel is the repo-local operating model that makes new AI sessions
behave consistently across chats by forcing them to reconstruct the same
architecture, rule system, runtime map, and truth-layer discipline every time.

Think of it as the SmartStat AI operating environment.

---

## Kernel Components

The SmartStat AI Kernel is composed of these files:

1. `docs/ai/SYSTEM_INDEX.md`
2. `docs/ai/PROJECT_BRAIN.md`
3. `docs/ai/ARCHITECTURE_ANCHOR.md`
4. `docs/ai/DEVELOPMENT_RULES.md`
5. `docs/ai/RUNTIME_PIPELINE.md`
6. `docs/ai/SMARTSTAT_RUNTIME_MAP.md`
7. `docs/ai/context_seed.md`
8. `docs/ai/SMARTSTAT_AI_BOOTSTRAP.md`
9. `docs/ai/AGENTS_SKILL_ROUTER.md`
10. `docs/ai/chat-starters.md`
11. `docs/ai/AI_SESSION_PROTOCOL.md`

These sit under the governance of:

- `AGENTS.md`
- `SESSION.md`
- `ROADMAP.md`
- `docs/viz-trio/`

---

## Kernel Startup Model

A strong SmartStat AI session should reconstruct context in this order:

### Phase 1 - Authority
- `SYSTEM_INDEX.md`
- `AGENTS.md`
- `SESSION.md`
- `ROADMAP.md`

### Phase 2 - Project Model
- `PROJECT_BRAIN.md`
- `ARCHITECTURE_ANCHOR.md`
- `DEVELOPMENT_RULES.md`

### Phase 3 - Runtime / Lane Interpretation
- `RUNTIME_PIPELINE.md`
- `SMARTSTAT_RUNTIME_MAP.md`

### Phase 4 - Session Boot
- `AGENTS_SKILL_ROUTER.md`
- `context_seed.md`
- `SMARTSTAT_AI_BOOTSTRAP.md`
- `chat-starters.md`
- `AI_SESSION_PROTOCOL.md`

---

## Kernel Guarantees

When used correctly, the AI Kernel should improve:

- architectural consistency across chats
- runtime vs semantic-lane separation
- SmartStat governance compliance
- Viz Trio doc-grounded reasoning
- prompt reuse quality
- Codex session startup reliability
- skill selection before planning

---

## Kernel Rules

### 1. Governance Wins
If the kernel conflicts with `AGENTS.md`, `SESSION.md`, or `ROADMAP.md`, those
repo governance files win.

### 2. Viz Trio Docs Win on Trio Behavior
No Trio behavior may be inferred without `docs/viz-trio/` support.

### 3. Truth Layers Must Stay Separate
Repo-truth, session-truth, implementation-truth, and proposed future-state must
not be merged casually.

### 4. Read-Only Lanes Do Not Authorize Runtime Changes
A docs/tests/tooling package is not runtime implementation.

### 5. Bootstrap Prompts Are Session Tools, Not Governance
Prompt files guide startup behavior but do not authorize code changes.

### 6. Skill Router Runs Before Planning
The correct repo skill should be identified before building plans, patches, or
recommendations whenever a matching domain skill exists.

---

## Recommended Chat Entry

For a new SmartStat chat, begin with:

```text
This chat is part of the SmartStat Development project.

Before answering, load the context defined in:
- docs/ai/SYSTEM_INDEX.md
- docs/ai/context_seed.md
```

For a deeper chat, add:

```text
Also align with:
- AGENTS.md
- SESSION.md
- ROADMAP.md
- docs/ai/SMARTSTAT_AI_BOOTSTRAP.md
```

---

## Recommended Codex Entry

For a new Codex run, begin with:

```text
You are working in the SmartStat repo.

Load and align with:
- docs/ai/SYSTEM_INDEX.md
- AGENTS.md
- SESSION.md
- ROADMAP.md
- docs/ai/PROJECT_BRAIN.md
- docs/ai/ARCHITECTURE_ANCHOR.md
- docs/ai/DEVELOPMENT_RULES.md
- docs/ai/RUNTIME_PIPELINE.md
- docs/ai/SMARTSTAT_RUNTIME_MAP.md
- docs/ai/AGENTS_SKILL_ROUTER.md
- docs/ai/AI_KERNEL.md
- docs/viz-trio/
```

---

## Maintenance

Update the kernel components when:

- architecture changes
- roadmap interpretation changes
- runtime bridge sequencing changes
- authority order changes
- startup prompt strategy changes
- skills inventory changes

The kernel should evolve deliberately, not casually.
