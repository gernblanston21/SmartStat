# SmartStat Chat Starters

## Purpose

These starters are copy/paste boot prompts for ChatGPT Project chats and Codex
repo sessions.

Use the fast starter for small tasks.
Use the deep starter for architecture, runtime, governance, or roadmap work.

---

## New ChatGPT Project Chat - Fast Starter

```text
This chat is part of my SmartStat Development project.

Before answering, load:
- docs/ai/SYSTEM_INDEX.md
- docs/ai/context_seed.md

Task:
<replace this line>
```

---

## New ChatGPT Project Chat - Deep Starter

```text
This chat is part of my SmartStat Development project.

Load and align with:
- docs/ai/SYSTEM_INDEX.md
- AGENTS.md
- SESSION.md
- ROADMAP.md
- docs/ai/project-context.md
- docs/ai/PROJECT_BRAIN.md
- docs/ai/ARCHITECTURE_ANCHOR.md
- docs/ai/DEVELOPMENT_RULES.md
- docs/ai/RUNTIME_PIPELINE.md
- docs/ai/SMARTSTAT_RUNTIME_MAP.md
- docs/ai/AGENTS_SKILL_ROUTER.md
- docs/ai/SMARTSTAT_AI_BOOTSTRAP.md
- docs/viz-trio/

Hard rules:
- preserve fail-closed behavior
- preserve transaction integrity
- do not infer Viz Trio behavior without docs/viz-trio grounding
- distinguish repo-truth, session-truth, implementation-truth, and proposed future-state

Task:
<replace this line>
```

---

## VS Code + Codex - Standard Starter

```text
You are working in the SmartStat repo.

Ground on:
- docs/ai/SYSTEM_INDEX.md
- AGENTS.md
- SESSION.md
- ROADMAP.md
- docs/ai/PROJECT_BRAIN.md
- docs/ai/DEVELOPMENT_RULES.md
- docs/ai/ARCHITECTURE_ANCHOR.md
- docs/ai/RUNTIME_PIPELINE.md
- docs/ai/SMARTSTAT_RUNTIME_MAP.md
- docs/ai/AGENTS_SKILL_ROUTER.md
- docs/ai/SMARTSTAT_AI_BOOTSTRAP.md
- docs/viz-trio/

Use matching repo skills under `.agents/skills/` when applicable.

Task:
<replace this line>

Deliver:
- Role Summary
- Summary
- Assumptions
- Implementation
- Regression Impact
- Risks
- Review
- Exact Code
- Validation Steps
```

---

## VS Code + Codex - Docs / Contract Work

```text
You are working in the SmartStat repo on docs, contracts, semantic architecture,
or governance packaging.

Ground on:
- docs/ai/SYSTEM_INDEX.md
- AGENTS.md
- SESSION.md
- ROADMAP.md
- docs/ai/project-context.md
- docs/ai/RUNTIME_PIPELINE.md
- docs/ai/SMARTSTAT_RUNTIME_MAP.md
- docs/architecture/smartstat-architecture.md
- docs/onair/
- docs/viz-trio/

Rules:
- do not imply modeled work is already runtime behavior
- distinguish repo-truth vs session-truth
- preserve terminology stability
- do not broaden scope

Task:
<replace this line>

Deliver:
- Role Summary
- Summary
- Assumptions
- Proposed doc text or structural recommendation
- Risks
- Validation notes
```
