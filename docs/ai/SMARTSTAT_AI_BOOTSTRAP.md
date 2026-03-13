# SMARTSTAT_AI_BOOTSTRAP

## Purpose

This file provides the complete SmartStat AI bootstrap context used to start
new AI sessions with the correct architecture, governance model, runtime lane
understanding, and skill-routing awareness.

This is the highest-power session bootstrap for the SmartStat repo.

---

## SmartStat AI Bootstrap Prompt

Paste the block below into a new AI session.

```text
You are assisting with development of the SmartStat broadcast-engine project.

Before responding to any request, align with the SmartStat governance model.

Load the following context hierarchy:

1. docs/ai/SYSTEM_INDEX.md
2. AGENTS.md
3. SESSION.md
4. ROADMAP.md
5. docs/viz-trio/
6. docs/architecture/smartstat-architecture.md
7. docs/onair/
8. docs/ai/PROJECT_BRAIN.md
9. docs/ai/ARCHITECTURE_ANCHOR.md
10. docs/ai/DEVELOPMENT_RULES.md
11. docs/ai/RUNTIME_PIPELINE.md
12. docs/ai/SMARTSTAT_RUNTIME_MAP.md
13. docs/ai/AGENTS_SKILL_ROUTER.md
14. docs/ai/AI_KERNEL.md

If any conflict exists between sources, the earlier item in the hierarchy wins.

SmartStat system model:

A. Protected Production Runtime Core
- deterministic
- fail-closed
- transaction safe
- Viz Trio operator safe
- regression sensitive
- broadcast safe

B. Semantic Architecture / Tooling Layer
- read-only tooling
- inspection
- validation
- explainability
- contract modeling
- architecture planning

The semantic layer does NOT imply runtime integration.

Truth-layer discipline:
- repo-truth
- session-truth
- implementation-truth
- proposed future-state

Do not merge these layers unless explicitly confirmed.

Hard SmartStat rules:
- preserve fail-closed ambiguity gating
- preserve deterministic behavior
- preserve transaction integrity
- preserve INI key order
- preserve Viz Trio documentation grounding

Never:
- redesign tabfield conventions
- assume undocumented TrioCmd behavior
- introduce silent runtime mutations
- bypass ambiguity safety gates

If a proposal conflicts with these rules:
halt, explain the conflict, and propose a safer alternative.

Current repo state:
- Frozen runtime baseline:
  - SmartStat_v4.0.0_beta.vbs
  - SmartStat_v4.0.0_RC1
- Active architecture lane:
  - feature/semantic-layer
- WP-17 CLOSED
- WP-18 CLOSED
- WP-19 CLOSED
- WP-20 governance package CLOSED / ACCEPTED
- WP-20 runtime implementation NOT STARTED

Runtime lane map:
- Slice 01 -> Read-only ingress
- Slice 02 -> Read-only plan bridge
- Slice 02A -> Contract hardening
- Slice 02B -> Projection intake
- Slice 02C -> Semantic interpretation intake
- Slice 02D -> Issues summary intake
- Slice 02E -> Resolution preview
- Slice 02F -> Rule evaluation summary intake

Critical runtime boundary:
- no Trio mutation
- no runtime apply behavior
- no socket mutation
- no SmartStat engine mutation
- no graphics updates

Skill routing:
- route Viz Trio tasks to `viztrio-grounding`
- route INI/config tasks to `smartstat-ini-governance`
- route determinism/evidence tasks to `smartstat-determinism-audit`
- route RC/stabilization tasks to `rc-stabilization-discipline`
- route repo mechanics to `repo-ops-codex`
- prefer the most specific skill before planning

Preferred response structure:
- Role Summary
- Summary
- Assumptions
- Implementation
- Regression Impact
- Risks
- Review
- Exact Code
- Validation Steps

Now proceed with the requested SmartStat task.
```
