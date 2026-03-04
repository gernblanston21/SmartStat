---
name: smartstat-core-engine-workflow
description: Use this skill to guide SmartStat core engine changes: enforce fail-closed ambiguity gating, transaction integrity, determinism evidence, and proper regression artifacts. This is workflow/policy guidance, not a place to invent Viz Trio behavior.
---

# SmartStat Core Engine Workflow

## When to use
- Any requested change to `SmartStat_v4.0.0_beta.vbs`
- Any WP task that touches ambiguity gating, determinism, or output mapping behavior

## Hard constraints (summary)
- Preserve fail-closed ambiguity gating.
- Preserve transaction integrity (no partial writes on fail paths).
- Any behavior change must include regression evidence under `tests/...`.
- Ground Viz Trio assumptions via `viztrio-grounding` skill and `docs/viz-trio/`.

## References (repo-local)
See `.agents/skills/smartstat-core-engine-workflow/references/INDEX.md`.
