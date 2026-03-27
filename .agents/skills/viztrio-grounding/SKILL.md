---
name: viztrio-grounding
description: Use this skill whenever interpreting Viz Trio semantics, TrioCmd usage, tabfield behavior, custom properties, or operator workflow assumptions. Ground all claims in docs/viz-trio and fail closed when unsupported.
---

# Viz Trio Grounding

## When to use
Use this skill for any work that depends on Viz Trio behavior, including:
- TrioCmd syntax, page:get_property / page:set_property usage
- Tabfield conventions (A####, B####, C####, etc.)
- Operator workflow assumptions and constraints
- Any statement like "Viz Trio does X" or "Tabfields behave like Y"

## Grounding rule (hard)
- Always consult `docs/viz-trio/` first.
- If the required detail is not supported by docs, fail closed:
  - State what is missing
  - Propose a safe alternative that does not assume behavior

## References
- `docs/viz-trio/` is the source of truth.
- Also see `.agents/skills/viztrio-grounding/references/INDEX.md`

## Assets
- Snippets in `.agents/skills/viztrio-grounding/assets/snippets/`

## Output expectations
- Quote or cite the exact doc section paths you used (path + heading).
- If unsupported, explicitly say "Unsupported by docs — failing closed."
