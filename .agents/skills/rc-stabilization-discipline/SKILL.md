---
name: rc-stabilization-discipline
description: Use this skill when working in RC mode. Enforce RC policy: doc/tests/log clarity only unless explicitly approved. Any behavior change requires new WP entry + Phase-4 style regression evidence.
---

# RC Stabilization Discipline

## When to use
- Any work mentioning RC, release candidate, stabilization, regression evidence
- Any change request that might alter behavior during RC

## Hard rules
- RC work is doc/tests/log clarity only unless explicitly approved as a roadmap item.
- Any behavior change requires:
  - A new WP entry
  - Phase-4 style regression evidence under `tests/...`
- All new harness artifacts must remain under `tests/...`

## References (repo-local)
See `.agents/skills/rc-stabilization-discipline/references/INDEX.md`
