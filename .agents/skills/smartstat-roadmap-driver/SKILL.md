---
name: smartstat-roadmap-driver
description: Drive SmartStat roadmap execution (WP-10 to WP-14) under RC discipline. Reads ROADMAP.md + SESSION.md + AGENTS.md, infers current mode (RC vs DEV), summarizes WP status (open/closed), and generates copy/paste Codex prompts + evidence templates. Fail closed if required files are missing or RC policy would be violated.
---

# SmartStat Roadmap Driver

## When to use
- "What’s next in the roadmap?"
- "Generate the Codex prompt for WP-11 / WP-12 / WP-13 / WP-14"
- "Are we in RC mode and what is allowed?"
- "Produce a PR evidence block for a WP target"

## Inputs (repo-local)
Reads:
- `ROADMAP.md` (required)
- `AGENTS.md` (required)
- `SESSION.md` (optional but preferred)
- Optional: `PROMPTS.md`

## Mode inference rules (priority)
1) If SESSION.md indicates RC (Release Candidate / RC1 / v4_RC) → RC
2) Else if ROADMAP.md Current State mentions RC1 / _RC / stabilization active → RC
3) Else → DEV

## Hard rules
- In RC mode, allow only docs/tests/log clarity work unless explicitly approved as a roadmap item.
- Any behavior change must include regression evidence under `tests/...`.
- Never invent Viz Trio behavior; use `viztrio-grounding` and `docs/viz-trio/`.

## Scripts
- `scripts/roadmap_status.ps1`: Mode + WP table (open/closed)
- `scripts/roadmap_next_steps.ps1`: Generates next-step suggestions based on mode
- `scripts/wp_prompt.ps1`: Builds a Codex-ready prompt block for a WP (+ optional target)
- `scripts/rc_guard.ps1`: Fail-closed gate for RC-allowed vs behavior-affecting requests
