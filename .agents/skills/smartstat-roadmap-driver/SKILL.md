---
name: smartstat-roadmap-driver
description: Drive SmartStat roadmap execution using current repo governance. Reads ROADMAP.md + SESSION.md + AGENTS.md, infers the active development posture, summarizes WP status (open/closed/hold), and generates copy/paste Codex prompts + evidence templates. Fail closed if required files are missing or if requested work would violate current lane, viewer, runtime, or release constraints.
---

# SmartStat Roadmap Driver

## Purpose

Use this skill to drive SmartStat roadmap execution from the **current**
repo governance state.

This skill helps the agent:

- determine the current SmartStat roadmap posture
- identify which WP / pass / slice is active, closed, or on hold
- generate Codex-ready prompt blocks
- generate evidence / validation / checkpoint blocks
- prevent outdated roadmap assumptions from leaking into planning

This skill is **workflow support**, not governance.
If it conflicts with repo authority, the higher-priority source wins.

Authoritative order:

1. `AGENTS.md`
2. `SESSION.md`
3. `ROADMAP.md`

If those sources conflict, fail closed and report the conflict.

---

## When to Use

Use this skill when the user asks things like:

- “What’s next in the roadmap?”
- “Which WP is active right now?”
- “Generate the Codex prompt for the next approved WP / pass / slice”
- “What should go into SESSION.md for this stage?”
- “What evidence do we need before implementation?”
- “Are we still in viewer work, semantic work, RC work, or runtime hold?”
- “Summarize current roadmap status before I continue in Codex”

Do **not** use this skill to override explicit instructions from governance files.

---

## Required Inputs

Read these repo-local files first:

### Required
- `AGENTS.md`
- `ROADMAP.md`

### Preferred
- `SESSION.md`

### Optional
- `PROMPTS.md`
- `docs/ai/context_seed.md`
- `docs/ai/SMARTSTAT_AI_BOOTSTRAP.md`

If `AGENTS.md` or `ROADMAP.md` is missing, stop and fail closed.

---

## Mode Inference Rules

Determine the current operating posture in this order:

1. `AGENTS.md`
   - active branch / lane
   - release discipline
   - viewer/runtime restrictions
   - runtime hold status
   - allowed / forbidden work

2. `SESSION.md`
   - current session posture
   - active phase / WP focus
   - current branch wording
   - frozen baselines / release target

3. `ROADMAP.md`
   - official WP ordering
   - current open / closed / hold markers
   - approved next-step structure

If any of these disagree in a material way:

- halt
- state the conflict explicitly
- do not guess
- do not synthesize a “best” answer

Fail closed.

---

## Hard Rules

- Never assume older WP ranges are still the active roadmap focus.
- Never reuse stale roadmap language without checking current governance.
- Respect the current active lane exactly as defined in `AGENTS.md`.
- Respect viewer-only boundaries when viewer work is active.
- Do not assume viewer completion authorizes runtime progression.
- Runtime execution, mutation, or apply work requires explicit authorization.
- Do not invent a next runtime slice unless governance explicitly authorizes one.
- In RC mode, allow only docs/tests/log clarity work unless explicitly approved as a roadmap item.
- Any behavior change must include regression evidence requirements.
- Never infer Viz Trio behavior; defer to `viztrio-grounding` and `docs/viz-trio/`.
- Never broaden a roadmap task beyond its approved boundary.
- Do not convert a planning request into an implementation request.

---

## Viewer / Runtime Boundary Enforcement

When the repo is in viewer-layer work:

- viewer changes must remain read-only
- viewer changes must remain deterministic
- viewer must only visualize runtime output
- viewer must not derive new logic
- viewer must not become a validation engine
- viewer must not simulate mutation/apply behavior

When the repo is in runtime hold posture:

- do not authorize runtime work automatically
- do not treat viewer completion as runtime unlock
- do not generate implementation prompts for runtime execution work
  unless explicit authorization exists

---

## Output Types This Skill May Produce

This skill may produce:

- current roadmap status summaries
- WP / pass / slice status breakdowns
- “what’s next” planning summaries
- Codex-ready prompt blocks
- validation / evidence requirement checklists
- SESSION.md update suggestions
- roadmap checkpoint / freeze / hold summaries

This skill must **not** produce:

- code patches
- runtime logic
- viewer logic
- inferred approvals
- speculative future roadmap branches

---

## Expected Output Pattern

When responding, prefer this structure:

### 1. Current Posture
- active lane
- active WP / pass / slice
- closed items
- hold items
- forbidden work

### 2. Safe Next Action
- what may happen next
- what may not happen next
- whether planning only or implementation is allowed

### 3. Evidence Requirements
- validation needed before implementation
- regression requirements
- determinism requirements
- fail-closed requirements

### 4. Codex Block
- a copy/paste prompt if appropriate
- bounded to current governance

---

## Failure Conditions

Fail closed if any of the following are true:

- `AGENTS.md` missing
- `ROADMAP.md` missing
- current lane unclear
- runtime/viewer boundary unclear
- user request conflicts with current release discipline
- roadmap status is ambiguous
- requested next step is not explicitly authorized
- session wording conflicts with roadmap/governance posture

In failure cases:

1. state exactly what is unclear
2. identify which file(s) must be reconciled
3. do not guess a next WP / pass / slice

---

## Script Hooks

If helper scripts exist, they may be used for support only.

Examples:

- `scripts/roadmap_status.ps1`
  - summarize current posture and WP state

- `scripts/roadmap_next_steps.ps1`
  - generate bounded next-step planning options from current governance

- `scripts/wp_prompt.ps1`
  - generate a Codex-ready prompt block for a specific approved WP / pass / slice

- `scripts/rc_guard.ps1`
  - enforce RC-safe work boundaries

If script output conflicts with governance files, governance wins.

---

## Recommended Companion Skills

Use with:

- `smartstat-core-engine-workflow`
  - for bounded implementation planning once roadmap posture is confirmed

- `smartstat-determinism-audit`
  - for regression / determinism evidence requirements

- `rc-stabilization-discipline`
  - when repo posture is RC-bound

- `viztrio-grounding`
  - when any roadmap work touches Trio semantics or operator workflow

- `repo-ops-codex`
  - for exact repo mechanics, diffs, file extraction, and line-context tasks

---

## Example Safe Uses

### Example 1
User asks:
> What is the next approved step after viewer PASS_16?

Safe response pattern:
- read `AGENTS.md`, `SESSION.md`, `ROADMAP.md`
- confirm viewer boundary still active
- confirm remaining viewer passes
- do not authorize runtime automatically

### Example 2
User asks:
> Generate the Codex prompt for the next runtime step.

Safe response pattern:
- first verify whether runtime work is authorized
- if runtime is on hold, provide planning-only prompt
- do not generate implementation prompt unless explicitly allowed

### Example 3
User asks:
> Summarize where SmartStat stands right now.

Safe response pattern:
- identify active lane
- identify frozen baselines
- identify closed WPs
- identify current hold posture
- identify allowed next planning action only

---

## Final Rule

This skill exists to keep roadmap work aligned with the **current**
SmartStat governance state.

It must never:

- revive outdated roadmap assumptions
- skip authorization boundaries
- convert hold state into implementation authority
- allow viewer work to leak into runtime work

When uncertain, fail closed.
