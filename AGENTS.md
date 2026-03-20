# AGENTS.md - SmartStat Core Governance

## Project Scope

SmartStat Core Engine (VBScript + INI system)

Includes:
- `SmartStat_v4.0.0_beta.vbs`
- all `SmartStat_Mappings*.ini` variants
- `SmartStat_StaticOverrides.ini`
- `SmartStat_TemplateConfig.ini`

Excludes:
- naming convention changes
- Viz Trio tabfield redesign
- external tool schema changes

---

# AI Kernel Initialization (Advanced Repo Pattern)

Before performing reasoning, planning, or code generation, the agent MUST align
with the SmartStat AI Kernel context system.

To minimize token usage while maximizing architectural grounding,
use the following fast-load sequence:

### Phase 1 - Kernel Index (fast architecture orientation)

Load:

- `docs/ai/SYSTEM_INDEX.md`

### Phase 2 - Governance Sources

Load only the authoritative governance files:

1. `AGENTS.md`
2. `SESSION.md`
3. `ROADMAP.md`

These override any AI context files.

### Phase 3 - Architecture Kernel

Load:

- `docs/ai/PROJECT_BRAIN.md`
- `docs/ai/ARCHITECTURE_ANCHOR.md`
- `docs/ai/DEVELOPMENT_RULES.md`
- `docs/ai/RUNTIME_PIPELINE.md`
- `docs/ai/SMARTSTAT_RUNTIME_MAP.md`
- `docs/ai/AI_KERNEL.md`

### Phase 4 - Domain Grounding

Load only if the task involves runtime behavior or Viz interaction:

- `docs/viz-trio/`
- `docs/architecture/smartstat-architecture.md`
- `docs/onair/`

### Phase 5 - Skill Router and Session Helpers

Load skill routing rules before planning:

- `docs/ai/AGENTS_SKILL_ROUTER.md`

Load prompt helpers only when needed:

- `docs/ai/context_seed.md`
- `docs/ai/SMARTSTAT_AI_BOOTSTRAP.md`
- `docs/ai/chat-starters.md`

These are prompt helpers, not governance.

### Hard Rule

AI context files improve reasoning but never override governance.

Authoritative hierarchy:

1. `AGENTS.md`
2. `SESSION.md`
3. `ROADMAP.md`
4. `docs/viz-trio/`
5. `docs/architecture/smartstat-architecture.md`
6. `docs/onair/`
7. `docs/ai/*`
8. `.agents/skills/*`

---

## AI Context System Alignment

This repo includes a repo-local AI grounding layer under `docs/ai/`.

Use it to improve session continuity and architecture accuracy,
but do not treat it as authoritative over repo governance.

If `docs/ai/*` conflicts with governance sources,
the higher-priority source wins.

---

## Viz Trio Grounding Requirement (Method A - Same Repo)

All SmartStat changes must align with `docs/viz-trio/`.

No Trio command, tabfield assumption, or operator workflow may be inferred
without documentation support.

Before proposing any SmartStat changes:

1. Read `docs/viz-trio/environment_constraints.md` (live-safe rules)
2. Read `docs/viz-trio/command_reference.md`
3. Read `docs/viz-trio/commands_full_index.md`
4. Read `docs/viz-trio/tabfields.md` (tabfield prefix heuristics)
5. Use `page_list.md`, `page_editor.md`, and `show_control.md`
   when reasoning about operator workflow

If uncertain:

- quote the relevant section
- justify the decision
- fail closed

If a proposal conflicts with those docs:

- stop
- propose a safer alternative

Use `PROMPTS.md` for standardized Codex kickoff blocks.

---

## Viewer Layer Boundary (WP-20 Enforcement)

The viewer layer is strictly:

- read-only
- deterministic
- a visualization of runtime output only

The viewer must NOT:

- implement runtime logic
- derive new data not present in runtime output
- infer relationships between entities
- perform validation or scoring
- simulate mutation or apply behavior

The viewer is:

- a truth surface
- a debug surface
- an inspection surface

The viewer is NOT:

- a reasoning engine
- a validation engine
- a runtime extension layer

Any violation:

1. Halt
2. Reject change
3. Re-scope to viewer-only

---

## Codex Agent Skills

This repo uses Codex-compatible skills stored in:

`.agents/skills/<skill-name>/SKILL.md`

Before planning, the agent should consult
`docs/ai/AGENTS_SKILL_ROUTER.md` to select the most specific matching skill.

When a task falls into one of these domains,
Codex should prefer the matching skill workflow:

- Viz Trio semantics / TrioCmd / tabfields / operator workflow
  -> `viztrio-grounding`
- INI contract/order/aliases/dup keys
  -> `smartstat-ini-governance`
- determinism evidence / run comparisons / stable hashing
  -> `smartstat-determinism-audit`
- RC-only allowed work / regression evidence requirements
  -> `rc-stabilization-discipline`
- standard repo mechanics (`-U5` diffs, line extraction)
  -> `repo-ops-codex`

---

## Release Discipline

### Beta (`v4.0.0_beta`)

- feature complete
- frozen for behavioral change

### RC (`v4.0.0_RC1`)

Allowed:

- stability validation
- log clarity
- determinism verification
- minor guardrails

Forbidden:

- resolver math changes
- transaction behavior changes
- schema changes
- feature additions

### v4.1+

Allowed:

- architectural improvements
- resolver enhancements
- performance work
- new harness capabilities

### RC Stabilization Rules

- RC work is doc/tests/log clarity only unless explicitly approved
  as a roadmap item.
- Any behavior change requires a new WP entry plus
  Phase-4 style regression evidence.
- All new harness artifacts must remain under `/tests/...`.

All changes must:

- preserve fail-closed ambiguity gating
- preserve transaction integrity
- avoid unintended behavioral drift

---

## Branch / Lane Separation (Mandatory)

- Single active lane is `feature/semantic-layer`
  for semantic architecture/tooling only.
- Do not mix runtime execution work with semantic architecture/tooling work.
- Runtime bridge/execution proposals require explicit approval
  before implementation.
- Runtime bridge implementation may require a separate dedicated branch.

If lane boundaries are unclear:

- halt
- request explicit scope confirmation

---

## Runtime Re-entry Authorization Gate

Runtime work may not resume automatically after viewer completion.

Before any runtime execution, mutation, or apply work:

The following must be explicitly defined:

1. A single bounded runtime objective
2. Allowed surfaces (what CAN change)
3. Forbidden surfaces (what MUST NOT change)
4. Determinism guarantees
5. Fail-closed behavior model
6. Regression validation requirements

Without this:

- Runtime work is NOT authorized
- Codex must halt and request scope confirmation

Current mandatory re-entry contract:

- `docs/onair/wp20_runtime_reentry_authorization_envelope_01.md`
- Status: `NOT AUTHORIZED` (runtime implementation must remain blocked unless a future superseding envelope explicitly authorizes one bounded objective).

---

## Code Delivery Rules

- Always provide unified diffs (`-U5` minimum)
- preserve INI formatting and key order
- full file required if partial patch is unsafe
- flag SmartStatTrayApp compatibility risks
- do not silently refactor unrelated code
- clearly state regression impact

---

## Codex Efficiency Profile (Enforced)

This section defines how Codex must control context, scope,
and execution cost during all SmartStat work.

All Codex work must follow the SmartStat efficiency model:

### Default Mode
- Precision Patch Mode (single-file, minimal context)

### Mode Escalation
Escalate ONLY if required:
- Read-Only Audit → for planning
- Cross-File Implementation → only after audit proves necessity
- New Slice Definition → for WP/runtime progression

### Hard Rules
- No no-op prompts
- No unnecessary repo-wide context
- No opportunistic refactors
- Preserve determinism and fail-closed behavior

### Closeout Enforcement
All approved passes must include:
- commit of approved files only
- SESSION.md update
- ROADMAP.md update (if needed)
- untracked file triage
- `_scratch` artifacts remain untracked

### Context Discipline
- Prefer tagged files or pasted code over full repo context
- Avoid IDE-wide context unless explicitly required

Violation of this section:
- halt
- rescope to minimal-context mode

---

## Conflict Policy

If a request:

- breaks determinism
- alters ambiguity gating
- reorders INI keys
- risks external compatibility
- conflicts with Viz Trio documentation

Then:

1. Halt
2. Explain conflict clearly
3. Propose a safe alternative implementation

Fail closed by default.

---

## No Derived Data Rule

AI and implementation layers must not create new data that is not explicitly present in:

- runtime output
- validated upstream artifacts (WP-17, WP-18, WP-19)

Forbidden:

- inferred relationships
- synthesized summaries not present in source
- heuristic grouping or scoring
- “helpful” explanations not backed by data

All outputs must be:

- traceable to source
- reproducible
- deterministic

Fail closed if data is incomplete.

---

## Current Release Target and Branch Strategy

- Frozen runtime/core baselines: `v4.0.0_beta` and `v4.0.0_RC1`.
- Single active development lane:
  `feature/semantic-layer` (semantic architecture/tooling plus bounded read-only runtime-slice implementation/validation/governance only).
- WP-18 CLOSED (validation layer).
- WP-19 CLOSED (viewer contracts).
- WP-20 governance package CLOSED / ACCEPTED.
- WP-20 read-only runtime slice chain is implemented and validated through `WP20_RUNTIME_SLICE_02H_READONLY_PLAN_BRIDGE_INELIGIBLE_EVIDENCE_PREVIEW`.
- Runtime continuation hardening chain `RUNTIME_CONTINUATION_PASS_01` through `RUNTIME_CONTINUATION_PASS_08` is implemented/validated/closed on existing Slice-01 through 02G surfaces.
- Current post-02H runtime-lane posture: `HOLD / GATED` (no active runtime implementation pass).
- Broader WP-20 runtime mutation/apply implementation remains NOT STARTED / NOT AUTHORIZED.

WP-20 allowed upstream inputs:

- WP-17 captured plans
- WP-18 validation outputs
- WP-19 projection contracts

WP-20 forbidden behavior before broader implementation:

- no runtime mutation/apply behavior
- no apply behavior
- no Trio integration
- no SmartStat engine calls
- no viewer implementation
- no socket mutation behavior
- no artifact mutation

Runtime baseline protection rule:

`SmartStat_v4.0.0_beta.vbs` may not be modified by WP-20 runtime work.

---

## Testing Governance

- Codex may create any harness files needed.
- All harness outputs must go under `/tests/wp-XX/target-YY/`.
- If WP/target unknown, use `/tests/_scratch/<task>/`.
- Root `tmp_*` files are forbidden.
