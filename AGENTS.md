# AGENTS.md - SmartStat Core Governance

## Project Scope
SmartStat Core Engine (VBScript + INI system)

Includes:
- SmartStat_v4.0.0_beta.vbs
- All SmartStat_Mappings*.ini variants
- SmartStat_StaticOverrides.ini
- SmartStat_TemplateConfig.ini

Excludes:
- Naming convention changes
- Viz Trio tabfield redesign
- External tool schema changes

---

## Viz Trio Grounding Requirement (Method A — Same Repo)

All SmartStat changes must align with `docs/viz-trio/`.

No Trio command, tabfield assumption, or operator workflow may be inferred without documentation support.

Before proposing ANY SmartStat changes:
1) Read `docs/viz-trio/environment_constraints.md` (live-safe rules)
2) Read `docs/viz-trio/command_reference.md`
3) Read `docs/viz-trio/commands_full_index.md`
4) Read `docs/viz-trio/tabfields.md` (tabfield prefix heuristics)
5) Use `page_list.md`, `page_editor.md`, and `show_control.md` when reasoning about operator workflow

If uncertain:
- Quote the relevant section.
- Justify the decision.
- Fail closed.

If a proposal conflicts with those docs:
- STOP.
- Propose a safer alternative.

Use `PROMPTS.md` for standardized Codex kickoff blocks.

---

## Codex Agent Skills

This repo uses Codex-compatible skills stored in:
- `.agents/skills/<skill-name>/SKILL.md`

When a task falls into one of these domains, Codex must prefer the matching skill workflow:
- Viz Trio semantics / TrioCmd / tabfields / operator workflow → `viztrio-grounding`
- INI contract/order/aliases/dup keys → `smartstat-ini-governance`
- Determinism evidence / run comparisons / stable hashing → `smartstat-determinism-audit`
- RC-only allowed work / regression evidence requirements → `rc-stabilization-discipline`
- Standard repo mechanics (-U5 diffs, line extraction) → `repo-ops-codex`

---

## Release Discipline

### Beta (v4.0.0_beta)
- Feature complete
- Frozen for behavioral change

### RC (v4.0.0_RC1)
Allowed:
- Stability validation
- Log clarity
- Determinism verification
- Minor guardrails

Forbidden:
- Resolver math changes
- Transaction behavior changes
- Schema changes
- Feature additions

### v4.1+
Allowed:
- Architectural improvements
- Resolver enhancements
- Performance work
- New harness capabilities

### RC Stabilization Rules
- RC work is doc/tests/log clarity only unless explicitly approved as a roadmap item.
- Any behavior change requires a new WP entry + Phase-4 style regression evidence.
- All new harness artifacts must remain under `/tests/...`.

All changes must:
- Preserve fail-closed ambiguity gating
- Preserve transaction integrity
- Avoid unintended behavioral drift

### Branch / Lane Separation (Mandatory)
- Single active lane is `feature/semantic-layer` for semantic architecture/tooling only.
- Do not mix runtime execution work with semantic architecture/tooling work in one pass.
- Runtime bridge/execution proposals require explicit approval before implementation.
- Runtime bridge/execution implementation may require a separate dedicated branch.
- If lane boundaries are unclear, halt and request explicit scope confirmation.

---

## Code Delivery Rules

- Always provide unified diffs (-U5 minimum)
- Preserve INI formatting and key order
- Full file required if partial patch is unsafe
- Flag SmartStatTrayApp compatibility risks
- Do not silently refactor unrelated code
- Clearly state regression impact

---

## Conflict Policy

If a request:
- Breaks determinism
- Alters ambiguity gating
- Reorders INI keys
- Risks external compatibility
- Conflicts with Viz Trio documentation

Then:
1. Halt
2. Explain conflict clearly
3. Propose safe alternative implementation

Fail closed by default.

---

## Current Release Target and Branch Strategy

- Frozen runtime/core baselines: `v4.0.0_beta` and `v4.0.0_RC1`.
- Single active development lane: `feature/semantic-layer` (semantic architecture/tooling only).
- WP-18 is CLOSED as a validation-layer-only package and must remain runtime-independent.
- WP-19 is CLOSED as a read-only viewer-contract package over WP-17/WP-18 artifacts.
- WP-19 contract surfaces may be consumed read-only only: projection contract, projection summary/consumption surfaces, and projection-to-view-model adapter contract.
- WP-20 kickoff gate, Target-01 approval package, Target-02 charter/plan gate, Target-03 rehearsal gate, Target-04 sign-off gate, Target-05 implementation-authorization decision gate, and Target-06 authorization-packet fill/verification gate are defined (governance/docs/tests only) and WP-20 remains NOT STARTED.
- WP-20 allowed upstream inputs are limited to: WP-17 captured-plan artifacts, WP-18 validation outputs (`validation_result`, `rule_evaluations`, refusal diagnostics, deterministic identities, `semantic_interpretation`), and WP-19 projection/adapter contract surfaces.
- WP-20 forbidden pre-implementation behavior: no runtime bridge code, no apply behavior, no Trio integration, no SmartStat engine/apply calls, no viewer implementation, and no artifact mutation.
- WP-20 implementation requires explicit approval, dedicated branch/lane separation, and an approved runtime-bridge regression/evidence plan before code changes begin.
- WP-20 kickoff checklist reference: `docs/onair/wp20_kickoff_checklist.md`.
- WP-20 approval requirements reference: `docs/onair/wp20_approval_requirements.md`.
- WP-20 lane charter reference: `docs/onair/wp20_lane_charter.md`.
- WP-20 regression/evidence plan reference: `docs/onair/wp20_regression_evidence_plan.md`.
- WP-20 rehearsal protocol reference: `docs/onair/wp20_rehearsal_protocol.md`.
- WP-20 rehearsal manifest template reference: `docs/onair/wp20_rehearsal_manifest_template.md`.
- WP-20 gate review checklist reference: `docs/onair/wp20_gate_review_checklist.md`.
- WP-20 implementation-authorization record reference: `docs/onair/wp20_implementation_authorization_record.md`.
- WP-20 branch-approval record reference: `docs/onair/wp20_branch_approval_record.md`.
- WP-20 authorization packet index template reference: `docs/onair/wp20_authorization_packet_index_template.md`.
- WP-20 packet completeness checklist reference: `docs/onair/wp20_packet_completeness_checklist.md`.
- WP-20 approval evidence template reference: `tests/wp-20/target-01/approval_evidence_template.md`.
- Runtime bridge/execution work is a distinct risk class and may require a separate dedicated branch.
- `v4_Dev` remains historical RC lineage baseline, not the active semantic feature lane.

---

## Testing Governance

- Codex may create any harness files needed.
- All harnesses and outputs must go under `/tests/wp-XX/target-YY/`.
- If WP/target is unknown, use `/tests/_scratch/<task>/`.
- Root `tmp_*` files are forbidden; move them into `/tests` before final output.
