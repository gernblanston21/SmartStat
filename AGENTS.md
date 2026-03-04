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

## Current Release Target

v4.0.0_RC1 frozen
Post-RC development continues on `v4_Dev` branch toward v4.1.0.

---

## Testing Governance

- Codex may create any harness files needed.
- All harnesses and outputs must go under `/tests/wp-XX/target-YY/`.
- If WP/target is unknown, use `/tests/_scratch/<task>/`.
- Root `tmp_*` files are forbidden; move them into `/tests` before final output.
