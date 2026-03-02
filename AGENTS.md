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

## Viz Trio Grounding Requirement
- All SmartStat changes must align with `docs/viz-trio/`; no Trio command, tabfield assumption, or operator workflow may be inferred without documentation support. If uncertain, quote the relevant section and fail closed.


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
- Architectural improvements
- Resolver enhancements
- Performance work
- New harness capabilities

---

## Code Delivery Rules
- Always provide unified diffs (-U5 minimum)
- Preserve INI formatting and key order
- Full file required if partial patch is unsafe
- Flag SmartStatTrayApp compatibility risks

---

## Conflict Policy
If a request:
- Breaks determinism
- Alters ambiguity gating
- Reorders INI keys
- Risks external compatibility

Then:
1. Halt
2. Explain conflict
3. Propose safe alternative

---

## Current Release Target
Preparing v4.0.0_RC1 as of 2026-02-28.
