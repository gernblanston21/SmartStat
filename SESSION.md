# SESSION — SmartStat Core Engine

## Session Identity (must be set at start)
- Working target version (user-declared): v4.0.0_beta
- SmartStat baseline identifier (user-declared): v4.0.0_beta
- Workspace folder: SMARTSTAT (this VS Code workspace)
- Branch: v4_Dev
- Source of truth: current workspace (v4_Dev branch). For cross-checking outside Codex, use the raw GitHub link for SmartStat_v4.0.0_beta.vbs.

## Notes on VERSION.txt (important)
- VERSION.txt value: 4.0.0
- Purpose: Used by VIZOR to display the repo version in the VIZOR UI.
- SmartStat usage: NOT USED by SmartStat core engine. Do not treat VERSION.txt as SmartStat baseline gating.

## Baseline artifacts (authoritative for this session)
- Script (core engine):
  - SmartStat_v4.0.0_beta.vbs
- Config (INI):
  - SmartStat_Mappings.ini
  - SmartStat_StaticOverrides.ini
  - SmartStat_TemplateConfig.ini
- Learn / alias layers:
  - SmartStat_Mappings.learn.ini
  - SmartStat_MappingsNBA.learn.ini
  - SmartStat_MappingsNHL.learn.ini
- Sport-specific mappings:
  - SmartStat_MappingsNBA.ini
  - SmartStat_MappingsNHL.ini
- Tools:
  - SmartStatValidator.exe
- Logs:
  - DiagLogs\
  - DiagLogs\Harness\

## Non-negotiables (session enforcement)
- Version numbers MUST be specified by the user at the beginning of each session.
  - For SmartStat: use the “user-declared” version above (NOT VERSION.txt).
- Do not change naming conventions.
- Do not redesign Viz Trio tabfield patterns.
- Preserve INI formatting, spacing, and key order within sections.
- No placeholders.
- If a patch is unsafe as a partial snippet, output the full file.
- Flag breaking changes that may affect external tools (SmartStatTrayApp).

## Codex execution policy (current)
- Codex MAY use shell commands for READ-ONLY inspection (rg/grep, cat/type, git show/diff, etc.) without asking.
- Codex MUST ask for approval before any change that writes/modifies repo files (apply/edit/patch).

## Roadmap execution order (current)
1) WP-05 Output Map Coverage ? (Completed)
2) WP-04 Ambiguity transparency (fail-closed stays strict) ? (Completed)
3) WP-01 Phase hardening ? (Completed)
4) WP-07 Overrides audit trail ? (Completed)
5) WP-08 Harness expansion ? (Completed)

## Completed in this session
- [x] WP-05 Output Map Coverage
  - Runtime inference of missing/partial `output_map`
  - Deterministic prefix + hundred-group pairing
  - Explicit map entries preserved
  - Hard fail `OUTMAP.EMPTY` when no usable targets exist
  - No INI schema changes
  - No ambiguity gating changes
- [x] WP-01 Phase hardening
  - Added phase helpers for begin/end/fail/early-exit logging
  - Converted silent early exits in key pipeline spots to logged early-exits
  - Ensured finalize/refresh behavior is preserved on fail paths
  - No resolver/mapping behavior changes
- [x] WP-07 Overrides audit trail
  - Added override audit logging (old -> new) with section + entity context
  - Added explicit “skip” log when overrides INI cannot load
  - No resolver/mapping behavior changes
- [x] WP-08 Harness expansion
  - Added HARNESS_CAPTURE_ONLY mode
  - Added harness snapshot + grouped diff artifacts (deterministic ordering)
  - Added harness artifacts on early exits / config failures
  - Centralized artifact emissions (avoid duplicates)
  - No resolver/mapping/Tx behavior changes
- [x] WP-04 Ambiguity transparency
  - All ambiguity-block abort paths emit exactly one EARLY EXIT marker:
    - TX: EARLY EXIT - AMBIGUOUS_GATE
    - TX: EARLY EXIT - AMBIGUOUS_CONTEXT_INVALID
  - No remaining unbounded ambiguity dump paths
  - Diag_WriteAmbiguitySummary() gated behind learn.allow_ambiguous_apply=True
  - BuildAmbiguityOperatorSection() gated behind learn.allow_ambiguous_apply=True
  - Bounded detail enforced (max 5 entries + truncation marker)
  - Fail-closed behavior unchanged
  - No resolver/mapping/apply behavior changes

## Next action (ready to proceed)
- [ ] WP-06 Resolver stability guardrails
  - Audit candidate sorting determinism
  - Verify tie-breaking stability
  - Add logging-only transparency if nondeterminism risk exists
  - Do NOT change scoring math, thresholds, ordering logic, ApplyPlan behavior, transaction behavior, phase order, or INI schemas

## Output Requirements
- Provide unified diff (-U5) for SESSION.md only.
- Minimal edits only.
- No commentary inside the diff.
- Stop after the diff and wait for approval.
