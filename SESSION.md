# SESSION — SmartStat Core Engine

## Session Identity (must be set at start)
- Working target version (user-declared): v4.0.0_beta
- SmartStat baseline identifier (user-declared): v4.0.0_beta
- Workspace folder: SMARTSTAT (this VS Code workspace)

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

## Non-negotiables (session enforcement)
- Version numbers MUST be specified by the user at the beginning of each session.
  - For SmartStat: use the “user-declared” version above (NOT VERSION.txt).
- Do not change naming conventions.
- Do not redesign Viz Trio tabfield patterns.
- Preserve INI formatting, spacing, and key order within sections.
- No placeholders.
- If a patch is unsafe as a partial snippet, output the full file.
- Flag breaking changes that may affect external tools (SmartStatTrayApp).

## Current status snapshot
- Status: PAUSED at ambiguity-system implementation checkpoint.
- Last confirmed analysis finding:
  - AMBIGUITY output is missing because:
    - There is no Ambiguity_Add function in the script, and
    - There is no initialization for CompilerContext("ambiguous"),
    - A shorthand guard calls a non-existent routine,
    - The call fails silently due to Main() running under On Error Resume Next,
    - Execution falls through to the unresolved path.
- No changes have been applied after this analysis point.

## Next action (ready to proceed)
- [ ] Proceed with Roadmap “In Progress” Item 1: Ambiguity system implementation.
