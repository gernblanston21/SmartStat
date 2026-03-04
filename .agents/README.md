# Codex Agent Skills (SmartStat)

This repository uses Codex-compatible skills stored in:

- `.agents/skills/<skill-name>/SKILL.md`

Each skill bundles:
- A narrow domain responsibility (so Codex can trigger reliably)
- Optional scripts (PowerShell-first) Codex can execute
- Optional references/assets to ground behavior

## How to use
In Codex chat, ask for tasks in natural language like:
- "Validate TemplateConfig INI ordering"
- "Run determinism smoke test on latest OperatorDiag"
- "Make a -U5 diff for this change"
- "Show me lines 820-900 from SmartStat_v4.0.0_beta.vbs"

Codex should select the matching skill and follow its workflow.

## Repo grounding sources
- Viz Trio grounding: `docs/viz-trio/`
- RC discipline: `governance/RC_POLICY_v4.md` and `governance/releases/`
- Determinism evidence: `tests/wp-10/` and `DiagLogs/`
