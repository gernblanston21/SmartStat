# SmartStat Repomix Context Instructions

You are reviewing a curated SmartStat repository slice.

Core rules:
- Treat SmartStat as a deterministic broadcast-engine project, not a generic scripting repo.
- Preserve fail-closed behavior.
- Do not broaden scope beyond the files present.
- Respect AGENTS.md and SESSION.md as governing documents.
- Do not assume naming convention changes are allowed.
- Do not redesign Viz Trio tabfield conventions.
- Flag any external compatibility risk, especially SmartStatTrayApp or related tools.
- Prefer exact drop-in edits over abstract refactors.
- Preserve INI ordering and formatting.
- When proposing code changes, reference exact file names and exact placement.
- Prioritize:
  1. determinism
  2. regression safety
  3. operator workflow safety
  4. config compatibility
  5. maintainability

Review emphasis:
- SmartStat_v4.0.0_beta.vbs core logic
- cross-file coupling among mappings / overrides / template config
- docs/viz-trio grounding when Trio behavior is relevant
- tests or scratch artifacts only as supporting evidence, not as permission to broaden production behavior
