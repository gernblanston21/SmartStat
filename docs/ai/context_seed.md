# SmartStat Context Seed

Use this file as the default fast-start seed for a new SmartStat chat.

```text
This chat is part of my SmartStat Development project.

Before answering, load the SmartStat AI context system in this order:

1. docs/ai/SYSTEM_INDEX.md
2. AGENTS.md
3. SESSION.md
4. ROADMAP.md
5. docs/ai/PROJECT_BRAIN.md
6. docs/ai/ARCHITECTURE_ANCHOR.md
7. docs/ai/DEVELOPMENT_RULES.md
8. docs/ai/RUNTIME_PIPELINE.md
9. docs/ai/SMARTSTAT_RUNTIME_MAP.md
10. docs/viz-trio/
11. docs/ai/SMARTSTAT_AI_BOOTSTRAP.md

SmartStat grounding model:
- SmartStat is a deterministic broadcast-engine for Viz Trio.
- Protected runtime/core baseline: SmartStat_v4.0.0_beta.vbs
- Primary config family:
  - SmartStat_Mappings.ini
  - SmartStat_Mappings.learn.ini
  - SmartStat_MappingsNBA.ini
  - SmartStat_MappingsNBA.learn.ini
  - SmartStat_MappingsNHL.ini
  - SmartStat_MappingsNHL.learn.ini
  - SmartStat_StaticOverrides.ini
  - SmartStat_TemplateConfig.ini
- Hard rules:
  - preserve fail-closed ambiguity gating
  - preserve transaction integrity
  - preserve INI ordering and contract discipline
  - do not infer Viz Trio behavior without docs/viz-trio grounding
  - flag SmartStatTrayApp and external compatibility risks
- Architecture model:
  - protected runtime core
  - separate semantic architecture/tooling lane
- Truth model:
  - keep repo-truth, session-truth, implementation-truth, and proposed future-state separate unless explicitly aligned

Task:
<replace with the current task>
```
