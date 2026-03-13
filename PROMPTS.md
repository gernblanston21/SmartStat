# PROMPTS.md — SmartStat / Codex Prompt Library

This file is a copy/paste library for consistent Codex runs.
It complements (does not replace) `AGENTS.md` governance.

---

## P0 — Universal Codex Kickoff (Viz Trio doc-gated)

You MUST treat `docs/viz-trio/` as the source of truth for Viz Trio behavior.

Before proposing ANY SmartStat changes:
1. Read `docs/viz-trio/environment_constraints.md` (live-safe rules)
2. Read `docs/viz-trio/command_reference.md` and `docs/viz-trio/commands_full_index.md`
3. Read `docs/viz-trio/tabfields.md`
4. Use `docs/viz-trio/page_list.md`, `page_editor.md`, and `show_control.md` when reasoning about operator workflow

If any proposed change conflicts with those docs, STOP and propose a safer alternative.
When unsure, QUOTE the relevant doc section and justify the decision.
Fail closed.

Now proceed with the task below.

---

## P5 — SmartStat Context Boot (Full Repo-Grounded)

You are working in the SmartStat repo.

Before proposing changes or analysis, align with these sources in order:

1. `docs/ai/SYSTEM_INDEX.md`
2. `AGENTS.md`
3. `SESSION.md`
4. `ROADMAP.md`
5. `docs/ai/PROJECT_BRAIN.md`
6. `docs/ai/ARCHITECTURE_ANCHOR.md`
7. `docs/ai/DEVELOPMENT_RULES.md`
8. `docs/ai/RUNTIME_PIPELINE.md`
9. `docs/ai/SMARTSTAT_RUNTIME_MAP.md`
10. `docs/ai/SMARTSTAT_AI_BOOTSTRAP.md`
11. `docs/viz-trio/`

Rules:
- Treat `AGENTS.md` as authoritative
- Treat `docs/viz-trio/` as the source of truth for Viz Trio behavior
- Preserve fail-closed behavior
- Preserve transaction integrity
- Preserve INI governance
- Distinguish repo-truth, session-truth, and implementation-truth
- Flag SmartStatTrayApp and external compatibility risks
- Do not broaden scope from scratch/test/doc artifacts

If repo sources conflict, say so explicitly instead of merging them silently.

Now proceed with the task below.

TASK:
[PASTE TASK HERE]

---

## P6 — Docs / Semantic Lane / Contract Work

You are working on SmartStat docs, contracts, semantic architecture, or planning.

Ground on:
- `docs/ai/SYSTEM_INDEX.md`
- `AGENTS.md`
- `SESSION.md`
- `ROADMAP.md`
- `docs/architecture/smartstat-architecture.md`
- `docs/onair/`
- `docs/ai/PROJECT_BRAIN.md`
- `docs/ai/RUNTIME_PIPELINE.md`
- `docs/ai/SMARTSTAT_RUNTIME_MAP.md`

Rules:
- Do not imply modeled or planned lanes are already live runtime behavior
- Distinguish repo-truth vs session-truth
- Preserve terminology stability
- Do not broaden scope

Deliver:
- Summary
- Assumptions
- Exact proposed doc text or structural recommendation
- Risks
- Validation notes

TASK:
[PASTE TASK HERE]

---

## P7 — INI Governance / Mapping / Config Work

You are working on SmartStat INI-governed config.

Ground on:
- `docs/ai/SYSTEM_INDEX.md`
- `AGENTS.md`
- `docs/ai/DEVELOPMENT_RULES.md`
- `SmartStat_Mappings.ini`
- `SmartStat_Mappings.learn.ini`
- `SmartStat_MappingsNBA.ini`
- `SmartStat_MappingsNBA.learn.ini`
- `SmartStat_MappingsNHL.ini`
- `SmartStat_MappingsNHL.learn.ini`
- `SmartStat_StaticOverrides.ini`
- `SmartStat_TemplateConfig.ini`

Rules:
- Preserve key order
- Preserve formatting and spacing
- Do not silently normalize beyond approved rules
- Flag cross-file compatibility risk
- Flag SmartStatTrayApp compatibility risk if config meaning changes

Deliver:
- Exact impacted file(s)
- Ordering/contract impact
- Risks
- Exact proposed edits
- Validation steps

TASK:
[PASTE TASK HERE]

---

## P8 — ChatGPT Project Continuity Export

Prepare a ChatGPT-ready continuity packet for SmartStat work.

Ground on:
- `docs/ai/SYSTEM_INDEX.md`
- `AGENTS.md`
- `SESSION.md`
- `ROADMAP.md`
- `docs/ai/project-context.md`
- `docs/ai/context_seed.md`
- `docs/ai/PROJECT_BRAIN.md`
- `docs/ai/ARCHITECTURE_ANCHOR.md`
- `docs/ai/DEVELOPMENT_RULES.md`
- `docs/ai/RUNTIME_PIPELINE.md`
- `docs/ai/SMARTSTAT_RUNTIME_MAP.md`
- `docs/ai/SMARTSTAT_AI_BOOTSTRAP.md`

Deliver:
1. a short ChatGPT starter
2. a deep ChatGPT starter
3. a current-state summary
4. any truth-layer warnings (repo-truth vs session-truth)

TASK:
[PASTE TASK HERE]
