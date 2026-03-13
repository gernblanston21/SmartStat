# SmartStat Architecture Anchor

## Purpose

This is the stable architecture boot file for SmartStat AI sessions.

Use it to quickly establish the SmartStat system shape before reasoning about
specific runtime, config, roadmap, or semantic-lane tasks.

---

## Canonical System Shape

SmartStat is a deterministic Viz Trio broadcast-engine project with:

- a protected runtime core
- a separately modeled semantic architecture/tooling lane

That split is intentional and must be preserved during AI reasoning.

---

## Stable Core Components

### Runtime Core
- `SmartStat_v4.0.0_beta.vbs`

### Configuration Family
- `SmartStat_Mappings.ini`
- `SmartStat_Mappings.learn.ini`
- `SmartStat_MappingsNBA.ini`
- `SmartStat_MappingsNBA.learn.ini`
- `SmartStat_MappingsNHL.ini`
- `SmartStat_MappingsNHL.learn.ini`
- `SmartStat_StaticOverrides.ini`
- `SmartStat_TemplateConfig.ini`

### Governance / State
- `AGENTS.md`
- `SESSION.md`
- `ROADMAP.md`

### Viz Trio Grounding
- `docs/viz-trio/`

### Architecture / Semantic Contracts
- `docs/architecture/smartstat-architecture.md`
- `docs/onair/`
- `docs/contracts/`

### AI / Agent Workflow
- `.agents/skills/`
- `docs/ai/`

---

## Architecture Split

### A. Protected Production Runtime Core

Properties:

- deterministic
- fail-closed
- transaction-safe
- operator-safe
- regression-sensitive
- not implicitly expandable

### B. Semantic Architecture / Tooling Lane

Properties:

- read-only by default
- contract-oriented
- explainability-oriented
- validation-oriented
- planning-oriented
- future-facing but non-authorizing

---

## Governing Principle

No architecture suggestion is valid if it:

- invents unsupported Viz Trio behavior
- bypasses fail-closed logic
- ignores transaction integrity
- breaks config compatibility
- broadens scope beyond repo evidence
- treats governance or test artifacts as runtime authorization

---

## AI Load Order

When booting a SmartStat session, read in this order:

1. `SYSTEM_INDEX.md`
2. `AGENTS.md`
3. `SESSION.md`
4. `ROADMAP.md`
5. `PROJECT_BRAIN.md`
6. `ARCHITECTURE_ANCHOR.md`
7. `DEVELOPMENT_RULES.md`
8. `RUNTIME_PIPELINE.md`
9. `SMARTSTAT_RUNTIME_MAP.md`
10. `docs/viz-trio/`

Use `project-context.md` when full re-grounding is needed.
