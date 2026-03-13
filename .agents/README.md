# SmartStat Agent Skills

This directory contains repo-local skills used by Codex and related AI tools.

These skills are part of the SmartStat AI operating system, but they are not
authoritative over repo governance.

Authority order still follows:

1. `AGENTS.md`
2. `SESSION.md`
3. `ROADMAP.md`
4. `docs/viz-trio/`
5. `docs/architecture/smartstat-architecture.md`
6. `docs/onair/`
7. `docs/ai/*`
8. `.agents/skills/*`

---

## Relationship to the AI Kernel

The SmartStat AI Kernel is defined in `docs/ai/AI_KERNEL.md` and indexed from
`docs/ai/SYSTEM_INDEX.md`.

Skills are loaded after governance and architecture context are established.

That means:

- skills guide workflow
- skills do not authorize changes
- skills do not override governance
- skills should be selected before planning when the domain is clear

---

## Expected Skill Flow

1. load governance
2. load kernel architecture
3. identify task domain
4. route to the correct skill
5. plan and execute within SmartStat rules

---

## Skill Inventory

Each skill lives under:

`.agents/skills/<skill-name>/SKILL.md`

The skill router is defined in:

`docs/ai/AGENTS_SKILL_ROUTER.md`
