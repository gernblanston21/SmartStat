# SmartStat Development Rules

## Purpose

This file restates SmartStat engineering rules in a compact AI-friendly format.

It does not replace repo governance.
It exists so AI outputs match repo discipline more reliably.

---

## Non-Negotiable Rules

### 1. Preserve fail-closed behavior
If resolution, config, grounding, or runtime assumptions are uncertain, the
correct response is refusal, gating, or a safer alternative.

### 2. Preserve transaction integrity
No partial-commit behavior on fail paths.

### 3. Preserve deterministic behavior
No change should introduce output drift across comparable runs without explicit
scope, versioning, and evidence.

### 4. Preserve Viz Trio grounding discipline
Do not infer TrioCmd behavior, tabfield semantics, or operator workflow without
support from `docs/viz-trio/`.

### 5. Preserve INI governance
- preserve existing order
- preserve formatting when editing
- avoid duplicate keys
- respect contract expectations
- flag cross-file compatibility risk

### 6. Preserve external compatibility awareness
Always flag possible impact to:
- SmartStatTrayApp
- related tooling
- downstream consumers
- repo-local contracts

### 7. Do not broaden scope
Supporting docs, scratch material, harness artifacts, or future-lane files do
not authorize production behavior changes.

### 8. Distinguish truth layers
Always separate:
- current repo-truth
- current session-truth
- current merged implementation
- proposed future behavior

### 9. No silent refactors
If structure changes, say so.
If naming changes would be needed, stop and propose a safe alternative.

### 10. Prefer exact drop-in results
When asked for implementation help:
- be specific
- be copy/paste ready
- avoid pseudo-logic
- avoid vague abstractions

---

## Response Discipline

For meaningful SmartStat work, prefer this structure:

- Role Summary
- Summary
- Assumptions
- Implementation
- Regression Impact
- Risks
- Review
- Exact Code
- Validation Steps

---

## Evidence Discipline

If a change touches runtime behavior, include:

- impacted surfaces
- regression risk summary
- validation approach
- determinism implications
- config compatibility implications

---

## AI Failure Rule

If the AI cannot support a claim from repo evidence, it must say so clearly and
fail closed.
