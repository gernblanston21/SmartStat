# SmartStat Runtime Architecture Map

## Purpose

This file provides a compact visual and conceptual map of how SmartStat is
structured, how the semantic lane relates to the runtime core, and how the
WP-20 runtime bridge slices fit into the broader system.

Use this file to help new AI sessions understand the architecture quickly.

---

## Top-Level System Map

```text
Viz Trio Page / Template State
            ↓
Tabfield / Property Acquisition
            ↓
SmartStat Runtime Core
            ↓
INI Mapping / Override / Template Config Layer
            ↓
Deterministic Syntax Generation
            ↓
Protected Output / Apply Boundary
```

---

## Semantic Architecture Sidecar Map

```text
Semantic Source View
        ↓
Resolution Explainability
        ↓
Plan Capture Contract (WP-17)
        ↓
Plan Validation Layer (WP-18)
        ↓
Viewer Contract Layer (WP-19)
        ↓
Runtime Bridge Governance / Slices (WP-20)
```

This sidecar path is read-only unless separately authorized.

---

## Two-Layer Model

### Protected Runtime Core

Responsibilities:

- read page and tabfield state
- resolve mappings/config
- generate deterministic syntax
- preserve fail-closed runtime behavior
- preserve transaction safety
- avoid undocumented Trio behavior

### Semantic Architecture / Tooling Layer

Responsibilities:

- inspect
- model
- validate
- explain
- package contracts
- define runtime bridge boundaries

The semantic layer supports runtime reasoning but does not automatically mutate
runtime behavior.

---

## Runtime Bridge Sequence (WP-20)

```text
Slice 01  -> Read-only ingress
Slice 02  -> Read-only plan bridge
Slice 02A -> Contract hardening
Slice 02B -> Projection intake
Slice 02C -> Semantic interpretation intake
Slice 02D -> Issues summary intake
Slice 02E -> Resolution preview
Slice 02F -> Rule evaluation summary intake
```

All slices through this map are read-only, deterministic, fail-closed,
and mutation-blocked.

---

## Runtime Boundary

The following are outside the allowed scope of the read-only runtime bridge
sequence unless separately authorized:

- Trio mutation
- runtime apply behavior
- socket mutation
- SmartStat engine mutation
- graphics updates
- hidden bridge activation

---

## Truth-Layer Overlay

When looking at this map, always ask which layer is being described:

- repo-truth
- session-truth
- implementation-truth
- proposed future-state

The same box in the map may be discussed in more than one truth layer.
Do not assume planned or modeled components are already live runtime behavior.
