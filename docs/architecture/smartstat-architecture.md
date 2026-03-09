# SmartStat Architecture Overview

Last updated: 2026-03-09\
Status: Current repo-ready architecture reference\
Scope: Post-RC SmartStat architecture after WP-17 closeout, governance
re-baseline, and WP-18 kickoff gate completion

------------------------------------------------------------------------

## Purpose

This document defines the current SmartStat architecture at a system
level so future work can proceed without scope drift.

It answers:

-   what is frozen
-   what is active
-   what is deferred
-   how the current semantic architecture lane relates to the protected
    runtime core
-   where WP-17 and WP-18 fit in the broader system

This document is intentionally high-level and governance-aligned. It is
not a runtime specification and does not redefine lower-level contracts
already documented elsewhere.

------------------------------------------------------------------------

## Current Operating Model

SmartStat is now organized into two clearly separated halves:

1.  **Protected production/runtime core**
2.  **Semantic architecture/tooling lane**

The production/runtime core remains frozen and protected.\
The semantic architecture/tooling lane is the single active development
lane on `feature/semantic-layer`.

Current architectural posture:

-   Runtime/core baseline is frozen
-   Semantic architecture/tooling is active
-   WP-17 is complete
-   WP-18 kickoff gate is complete
-   WP-18 is eligible to start but not implemented
-   WP-19 and WP-20 remain deferred

------------------------------------------------------------------------

## Architecture Diagram

``` text
                                  SMARTSTAT - CURRENT ARCHITECTURE
┌─────────────────────────────────────────────────────────────────────────────────────────────────────┐
│                                         FROZEN RUNTIME BASELINE                                    │
│                              v4.0.0_beta / v4.0.0_RC1 / WP-10..14 baseline                        │
│                                                                                                     │
│  SmartStat_v4.0.0_beta.vbs                                                                          │
│    ├─ deterministic resolver behavior                                                               │
│    ├─ ambiguity fail-closed gating                                                                  │
│    ├─ transaction / commit guarantees                                                               │
│    ├─ runtime config loading                                                                        │
│    ├─ sport-aware mapping selection                                                                 │
│    └─ output generation / Trio-facing execution                                                     │
│                                                                                                     │
│  Protected / Frozen:                                                                                │
│    - resolver math                                                                                  │
│    - ambiguity handling                                                                             │
│    - transaction semantics                                                                          │
│    - production INI layout/order                                                                    │
└─────────────────────────────────────────────────────────────────────────────────────────────────────┘
                                                   │
                                                   ▼
┌─────────────────────────────────────────────────────────────────────────────────────────────────────┐
│                                   CONFIG + CONTRACT FOUNDATION                                      │
│                                                                                                     │
│  Production Config                                                                                  │
│    ├─ SmartStat_TemplateConfig.ini                                                                  │
│    ├─ SmartStat_Mappings.ini                                                                        │
│    ├─ SmartStat_MappingsNBA.ini                                                                     │
│    ├─ SmartStat_MappingsNHL.ini                                                                     │
│    ├─ SmartStat_Mappings.learn.ini                                                                  │
│    ├─ SmartStat_MappingsNBA.learn.ini                                                               │
│    ├─ SmartStat_MappingsNHL.learn.ini                                                               │
│    └─ SmartStat_StaticOverrides.ini                                                                 │
└─────────────────────────────────────────────────────────────────────────────────────────────────────┘
                                                   │
                                                   ▼
┌─────────────────────────────────────────────────────────────────────────────────────────────────────┐
│                             ACTIVE LANE: SEMANTIC ARCHITECTURE / TOOLING                            │
│                                       branch: feature/semantic-layer                                │
└─────────────────────────────────────────────────────────────────────────────────────────────────────┘
                                                   │
                                                   ▼
┌─────────────────────────────────────────────────────────────────────────────────────────────────────┐
│                                      WP-15 / WP-16 FOUNDATION                                       │
│                                     semantic inspection layer                                       │
└─────────────────────────────────────────────────────────────────────────────────────────────────────┘
                                                   │
                                                   ▼
┌─────────────────────────────────────────────────────────────────────────────────────────────────────┐
│                                WP-17 - PLAN CAPTURE CONTRACT LAYER                                  │
│                                         STATUS: CLOSED                                              │
└─────────────────────────────────────────────────────────────────────────────────────────────────────┘
                                                   │
                                                   ▼
┌─────────────────────────────────────────────────────────────────────────────────────────────────────┐
│                             WP-18 - PLAN VALIDATION CONTRACT LAYER                                  │
│                             STATUS: ELIGIBLE TO START / NOT IMPLEMENTED                             │
└─────────────────────────────────────────────────────────────────────────────────────────────────────┘
                                                   │
                                                   ▼
┌─────────────────────────────────────────────────────────────────────────────────────────────────────┐
│                                   WP-19 - PLAN VIEWER (DEFERRED)                                    │
└─────────────────────────────────────────────────────────────────────────────────────────────────────┘
                                                   │
                                                   ▼
┌─────────────────────────────────────────────────────────────────────────────────────────────────────┐
│                               WP-20 - RUNTIME BRIDGE / EXECUTION (DEFERRED)                         │
└─────────────────────────────────────────────────────────────────────────────────────────────────────┘
```

------------------------------------------------------------------------

## Branch Strategy

``` text
v4_Dev
└── Historical RC lineage baseline

feature/semantic-layer
└── single active lane
    ├── semantic architecture
    ├── docs/tests/tooling
    ├── read-only inspection
    └── WP-18 validation work (future)

future runtime-bridge branch (if approved)
└── used only when WP-20 execution work is explicitly approved
```
