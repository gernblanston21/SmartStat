# Resolution Explainability Guide

Phase 8 starts a read-only architecture scaffold for semantic resolution explainability.

## Purpose

This layer explains how a normalized semantic record is positioned for future planning work.
It does not execute runtime behavior, does not apply plans, and does not perform planner execution.

## Search Explainability vs Resolution Explainability

Search explainability:
- answers: "why did this record match this search term?"
- driven by deterministic search ranking (`exact_id`, `prefix`, `substring`, source-hint matching)
- output focus: ranked result lists and match reasons

Resolution explainability (Phase 8 scaffold):
- answers: "how is this semantic record framed for future deterministic planning/execution explanation?"
- driven by deterministic semantic context (record scope, lineage, relationships, query-path count)
- output focus: ordered explainability steps and deferred runtime-bridge boundaries

## Phase 8 Data Model

Current model is `resolution-explainability.v0` with:
- fixed ordered step kinds
- stable refs inside each step
- explicit deferred boundaries (`runtime apply`, `planner execution`, runtime integration)

## Phase 9 Relationship

Phase 9 adds a separate candidate-resolution scaffold (`candidate-resolution.v0`) that sits after
search explainability and alongside this selected-record explainability model.
It remains read-only semantic tooling and does not claim runtime resolver behavior.

## Why This Is Separate From Runtime

The scaffold is intentionally read-only.
It exists to bridge semantic inspection to future plan capture without coupling to SmartStat runtime behavior.

## Deferred

- runtime resolver/apply execution
- plan generation/execution
- runtime bridge activation logic
