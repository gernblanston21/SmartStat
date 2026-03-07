# Candidate Resolution Guide

Phase 9 adds a deterministic candidate-resolution scaffold for read-only semantic tooling.

## Purpose

This model explains a candidate set for a semantic search context:
- input term and normalized term
- preferred candidate
- alternate/rejected candidates
- deterministic ranking metadata
- ambiguity indicator
- provenance/supporting refs

It does not run runtime resolution, apply behavior, or planner execution.

## How It Differs From Other Explainability Layers

Search explainability:
- explains why a record matched a search term and how search ranking works

Selected-record resolution explainability (Phase 8):
- explains semantic context for one selected normalized record

Candidate-resolution explainability (Phase 9):
- explains deterministic candidate-set structure between search inspection and later plan-capture phases

Runtime resolution / plan execution:
- not implemented here
- explicitly deferred

## Deterministic Contract

Model version: `candidate-resolution.v0`

Ordering contract:
1. deterministic search score
2. record type
3. league
4. record id
5. selected-record preferred override (inspection-only)

Candidate statuses:
- `preferred`
- `alternate`
- `rejected` (source-only hints excluded from normalized candidate set)

## Deferred

- runtime resolver competition logic
- SmartStat runtime integration
- plan generation/execution
- apply/runtime bridge behavior
