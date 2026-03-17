# Viewer Read-Only Boundary

Status: `SMARTSTAT_VIEWER_LAYER_IMPLEMENTATION_PASS_01`  
Scope: Boundary contract for viewer/tooling layer.

## Purpose

Lock the Viewer Layer to a read-only, non-authorizing architecture that consumes
frozen runtime preview payloads without introducing runtime behavior.

## Allowed

1. Read frozen preview artifacts from test/tooling locations.
2. Validate contract shape and deterministic ordering.
3. Adapt approved payload fields into a read-only view-model.
4. Emit fail-closed contract errors for malformed artifacts.
5. Run deterministic adapter tests and hash/equality checks.

## Forbidden

1. Runtime mutation/apply behavior.
2. Trio command execution or integration requirements.
3. Socket behavior or integration requirements.
4. SmartStat engine mutation or apply calls.
5. Runtime slice reopening or new slice invention.
6. Artifact mutation of upstream WP-17/WP-18/WP-19/WP-20 sources.
7. UI action controls that imply execution privileges.
8. Projection-contract expansion in viewer adapter scope.

## Contract Discipline

1. Adapter input must fail closed on missing/malformed required keys.
2. Adapter output must expose allow-listed fields only.
3. Adapter must reject inconsistent overlap between:
   - `projection_metadata`
   - `rule_evaluation_trace_preview`
4. Adapter must preserve deterministic ordering across identical inputs.

## Runtime-Lane Separation

1. Viewer tooling consumes frozen read-only payloads.
2. Viewer tooling does not authorize runtime lane progression.
3. Runtime lane remains frozen through 02G and on HOLD unless separately approved.

## Operator Safety Alignment

Viewer output must align with live-safe principles:

1. non-blocking tooling behavior
2. fail-closed uncertainty handling
3. no hidden execution side effects

Reference:

1. `docs/viz-trio/environment_constraints.md`

## Non-Authorizing Statement

This boundary file is governance for tooling behavior only.  
It does not authorize runtime/apply integration.
