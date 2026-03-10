# WP-20 Approval Requirements (Target-01)

Status date: `2026-03-10`  
Scope: `WP-20 approval-evidence package (governance/docs/tests only)`  
Implementation state: `NOT STARTED`

## Purpose

Define the mandatory approval package that must exist before any WP-20
runtime-bridge implementation pass is allowed to begin.

This document does not authorize implementation.

## Required Approval Package (Pre-Implementation)

All items below are required before any runtime-bridge code changes:

1. `Runtime-Bridge Lane Charter`
- Confirms WP-20 as a separate risk-class lane.
- Defines exact in-scope runtime bridge surfaces and explicit non-goals.
- Confirms no hidden coupling to WP-17/WP-18/WP-19 read-only lanes.

2. `Branch Isolation Approval`
- Dedicated runtime-bridge branch/lane approved.
- Clear branch boundary from `feature/semantic-layer`.
- Merge and rollback ownership defined.

3. `Regression/Evidence Plan`
- Baseline evidence sources defined before implementation starts.
- Determinism comparison strategy defined.
- Fail-closed safety checks defined.

4. `Rollback Plan`
- Trigger conditions and rollback commands/procedure documented.
- Partial-change containment and abort ownership defined.
- Reversion scope defined for runtime bridge work only.

5. `Abort Criteria`
- Immediate halt criteria defined for deterministic drift, ambiguity leakage,
  transaction risk, or uncontrolled side effects.
- Escalation/decision owner defined.

6. `Prototype Acceptance Gate Matrix`
- Minimum gates for any WP-20 prototype listed and testable.
- Requires pass/fail evidence artifact per gate.

## Required Evidence Categories (Future Implementation Lane)

Future WP-20 implementation must produce evidence in these categories:

1. `Contract Traceability`
- Upstream inputs traced to WP-17/WP-18/WP-19 contract surfaces only.
- No undocumented input surfaces.

2. `Determinism and Replay Stability`
- Before/after comparison evidence for deterministic behavior.
- Any intentional differences explicitly justified and reviewed.

3. `Fail-Closed Safety`
- Evidence that ambiguity and validation safety posture remains fail-closed.
- Evidence that unexpected states halt rather than auto-repair.

4. `Runtime Boundary Control`
- Evidence that bridge behavior is explicit, bounded, and reviewable.
- No hidden apply/bridge side effects.

5. `Operational Reversibility`
- Rollback rehearsed or otherwise evidenced.
- Abort path demonstrated and documented.

6. `Change Surface Accounting`
- Exact file/class/component change inventory.
- Explicit statement of protected surfaces left unchanged where required.

## Rollback Criteria (Future Implementation Lane)

Rollback is mandatory if any of the following occur:

1. Deterministic behavior deviates without approved justification.
2. Fail-closed behavior regresses (ambiguity/validation gating weakens).
3. Runtime bridge behavior escapes approved scope.
4. Unexpected side effects appear outside approved runtime bridge surfaces.
5. Protected files/surfaces are modified without prior approval.

## Abort Criteria (Immediate Stop)

Implementation lane must halt immediately when:

1. Unreviewed runtime/apply/bridge coupling is discovered.
2. Determinism evidence is inconclusive or contradictory.
3. Safety regression cannot be explained within approved scope.
4. Rollback path is unclear or unverified.

## Minimum Acceptance Gates for Any WP-20 Prototype

All gates are mandatory:

1. `Gate A: Scope Gate`
- Prototype scope matches approved lane charter exactly.

2. `Gate B: Determinism Gate`
- Determinism evidence passes the approved comparison plan.

3. `Gate C: Fail-Closed Gate`
- Fail-closed safety checks pass with explicit evidence.

4. `Gate D: Boundary Gate`
- No unapproved side effects, no hidden runtime/apply behavior expansion.

5. `Gate E: Rollback Gate`
- Rollback procedure is validated and executable.

6. `Gate F: Governance Gate`
- Required evidence artifacts are present and reviewed before merge.

## Branch / Lane Rule

WP-20 runtime-bridge implementation is a distinct risk class and must run on a
separate dedicated branch/lane. It must not be mixed with the current
read-only semantic architecture lane.

## Current Target-01 Outcome

WP-20 Target-01 defines approval requirements and evidence expectations only.
No runtime bridge implementation starts in this pass.
