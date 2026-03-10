# WP-20 Regression and Evidence Execution Plan (Target-02)

Status date: `2026-03-10`  
Scope: `WP-20 Target-02 governance gate (docs/tests only)`  
Implementation state: `NOT STARTED`

## Purpose

Define the pre-approved regression and evidence execution structure that a
future WP-20 runtime-bridge implementation lane must execute before merge.

This plan does not authorize implementation.

## Evidence Collection Categories

Future WP-20 implementation lane must collect evidence in these categories:

1. `Scope Compliance`
- Proof that changed files/components match the approved lane charter.

2. `Determinism Stability`
- Repeat-run comparisons showing deterministic behavior under identical inputs.

3. `Fail-Closed Safety`
- Evidence that unsafe or ambiguous conditions refuse/halt as designed.

4. `Boundary Integrity`
- Evidence that unapproved runtime/apply/bridge/viewer surfaces were not used.

5. `Rollback Reversibility`
- Evidence that rollback triggers, actions, and post-rollback state are valid.

6. `Protected Surface Integrity`
- Evidence that protected files/contracts remained unchanged unless explicitly
  approved in a separate gate.

## Required Checkpoints (Execution Structure)

The future implementation lane must execute checkpoints in this order:

1. `Checkpoint 0: Baseline Capture`
- Capture pre-change branch hash, environment metadata, and file inventory.

2. `Checkpoint 1: Scope/Boundary Preflight`
- Record allowed/forbidden surface assertions before code changes.

3. `Checkpoint 2: Determinism Pack`
- Run approved deterministic replay/compare packs for bridge scope.

4. `Checkpoint 3: Fail-Closed Pack`
- Run refusal/safety assertions for ambiguous/unsafe bridge conditions.

5. `Checkpoint 4: Boundary Pack`
- Prove no unapproved Trio/apply/viewer coupling or hidden side effects.

6. `Checkpoint 5: Rollback Rehearsal`
- Execute rollback procedure and capture evidence artifacts.

7. `Checkpoint 6: Merge-Gate Summary`
- Produce final evidence index and gate PASS/FAIL matrix.

## Required Regression Packs (Future Lane)

At minimum, the following packs must exist and run:

1. Determinism repeat-run pack.
2. Fail-closed safety/refusal pack.
3. Boundary integrity pack.
4. Rollback rehearsal pack.
5. Protected-surface integrity pack.

## Rollback Evidence Requirements

Rollback evidence must include:

1. Rollback trigger statement (what failed and why rollback was required).
2. Exact rollback procedure transcript.
3. Post-rollback file integrity report.
4. Post-rollback determinism/boundary sanity results.
5. Gate owner sign-off on rollback completeness.

## Merge Gate Requirements (Future Prototype Branch)

Prototype branch is merge-ready only when all gates pass:

1. `Gate 1 Scope` - approved charter scope only.
2. `Gate 2 Determinism` - deterministic evidence accepted.
3. `Gate 3 Fail-Closed` - safety/refusal evidence accepted.
4. `Gate 4 Boundary` - boundary integrity evidence accepted.
5. `Gate 5 Rollback` - rollback evidence accepted.
6. `Gate 6 Governance` - required approvals/sign-offs recorded.

Any failed gate blocks merge.

## Evidence Artifact Location Rule

All WP-20 evidence artifacts must be recorded under:

- `tests/wp-20/target-XX/artifacts/`

No root-level temporary evidence files are permitted.

## Target-02 Outcome

Regression/evidence execution plan is defined.  
WP-20 implementation remains NOT STARTED in this pass.
