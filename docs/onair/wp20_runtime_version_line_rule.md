# WP-20 Runtime Version-Line Fork Rule (Target-13)

Status date: `2026-03-10`  
Scope: `Governance/docs/tests-only runtime version-line governance rule`  
Implementation state: `NOT STARTED`

## Purpose

Define the mandatory runtime version-line fork rule that must govern any future
explicitly authorized WP-20 runtime implementation work.

## Non-Goals

1. Does not authorize implementation.
2. Does not execute runtime bridge/apply/Trio/engine/viewer integration work.
3. Does not create, rename, or modify runtime `.vbs` files in this pass.
4. Does not modify protected runtime/config/schema/validation surfaces.
5. Does not start WP-20 implementation.

## Protected-Baseline Rule

`SmartStat_v4.0.0_beta.vbs` is a frozen protected baseline and must not be
modified by any WP-20 runtime implementation target.

## Frozen-Baseline Rationale

The frozen baseline must be retained for:

1. Regression comparison against pre-implementation behavior.
2. Rollback safety if future runtime-line work must be reverted.
3. Governance comparison and audit traceability across version lines.

## Runtime Version-Line Options

Default runtime fork line:

1. `SmartStat_v4.1.0.vbs`

Allowed alternate runtime fork line (governance decision required first):

1. `SmartStat_v4.2.0.vbs`

## Version-Line Decision Gate (Mandatory Before Runtime Code Work)

Governance must explicitly choose one runtime version line before any runtime
code changes begin:

1. `SmartStat_v4.1.0.vbs` (default), or
2. `SmartStat_v4.2.0.vbs` (allowed alternate with explicit governance choice).

If no explicit version-line decision is recorded, runtime implementation may not
start.

## Chosen-Line Enforcement Rule

Once the runtime version line is chosen, all WP-20 runtime implementation work
must use that chosen new runtime file/version line.

Cross-line mixing or fallback to direct edits on the frozen baseline is not
permitted.

## Direct-Edit Prohibition

Direct WP-20 runtime implementation edits to `SmartStat_v4.0.0_beta.vbs` are
explicitly prohibited.

## Explicit Non-Authorizing Boundary

This rule document is governance-only and non-authorizing. WP-20 remains
`NOT STARTED`, and implementation still requires explicit recorded
implementation authorization and separate branch/lane entry controls.
