# WP-20 Acceptance Planning (Governance Package)

Status date: `2026-03-10`  
WP-20 implementation state: `NOT STARTED`  
Governance package state: `COMPLETE THROUGH TARGET-17`

## Purpose

Define closeout-planning truth for the WP-20 governance package without
authorizing implementation and without starting runtime-bridge work.

## Governance Package Complete Means

1. Target-01 through Target-17 governance/docs/tests gates are defined.
2. Approval, charter, rehearsal, sign-off, authorization, and packet templates
   are present.
3. Sample packet and sample dry-run structures are present as non-authorizing
   examples only.
4. Branch/lane separation and fail-closed constraints are explicitly documented.
5. Runtime version-line fork governance rule is defined and preserves
   `SmartStat_v4.0.0_beta.vbs` as a frozen protected baseline.
6. Runtime version-line decision record template and guidance are defined, and
   runtime implementation remains blocked without explicit authorization,
   explicit version-line decision, and lane-entry satisfaction.
7. Runtime version-line evidence checklist and evidence schema are defined, and
   runtime implementation remains blocked without evidence completeness.
8. Runtime version-line evidence review procedure and signoff template are
   defined, and runtime implementation remains blocked without evidence
   review/signoff completion.
9. Governance closeout criteria and stop-or-advance decision template are
   defined, and governance completion remains non-authorizing for runtime
   implementation.

## Governance Package Complete Does Not Mean

1. Implementation authorization has been granted.
2. Runtime-bridge code may start by default.
3. Apply behavior, Trio integration, or SmartStat engine/apply calls are
   approved.
4. Protected runtime/config surfaces may be modified.

## Authorization Boundary

Implementation start still requires an explicit recorded authorization decision,
including branch-approval record and implementation-authorization record.

Closeout-planning alignment is non-authorizing by design.

## Branch / Lane Rule

Runtime-bridge implementation remains a distinct risk-class lane and must run on
a separate explicitly authorized implementation lane.

## Protected Surface Reminder

The following remain protected in this governance pass:

1. `SmartStat_v4.0.0_beta.vbs`
2. `SmartStat_TemplateConfig.ini`
3. `SmartStat_Mappings.ini`
4. `SmartStat_StaticOverrides.ini`
5. `docs/onair/plan-capture.schema.json`
6. `docs/onair/plan-validation-contract.md`

## Explicit Future Choices

1. Stop at governance completion.
2. Collect real authorization inputs and evidence.
3. Explicitly authorize a separate implementation lane.

No implied activation is permitted from closeout-planning alone.
