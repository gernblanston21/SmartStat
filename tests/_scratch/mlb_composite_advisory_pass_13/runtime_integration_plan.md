# PASS_13 Runtime Integration Plan

## Runtime Goal
Add the smallest possible read-only advisory-composite invocation point inside `SmartStat_v4.1.0.vbs` without introducing execution authority, tabfield writes, socket/apply behavior, or runtime payload mutation.

## Candidate Insertion Point
1. Primary (safe for first slice): `Slice2PlanBridge_RunAndExit` (`SmartStat_v4.1.0.vbs`, after `Slice2PlanBridge_Execute(outcome)` and before `Slice2PlanBridge_EmitEvidence(outcome)`).
- Why this is safest:
- The run already has fully populated read-only context (`outcome`, `preview_kind`, `tabfield_records`).
- It is already a diagnostic boundary with fail-closed error handling and explicit trace logging.
- Advisory evaluation can be added as a side-channel diagnostic surface without changing preview payload shape.

2. Secondary (not safe for first slice): `Slice2PlanBridge_BuildPreviewPayloadJson` (adjacent to `operator_hint_surface`, `operator_hint_detail`, `operator_hint_summary`, `operator_next_step_preview`).
- Why deferred:
- Any new preview field here changes emitted payload contract/surface.
- Current governance posture is safer with no new preview payload mutation in first integration step.

## Required Inputs
1. `composite_input_raw` (expression string to evaluate).
- Safe source for first slice: explicit optional named arg via existing `Slice1Ingress_GetNamedArg(...)` plumbing.
- Existing but unsafe-for-first-slice source: implicit extraction from `outcome("tabfield_records")` (ambiguous field ownership in current runtime flow).

2. `preview_kind` (already present in `outcome("preview_kind")`).

3. Mapping authority context (read-only): `SmartStat_Mappings.ini` as already used by `SmartStat_AdvisoryComposite.vbs`.

## Advisory Call Boundary
- Invoke advisory engine only if `composite_input_raw` is non-empty.
- Keep invocation read-only and deterministic.
- Wrap call in local fail-closed error boundary (`On Error Resume Next` + deterministic diagnostic code path).
- Do not let advisory result alter `status`, `error_code`, `shouldPass`, or any execution path.

## Allowed Operator-Visible Output
For the first slice, only advisory-diagnostic/evidence visibility is allowed (no execution authority, no tabfield mutation):
- `recognized` (derived from classification)
- `reason_code`
- `confidence_bucket`
- `has_ambiguity`

Notes:
- Surface via deterministic diagnostic/evidence lines only in first slice.
- Keep existing `SMARTSTAT_SLICE2_STATUS` and `SMARTSTAT_SLICE2_ERROR_CODE` behavior unchanged.

## Forbidden Surfaces
- `tabfield_writes` (`tabfield:set_custom_property`, `page:set_property`)
- `socket_operations` (`sock:*`)
- `apply_take_cue` (`page:take`, `page:cue`, `page:takeout`, apply/take/cue semantics)
- `runtime_payload_mutation` (no new/changed preview payload keys in first slice)
- `authoritative_branching` (no advisory-driven runtime decisions)
- output-map/syntax generation mutation

## Recommended First Implementation Slice
`MLB_ONAIR_COMPOSITE_ADVISORY_RUNTIME_INTEGRATION_SLICE_14_DIAG_ONLY`

Bounded scope:
1. Add one guarded advisory invocation point in `Slice2PlanBridge_RunAndExit`.
2. Accept only explicit optional `composite_input` named arg as the advisory expression source.
3. Call the existing advisory helper and emit one deterministic diagnostic evidence line containing:
- `recognized`
- `reason_code`
- `confidence_bucket`
- `has_ambiguity`
4. Do not change preview payload JSON, runtime status/error semantics, TrioCmd surfaces, or mapping/config files.

## Risks
- Helper loading boundary: `SmartStat_v4.1.0.vbs` currently does not import `SmartStat_AdvisoryComposite.vbs`; integration method must be explicit and fail-closed.
- Input provenance risk: implicit tabfield-derived expression selection is not currently canonical; first slice should avoid this.
- Governance risk: preview payload expansion would exceed first-slice minimal-mutation boundary.
- Mapping overlap hard-stop risk: helper already hard-fails on category/qualifier key overlap; integration wrapper must preserve fail-closed behavior.

## Go / No-Go Recommendation
GO, but only for the bounded diag-only first slice above.

No-Go condition for this pass scope:
- Any attempt to add new preview payload fields, execution branching, Trio/apply/cue/tabfield/socket surfaces, or implicit tabfield inference beyond explicit input contract.
