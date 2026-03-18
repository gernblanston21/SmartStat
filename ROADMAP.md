# ROADMAP - SmartStat v4 Lifecycle

## Current State

- `v4.0.0_beta` and `v4.0.0_RC1` are frozen historical baselines.
- WP-10 through WP-14 are complete and retained as post-RC stabilization evidence.
- Active working lane in this repo is `feature/semantic-layer` (`v4.2.0` architecture/tooling track).
- Current semantic posture: Phase 9 candidate-resolution scaffold is present in `tools/semantic-source-view/`.
- WP-17 defines a versioned, deterministic plan-capture contract layer (docs/tests/tooling only).
- WP-18 is CLOSED as a validation-layer package (docs/tests/tooling only; runtime-independent).
- WP-19 is CLOSED as a read-only viewer-contract package over WP-17/WP-18 artifacts (docs/tests/tooling only; no UI/runtime behavior).
- WP-20 governance package is CLOSED / ACCEPTED as a governance/docs/tests-only runtime-bridge lane package; broader runtime mutation/apply implementation is NOT STARTED and is not implied by WP-19 closeout.

---

## Lane Re-Baseline (2026-03-08)

- Single active lane: `feature/semantic-layer` for semantic architecture/tooling only (docs/tests/read-only tooling).
- Frozen runtime/core baseline: `v4.0.0_beta` and `v4.0.0_RC1` (no implicit runtime execution lane is active).
- WP-18 lane state: CLOSED (2026-03-09) with acceptance evidence under `tests/wp-18/artifacts/wp18_validator_runs/`.
- WP-19 lane state: CLOSED (2026-03-09) with acceptance evidence under `tests/wp-19/`.
- WP-20 lane state: GOVERNANCE PACKAGE CLOSED / ACCEPTED (2026-03-10 governance/docs/tests only; broader runtime mutation/apply implementation not started; explicit approval still required).
- Branch boundary: runtime bridge/execution work requires explicit approval and should run on a separate dedicated branch when started.
- WP-15 through WP-17 closeout does not imply runtime bridge/apply integration.

---

## AI Context Alignment

The repo-local context system under `docs/ai/` must remain aligned to this roadmap.

When roadmap state changes, review at least:
- `docs/ai/project-context.md`
- `docs/ai/PROJECT_BRAIN.md`
- `docs/ai/ARCHITECTURE_ANCHOR.md`
- `docs/ai/RUNTIME_PIPELINE.md`

The roadmap remains authoritative over those summary files.

---

## Determinism Doctrine (v4+)

SmartStat guarantees within a given version plus config:
- identical input state plus config -> identical output
- identical STRICT harness run (same version/config) -> identical artifacts
- no implicit precedence via iteration order in resolver logic
- no filesystem-order-dependent behavior

---

# RC1 Phase (Historical Baseline - Locked)

## Objectives

- Validate STRICT harness across representative templates
- Confirm no behavioral drift since beta freeze
- Verify determinism logging (no math changes)
- Confirm transaction gating integrity

## Definition of Done

- No runtime errors
- No ambiguity leakage
- No unintended diff commits
- `RC1_CHECKLIST.md` fully executed
- Tag `v4.0.0_RC1` created

---

# Post-RC Tracks (`v4.1.0` -> `v4.2.0`)

## Guardrails (Post-RC)

- RC1 behavior is the baseline; changes must be intentional, scoped, and validated.
- No naming convention changes.
- No INI key reordering.
- Fail-closed semantics preserved (ambiguity + validation gates).
- Each WP must include: scope, DoD, regression plan, and validation artifacts.

## WP-10 (`v4.1.0`): Determinism Surface Stabilization (Explicit Ordering)

### Objective

Eliminate nondeterministic iteration surfaces that affect:
- visible output stability
- log emission stability
- harness artifact reproducibility
- resolver decision consistency
- INI traversal consistency

This is a controlled architectural stabilization pass.
No math, resolution logic, or precedence changes are permitted.

---

### Determinism Surface Classes (Mapped)

The following unordered iteration classes are in scope:

1. Dictionary key iteration (runtime collections)
2. Multi-source merge order (config overlays)
3. First-match resolver scans
4. `output_map` emission ordering
5. Log/diagnostic dump ordering
6. INI section traversal
7. INI key traversal (if iterative)
8. Filesystem enumeration (if used for config load)
9. List parsing rehydration into unordered containers
10. Candidate set construction during heuristic scans

Each surface must be classified as:
- behavior-affecting
- presentation-only
- non-impacting

HIGH-risk surfaces (behavioral) must be stabilized before emission-level sorting.

---

### Explicit Non-Scope

WP-10 must NOT:
- change resolution precedence rules
- alter ambiguity detection behavior
- modify fail-closed gating
- reorder INI files
- refactor resolver algorithm design
- introduce silent precedence changes
- mask latent ambiguity bugs via sorting

---

### Implementation Phases

Phase 1 - Surface Audit (no code changes)
- confirm actual presence of each determinism surface class
- identify behavioral vs presentation-only cases
- document implicit precedence dependencies

Phase 2 - Emission Stabilization
- stabilize `output_map` ordering
- stabilize ambiguity dump ordering
- stabilize harness artifact serialization

Phase 3 - Behavioral Surface Hardening
- stabilize dictionary iteration used in resolver decisions
- stabilize overlay merge ordering explicitly
- stabilize candidate evaluation order

### WP-10 Phase 3 - Behavioral Surface Hardening (Targets)

Completed:
- Target #1 - Commit ordering determinism (`ApplyPlan.Keys` sorted before non-atomic commit loop)
- Target #2 - Resolver tie determinism (two-pass tie detection; ties fail closed)
- Target #3 - `TRANSFORMS_REGEX` determinism (sorted load/apply; strict conflict fail-closed)
- Target #4 - `AmbiguityContext` lifecycle determinism plus strict invariants (`AMBIGUOUS_CONTEXT_INVALID`; stable ambiguity emissions)
- Target #5 - `TryCanonLookupFlexible` deterministic first-match behavior (normalize-collision handling)
- Target #6 - `LoadIniSectionDictNormalized` / `LoadIni` deterministic normalized-key collision handling
- Target #7 - `SuggestQualifierMapping` / `ResolveFilterFragments` deterministic first-hit scanning (containment + fuzzy fallback)
- Target #8 - `ResolveQualifierSmart` candidate pool ordering determinism
- Target #9 - `ResolveCategorySmart` alias/canonical merge determinism
- Target #10 - Heuristic scanner input normalization
- Target #11 - Residual normalized lookup surfaces (if discovered) (completed)

Phase-3 complete: all documented HIGH nondeterministic behavioral surfaces stabilized with ordering-only fixes.
No ambiguity, scoring, resolver, or logging drift introduced.

Phase 4 - Regression Verification
- STRICT harness repeat-run validation
- confirm identical resolution outcomes
- confirm diffs are ordering-only
- document validation artifact record

WP-10 CLOSED - Determinism surface stabilization and regression verification complete.

---

### Definition of Done (Expanded)

- all HIGH-risk determinism surfaces stabilized
- all MEDIUM-risk surfaces stabilized for artifact consistency
- STRICT harness repeat-run produces identical artifacts
- no new ambiguity leakage
- no resolution outcome changes
- changelog plus validation record committed

---

### Versioning

This is a minor version bump (`v4.1.0`) because:
- output ordering changes
- diff behavior changes
- determinism guarantees are strengthened

### WP-10 Retrospective - Determinism Stabilization

WP-10 hardened SmartStat's resolution engine to guarantee deterministic behavior across runs.
All previously identified HIGH-risk nondeterministic surfaces (dictionary iteration order, candidate pool construction, alias/canonical merge order, normalize-first lookup helpers, and heuristic scanner ingress) were stabilized without altering resolver math, scoring rules, or ambiguity policy.

The work focused exclusively on deterministic ordering and fail-closed ambiguity preservation so that identical input state and configuration now produce identical outputs every time.

Phase-4 regression verification validated repeat-run determinism in both STRICT and runtime harness modes, with archived evidence under `/tests/wp-10/phase-4/`.

This milestone establishes SmartStat's first fully verified deterministic core and provides a stable foundation for future resolver enhancements and feature work.

No patch release permitted for this scope.

## WP-11 (`v4.1.0`): Harness Regression Pack Framework

### Scope

- define a repeatable regression-pack set of templates/scenarios
- standardize how artifacts are stored and compared

### Definition of Done

- regression pack documented and runnable
- artifact locations standardized
- clear pass/fail criteria captured

WP-11 CLOSED - Harness regression pack framework implemented (tests-only).
Evidence: `tests/wp-11/regression-pack/artifacts/compare/wp11_runA__wp11_runB/` (`PACK_PASS=True`)

## WP-12 (`v4.1.0`): Enhanced Learn System Validation

### Scope

- verify learn file writes are correct, stable, and governed
- ensure learn updates are validated (format + intent) before acceptance

### Definition of Done

- learn write validation rules implemented
- bad/partial writes are blocked or quarantined (fail-closed)
- validation artifacts recorded

WP-12 CLOSED - Enhanced learn system validation implemented (tests-only).
Evidence: `tests/wp-12/learn-validation/artifacts/wp12_runC/` (`RUN_PASS=True`)

## WP-13 CLOSED (`v4.1.0`): Resolver Performance Optimization

### Scope

- optimize resolver hot paths without changing resolution outcomes
- preserve determinism and logging semantics

### Definition of Done

- performance improvement measured
- no behavioral diffs in STRICT harness regression pack

RC note:
- Resolver performance optimization validated via regression harness under `tests/wp-13/resolver-perf/`.
- Optimization itself modifies runtime code paths and is therefore deferred until post-RC.
- Implementation change is preserved in local stash `WP-13 perf micro-opt (post-RC)`.

Evidence: `tests/wp-13/resolver-perf/` (STRICT regression harness validation)

## WP-14 (`v4.1.0`): TrayApp Alignment Preparation

### Scope

- define and stabilize the contract between SmartStat core + TrayApp
- ensure mappings/config expectations are explicit and version-safe

### Definition of Done

- contract documented
- validator suite implemented
- harness execution passing
- no SmartStat core behavior changes

WP-14 implementation package (RC-safe, docs/tests/tooling only):
- Contract spec: `docs/contracts/smartstat_trayapp_contract.md`
- Validator suite: `tests/wp-14/contract-validators/`
- Harness runner: `tests/wp-14/run_wp14.ps1`
- Run command: `pwsh -NoProfile -ExecutionPolicy Bypass -File tests/wp-14/run_wp14.ps1 -RunLabel <label>`
- Evidence location: `tests/wp-14/contract-validators/artifacts/<runLabel>/`
- RC constraint: no SmartStat core VBScript behavior changes; no production INI schema/order mutations

WP-14 CLOSED - TrayApp contract + validator harness implemented (docs/tests/tooling only).
Evidence: `tests/wp-14/contract-validators/artifacts/wp14_runB/` (historical closeout record) and `tests/wp-14/contract-validators/artifacts/recovery_wp14_20260308/` (`RUN_PASS=True`, `REPO_GATE_FAILURES=0`; expanded sport-aware coverage).

## WP-15 (`v4.2.0`): Semantic Inspection Foundation

### Scope

- establish semantic inspection as the first architecture baseline for post-RC planning work
- lock the existing Semantic Source View capability as the canonical read-only inspection surface
- preserve deterministic search/inspection contracts as reusable inputs for later explainability and planning work

### Definition of Done

- semantic inspection baseline documented (records, relationships, query paths, traceability, deterministic search)
- deterministic inspection contracts are explicit and reusable
- read-only inspection workflow is reproducible on `feature/semantic-layer`
- no SmartStat runtime behavior changes introduced

WP-15 implementation package (read-only tooling):
- React Semantic Source View baseline: `tools/semantic-source-view/` (Phase 5)
- Deterministic search explainability/ranking: Phase 6
- Deterministic search contract + lightweight tests: Phase 7
- Checkpoint tag: `semantic-view-phase7` (`30d15b6`)

WP-15 CLOSED - Semantic inspection foundation accepted (read-only tooling only).

## WP-16 (`v4.2.0`): Resolution Explainability

### Scope

- add a read-only explainability scaffold that models deterministic semantic resolution reasoning
- keep this as architecture/tooling only (no runtime apply behavior, no planner execution)
- bridge semantic search inspection outputs to future plan-capture contracts

### Definition of Done

- resolution explainability data model/schema defined for semantic tooling
- deterministic explainability fixture/scaffold available for local inspection
- explainability boundaries documented against search explainability and plan execution
- no SmartStat runtime behavior changes introduced

WP-16 implementation package (read-only tooling):
- Explainability guide: `tools/semantic-source-view/RESOLUTION_EXPLAINABILITY_GUIDE.md`
- Explainability schema: `tools/semantic-source-view/resolution-explainability.schema.json`
- Deterministic fixture + read-only panel scaffold: `tools/semantic-source-view/src/data/resolutionExplainability.fixture.ts`
- Scaffold checkpoint tag: `semantic-view-phase8-scaffold` (`55c56ef`)

WP-16 CLOSED - Phase 8 and Phase 9 read-only explainability/candidate-resolution scaffolds are accepted.
WP-16 handoff into WP-17 contract work is complete.

Validation evidence (2026-03-08):
- `npm run build` passed
- `npm run test` passed (`3` files, `13` tests)
- baseline UI render, Phase 8 scaffold render, Phase 9 candidate-resolution scaffold render, search-driven selection updates, browse-mode behavior, zero-normalized-result behavior, and deterministic visual repeatability all passed
- boundary checks passed: no runtime/apply behavior, no planner execution, and no runtime integration implied

## WP-17 (`v4.2.0`): Plan Capture Contract Layer (Docs/Tests/Tooling Only)

### Scope

- define a versioned captured-plan artifact contract and schema
- enforce contract shape via validator tooling and deterministic fixture harness
- keep capture layer read-only and explicitly separated from validation, viewer, and runtime execution

### Definition of Done

- contract doc + schema published
- validator + runner implemented under `tests/wp-17/`
- good/bad fixtures pass expected outcomes
- deterministic replay hash check passes in runner summary
- no SmartStat runtime behavior changes introduced

WP-17 implementation package (read-only docs/tests/tooling):
- Contract doc: `docs/onair/plan-capture-contract.md`
- Contract schema: `docs/onair/plan-capture.schema.json`
- Validator: `tests/wp-17/contract-validators/validate_plan_capture_contract.ps1`
- Harness runner: `tests/wp-17/run_wp17.ps1`
- Run command: `pwsh -NoProfile -ExecutionPolicy Bypass -File tests/wp-17/run_wp17.ps1 -RunLabel <label>`
- Evidence location: `tests/wp-17/contract-validators/artifacts/<runLabel>/`
- Constraint: no SmartStat runtime behavior changes; no planner execution; no runtime bridge/apply behavior

WP-17 CLOSED - Standalone capture contract layer defined (docs/tests/tooling only).
Evidence: `tests/wp-17/contract-validators/artifacts/wp17_contract_20260308/` (`RUN_PASS=True`, `DETERMINISM_REPLAY_PASS=True`).

## WP-18 (`v4.2.0`): Plan Validation

Status: CLOSED (accepted on 2026-03-09; validation-layer-only, runtime-independent).

### Scope

- define the Plan Validation Contract Layer between WP-17 capture and future WP-19/WP-20 layers
- validate captured plans using deterministic, read-only architecture rules
- keep validator work runtime-independent (no runtime execution, no Trio/apply behavior)

Kickoff governance artifacts:
- `docs/onair/plan-validation-contract.md`
- `docs/onair/wp18_kickoff_checklist.md`

WP-18 explicit boundaries:
- Allowed: captured-plan validation, deterministic rule evaluation, validator tooling, schema compatibility checks
- Not allowed: runtime execution, Trio integration, SmartStat engine calls, applying stats to graphics, captured-plan mutation

### Definition of Done

- plan validator implemented
- fixture suite passes expected good/bad cases
- harness runner reports `RUN_PASS=True` for plan validation scope
- deterministic replay check is stable across repeated runs
- no SmartStat runtime behavior changes introduced by validation tooling

WP-18 implementation package (read-only docs/tests/tooling):
- Validator runner: `tests/wp-18/validator/validator_runner.py`
- Result-model doc: `tests/wp-18/validator/validation_result_model.md`
- Validation phase/readme doc: `tests/wp-18/README.md`
- Replay/harness tests:
    - `tests/wp-18/replay/deterministic_replay_test.py`
    - `tests/wp-18/replay/determinism_rule_layer_test.py`
    - `tests/wp-18/replay/boundary_rule_layer_test.py`
    - `tests/wp-18/replay/result_model_hardening_test.py`
- Acceptance note: `tests/wp-18/ACCEPTANCE.md`
- Evidence location: `tests/wp-18/artifacts/wp18_validator_runs/target06/`

WP-18 CLOSED - Plan validation layer accepted with:
- structural + semantic + determinism + boundary rules
- hardened result model
- deterministic semantic interpretation metadata for default stats scope (`omitted scope => career` in stats context)
- deterministic replay/harness evidence

WP-19 consumption boundary (allowed):
- consume WP-18 `validation_result` payloads as read-only artifacts only
- consume rule evaluations, refusal diagnostics, deterministic identities, and semantic interpretation metadata
- must not imply runtime apply behavior or runtime bridge activation

## WP-19 (`v4.2.0`): Plan Viewer

Status: CLOSED (accepted 2026-03-09; read-only viewer-contract package only).

### Scope

- extend read-only inspection UX to include deterministic plan-view semantics
- provide explainable plan browsing/debugging without apply/runtime integration
- keep viewer contracts versioned and compatible with semantic + validation outputs

### Kickoff Gate (Read-Only Consumer Layer)

WP-19 may consume the following immutable inputs:
- WP-17 captured-plan artifacts
- WP-18 `validation_result` payloads
- WP-18 `rule_evaluations`
- WP-18 refusal diagnostics (`errors`, `warnings`, refusal codes/messages)
- WP-18 deterministic identities (`input_identity`, `normalized_plan_hash`, `replay_identity`, `validator_run_identity`)
- WP-18 `semantic_interpretation` metadata

### Forbidden Behaviors (Must Not)

WP-19 must not:
- execute runtime behavior
- trigger apply behavior
- introduce bridge behavior
- call SmartStat engine/apply surfaces
- mutate captured-plan artifacts
- mutate validation artifacts

### Minimal Viewer Contract Surfaces (Pre-Implementation)

- Artifact intake surface: accepts WP-17/WP-18 artifacts as read-only inputs
- Validation summary surface: exposes `status`, deterministic `errors`, and deterministic `warnings`
- Rule evaluation surface: exposes stable phase/category/rule ordering from `rule_evaluations`
- Deterministic identity surface: exposes artifact identity and replay-normalization identities for traceability
- Semantic interpretation surface: exposes `semantic_interpretation` exactly as validation metadata (no runtime inference)

### Definition of Done

- read-only plan viewer contract documented
- read-only projection and adapter contracts documented and validated with deterministic harness evidence
- compatibility rules documented across semantic/explainability/plan contracts
- no runtime apply behavior changes introduced

### Accepted WP-19 Package Surface (Docs/Tests/Tooling Only)

- Contract doc: `docs/onair/plan-viewer-contract.md`
- Harness suite:
    - `tests/wp-19/harness/read_only_intake_contract_test.py`
    - `tests/wp-19/harness/viewer_projection_contract_test.py`
    - `tests/wp-19/harness/projection_consumption_contract_test.py`
    - `tests/wp-19/harness/projection_to_view_model_adapter_contract_test.py`
- Target evidence/readme tree:
    - `tests/wp-19/target-01/` through `tests/wp-19/target-06/`
- Acceptance note:
    - `tests/wp-19/ACCEPTANCE.md`

### Future Implementation Lane May Consume (Read-Only)

- WP-19 projection contract surfaces:
    - `projection_contract`, `projection_kind`, `input_artifact`, `input_identity`, `status_summary`, `issues_summary`, `rule_evaluation_summary`, `deterministic_identity_summary`, `semantic_interpretation_summary`
- WP-19 adapter contract surfaces:
    - adapter identity fields plus `view_model` sections: `status_view`, `issues_view`, `rules_view`, `semantic_view`, `trace_view`
- deterministic ordering guarantees established by WP-19 harnesses
- explicitly not included:
    - runtime execution, apply behavior, bridge behavior, mutation, or runtime inference

## WP-20 (`v4.2.0+`): Runtime Bridge

Status: GOVERNANCE PACKAGE CLOSED / ACCEPTED (2026-03-10 governance/docs/tests only; implementation NOT STARTED).

### Scope

- introduce controlled bridge points from validated plan artifacts toward runtime integration
- defer runtime bridge implementation until semantic inspection, explainability, plan capture, and plan validation are stable
- preserve fail-closed and deterministic discipline while defining bridge constraints

### Governance Package (Closed/Accepted, Non-Authorizing)

- WP-20 is explicitly separate from WP-17/WP-18/WP-19 read-only architecture packages.
- WP-20 governance package is planning/docs/tests only in this pass and does not authorize implementation.
- Gate checklist reference: `docs/onair/wp20_kickoff_checklist.md`
- Approval requirements reference: `docs/onair/wp20_approval_requirements.md`
- Target-01 evidence template reference: `tests/wp-20/target-01/approval_evidence_template.md`
- Target-02 lane charter reference: `docs/onair/wp20_lane_charter.md`
- Target-02 regression/evidence plan reference: `docs/onair/wp20_regression_evidence_plan.md`
- Target-03 rehearsal protocol reference: `docs/onair/wp20_rehearsal_protocol.md`
- Target-03 scaffold reference: `tests/wp-20/target-03/README.md`
- Target-04 rehearsal manifest template reference: `docs/onair/wp20_rehearsal_manifest_template.md`
- Target-04 gate review checklist reference: `docs/onair/wp20_gate_review_checklist.md`
- Target-04 scaffold reference: `tests/wp-20/target-04/README.md`
- Target-05 implementation-authorization record reference: `docs/onair/wp20_implementation_authorization_record.md`
- Target-05 branch-approval record reference: `docs/onair/wp20_branch_approval_record.md`
- Target-05 scaffold reference: `tests/wp-20/target-05/README.md`
- Target-06 authorization packet index template reference: `docs/onair/wp20_authorization_packet_index_template.md`
- Target-06 packet completeness checklist reference: `docs/onair/wp20_packet_completeness_checklist.md`
- Target-06 scaffold reference: `tests/wp-20/target-06/README.md`
- Target-07 authorization packet draft template reference: `docs/onair/wp20_authorization_packet_draft_template.md`
- Target-07 packet verification dry-run template reference: `docs/onair/wp20_packet_verification_dry_run_template.md`
- Target-07 scaffold reference: `tests/wp-20/target-07/README.md`
- Target-08 authorization packet sample reference: `docs/onair/wp20_authorization_packet_sample.md`
- Target-08 packet verification dry-run sample reference: `docs/onair/wp20_packet_verification_dry_run_sample.md`
- Target-08 scaffold reference: `tests/wp-20/target-08/README.md`
- Closeout-planning note reference: `tests/wp-20/ACCEPTANCE_PLANNING.md`
- Target-09 authorization-input collection template reference: `docs/onair/wp20_authorization_input_collection_template.md`
- Target-09 authorization-input evidence register reference: `docs/onair/wp20_authorization_input_evidence_register.md`
- Target-09 scaffold reference: `tests/wp-20/target-09/README.md`
- Target-10 packet-population readiness plan reference: `docs/onair/wp20_packet_population_readiness_plan.md`
- Target-10 authorization-input owner assignment template reference: `docs/onair/wp20_authorization_input_owner_assignment_template.md`
- Target-10 scaffold reference: `tests/wp-20/target-10/README.md`
- Target-11 authorization-input tracking ledger reference: `docs/onair/wp20_authorization_input_tracking_ledger.md`
- Target-11 readiness status report template reference: `docs/onair/wp20_readiness_status_report_template.md`
- Target-11 scaffold reference: `tests/wp-20/target-11/README.md`
- Target-12 tracking/reporting operating-rhythm reference: `docs/onair/wp20_tracking_reporting_operating_rhythm.md`
- Target-12 readiness review meeting template reference: `docs/onair/wp20_readiness_review_meeting_template.md`
- Target-12 scaffold reference: `tests/wp-20/target-12/README.md`
- Target-13 runtime version-line fork rule reference: `docs/onair/wp20_runtime_version_line_rule.md`
- Target-13 runtime implementation lane entry checklist reference: `docs/onair/wp20_runtime_implementation_lane_entry_checklist.md`
- Target-13 scaffold reference: `tests/wp-20/target-13/README.md`
- Target-14 runtime version-line decision record template reference: `docs/onair/wp20_runtime_version_line_decision_record_template.md`
- Target-14 runtime version-line decision guidance reference: `docs/onair/wp20_runtime_version_line_decision_guidance.md`
- Target-14 scaffold reference: `tests/wp-20/target-14/README.md`
- Target-15 runtime version-line evidence checklist reference: `docs/onair/wp20_runtime_version_line_evidence_checklist.md`
- Target-15 runtime version-line evidence schema reference: `docs/onair/wp20_runtime_version_line_evidence_schema.md`
- Target-15 scaffold reference: `tests/wp-20/target-15/README.md`
- Target-16 runtime version-line evidence review procedure reference: `docs/onair/wp20_runtime_version_line_evidence_review_procedure.md`
- Target-16 runtime version-line evidence signoff template reference: `docs/onair/wp20_runtime_version_line_evidence_signoff_template.md`
- Target-16 scaffold reference: `tests/wp-20/target-16/README.md`
- Target-17 governance closeout criteria reference: `docs/onair/wp20_governance_closeout_criteria.md`
- Target-17 stop-or-advance decision template reference: `docs/onair/wp20_stop_or_advance_decision_template.md`
- Target-17 scaffold reference: `tests/wp-20/target-17/README.md`
- Target-18 governance package closeout summary reference: `docs/onair/wp20_governance_package_closeout_summary.md`
- Target-18 governance acceptance record template reference: `docs/onair/wp20_governance_acceptance_record_template.md`
- Target-18 scaffold reference: `tests/wp-20/target-18/README.md`
- Target-19 governance evidence index reference: `docs/onair/wp20_governance_evidence_index.md`
- Target-19 final non-authorizing closure note reference: `docs/onair/wp20_final_non_authorizing_closure_note.md`
- Target-19 scaffold reference: `tests/wp-20/target-19/README.md`
- Governance closeout acceptance note reference: `docs/onair/wp20_governance_closeout_acceptance_note.md`
- Post-closeout runtime boundary note reference: `docs/onair/wp20_post_closeout_runtime_boundary_note.md`
- Closeout scaffold reference: `tests/wp-20/closeout/README.md`

### Allowed Upstream Inputs (Read-Only)

- WP-17 captured-plan artifacts and contract/schema evidence
- WP-18 validation outputs: `validation_result`, `rule_evaluations`, refusal diagnostics, deterministic identities, and `semantic_interpretation` metadata
- WP-19 projection/consumption/adapter contract surfaces and harness evidence

### Forbidden Pre-Implementation Behaviors

- no runtime mutation/apply code paths beyond separately authorized read-only preview slices
- no apply behavior implementation
- no Trio integration
- no SmartStat engine/apply calls
- no mutation of captured/validated/projected/adapter artifacts
- no runtime side-effect inference beyond read-only artifacts

### Required Approval Conditions Before Implementation

- explicit WP-20 implementation approval recorded in governance docs
- dedicated branch/lane confirmed for runtime bridge risk isolation
- pre-implementation regression/evidence plan approved (determinism, fail-closed behavior, rollback plan)
- runtime bridge risk-class review approved before any code mutation
- runtime version line is explicitly selected before runtime code changes (default `SmartStat_v4.1.0.vbs`; allowed alternate `SmartStat_v4.2.0.vbs`) and `SmartStat_v4.0.0_beta.vbs` remains frozen for regression/rollback/governance comparison
- runtime version-line decision record is completed and linked to Target-05 authorization artifacts plus Target-13 runtime lane-entry controls before runtime code changes begin
- runtime version-line decision evidence checklist/schema completeness is verified before runtime code changes begin
- runtime version-line evidence review is completed and signoff status is `signoff_complete` before runtime code changes begin
- governance closeout stop-or-advance decision is recorded; an advance outcome is non-authorizing and separate explicit implementation authorization remains required before runtime code changes begin

### Target-01 through Target-19 Governance Outcomes

- approval prerequisites are defined before any WP-20 code changes
- required evidence categories are defined for a future implementation lane
- rollback criteria and abort/fail-closed triggers are defined
- minimum acceptance gates are defined for any future runtime bridge prototype
- concrete lane charter is defined for branch/scope isolation
- pre-approved regression/evidence execution structure is defined
- pre-implementation rehearsal protocol is defined for dry-run flow validation
- rehearsal artifact pack layout, naming rules, checkpoint order, and fail/abort handling are defined
- formal rehearsal manifest template is defined for deterministic evidence intake
- formal sign-off gate checklist is defined with required roles and review surfaces
- gate outcomes are defined: `implementation_ready`, `hold`, and `fail`
- minimum evidence requirements are defined before implementation authorization
- formal implementation-authorization record template is defined with required decision inputs and outcomes
- formal implementation branch-approval record template is defined with branch isolation and ownership fields
- authorization and revocation ownership expectations are explicitly defined
- gate outcomes are defined: `authorized_to_start_implementation`, `hold`, and `denied`
- formal authorization packet index template is defined to bind required pre-implementation inputs
- formal packet completeness checklist is defined with verification outcomes: `packet_complete`, `packet_incomplete`, and `packet_invalid`
- packet verification is explicitly defined as non-authorizing; separate authorization decision still governs implementation start
- formal authorization packet draft template is defined for future packet-instance population
- formal packet verification dry-run template is defined with required review sections and outcomes: `dry_run_pass`, `dry_run_hold`, and `dry_run_fail`
- dry-run success is explicitly non-authorizing; separate implementation authorization remains required
- sample-filled authorization packet instance and sample dry-run report structures are defined as non-authorizing examples only
- hypothetical dry-run evidence-pack layout references are defined for structural guidance only
- formal authorization-input collection template is defined for future real input intake
- formal authorization-input evidence register is defined with fail-closed missing/unverified evidence handling
- authorization-input collection remains non-authorizing and does not start implementation
- formal packet-population readiness plan is defined for future real input collection execution planning
- formal authorization-input owner assignment template is defined with fail-closed missing-owner handling
- no real authorization-input collection is executed in this pass
- formal authorization-input tracking ledger is defined with fail-closed stale/missing/unknown status handling
- formal readiness status report template is defined for summary + rollup + blocker/next-action reporting
- tracking/reporting setup remains non-authorizing and does not start implementation
- formal tracking/reporting operating rhythm is defined with fail-closed handling for skipped/missed/unclear cadence
- formal readiness review meeting template is defined with non-authorizing decision vocabulary and required blocker/action capture
- formal runtime version-line fork rule is defined with explicit prohibition on direct WP-20 runtime edits to `SmartStat_v4.0.0_beta.vbs`
- formal runtime implementation lane entry checklist is defined with mandatory version-line decision gate and fail-closed undecided/ambiguous handling
- formal runtime version-line decision record template is defined with exactly-one-selection and fail-closed blank/multiple/conflicting/ambiguous handling
- formal runtime version-line decision guidance is defined and tied to Target-05 authorization linkage and Target-13 lane-entry controls
- formal runtime version-line evidence checklist is defined for decision review, exactly-one-selection verification, required linkage verification, and frozen-baseline preservation evidence checks
- formal runtime version-line evidence schema is defined with required artifact shape, selected-version constraints, frozen-baseline confirmation, and fail-closed invalid/missing/ambiguous handling
- formal runtime version-line evidence review procedure is defined with required linkage checks, frozen-baseline checks, review outcomes, and fail-closed handling
- formal runtime version-line evidence signoff template is defined with required identity/signer/outcome/linkage fields and fail-closed unsigned/incomplete/conflicting handling
- formal governance closeout criteria are defined with required evidence categories and fail-closed missing/incomplete/conflicting handling
- formal stop-or-advance decision template is defined with only two allowed outcomes and explicit non-authorizing advance-path handling
- formal governance package closeout summary is defined with major-gate references and explicit non-authorizing closeout boundary
- formal governance acceptance record template is defined with required completeness/frozen-baseline/non-authorizing acknowledgement fields and fail-closed handling
- formal governance evidence index is defined with categorized artifact traceability and explicit non-authorizing indexing boundary
- formal final non-authorizing closure note is defined with explicit runtime-not-started closure language and fail-closed ambiguity handling
- formal governance closeout acceptance note is defined with explicit governance-only acceptance language and fail-closed overstated-language handling
- formal post-closeout runtime boundary note is defined with explicit runtime-permission separation and fail-closed ambiguous-boundary handling
- this pass does not start runtime bridge implementation

### Branch / Lane Rule

- runtime bridge/execution implementation must run on a separate dedicated lane from the current read-only semantic lane
- `feature/semantic-layer` remains governance/docs/tests/tooling for kickoff gating until explicit implementation approval is granted

### Closeout-Planning Alignment (Non-Authorizing)

- WP-20 governance package is CLOSED / ACCEPTED as governance/docs/tests packaging, while broader runtime mutation/apply implementation remains NOT STARTED
- governance package completion does not equal implementation authorization
- implementation authorization still requires explicit recorded decision input + approval records
- broader runtime mutation/apply implementation remains blocked until explicit authorization, explicit recorded version-line decision, evidence completeness, evidence review/signoff, and runtime lane-entry conditions are all satisfied
- runtime bridge remains a distinct risk-class lane with mandatory branch separation
- protected surfaces remain protected until separately authorized implementation work

Explicit future choices:
1. Stop at governance completion.
2. Collect real authorization inputs.
3. Explicitly authorize a separate implementation lane.

### Runtime Lane Sequencing (Post Slice-01)

- Next planned runtime lane after `WP20_RUNTIME_SLICE_01_READONLY_INGRESS`: `WP20_RUNTIME_SLICE_02_READONLY_PLAN_BRIDGE`
- Objective: produce a deterministic read-only plan-bridge output/preview payload
- Lane remains read-only
- No apply behavior
- No Trio mutation behavior
- No socket mutation behavior
- Required validation: gate-OFF parity, gate-ON determinism, fail-closed negatives, boundary audit, and slice-1 carry-forward sanity check

- Next authorized step after the frozen current scaffold scope of `WP20_RUNTIME_SLICE_02_READONLY_PLAN_BRIDGE`: `WP20_RUNTIME_SLICE_02A_READONLY_PLAN_BRIDGE_CONTRACT_HARDENING`
- Objective: harden deterministic preview payload contract shape and contract-level fail-closed diagnostics
- Lane remains read-only
- No apply behavior
- No Trio mutation behavior
- No socket mutation behavior
- Required validation: gate-OFF parity, gate-ON determinism, contract-negative fail-closed checks, mutation boundary audit, and carry-forward validation of existing slice-02 scaffold positives and negatives

- Next authorized step after the frozen current scope of `WP20_RUNTIME_SLICE_02A_READONLY_PLAN_BRIDGE_CONTRACT_HARDENING`: `WP20_RUNTIME_SLICE_02B_READONLY_PLAN_BRIDGE_PROJECTION_INTAKE`
- Objective: intake exactly one upstream WP-19 projection artifact into the existing read-only plan-bridge path
- Lane remains read-only
- Joined preview output remains deterministic
- Joined preview output must keep `mutation_authorized=false`
- Fail closed if the projection artifact is missing, malformed, unsupported, or not runtime-eligible
- No apply behavior
- No Trio mutation behavior
- No socket mutation behavior
- Required validation: gate-OFF parity, deterministic repeat-run hash equality, carry-forward validation of existing slice-02A positives and negatives, projection-negative fail-closed checks, mutation boundary audit, and boundary check confirming no new Trio/socket/apply surfaces were introduced

- Next authorized step after the frozen current scope of `WP20_RUNTIME_SLICE_02B_READONLY_PLAN_BRIDGE_PROJECTION_INTAKE`: `WP20_RUNTIME_SLICE_02C_READONLY_PLAN_BRIDGE_SEMANTIC_INTERPRETATION_INTAKE`
- Objective: intake exactly the `semantic_interpretation_summary` metadata surface from the existing WP-19 projection artifact into the existing read-only plan-bridge path
- Intake limited to:
    - `semantic_interpretation_summary.scope_resolution`
    - `semantic_interpretation_summary.effective_scope`
    - `semantic_interpretation_summary.evidence_source`
- Lane remains read-only
- Joined preview output remains deterministic
- Joined preview output must keep `mutation_authorized=false`
- Fail closed if required semantic-interpretation fields are missing or empty
- No apply behavior
- No Trio mutation behavior
- No socket mutation behavior
- Required validation: gate-OFF parity, carry-forward validation of existing slice-02A positives and negatives, carry-forward validation of existing slice-02B positives and negatives, deterministic repeat-run hash equality for semantic-interpretation joined preview, semantic-interpretation negative fail-closed checks, mutation boundary audit, and boundary check confirming no new Trio/socket/apply surfaces were introduced

- Next authorized step after the frozen current scope of `WP20_RUNTIME_SLICE_02C_READONLY_PLAN_BRIDGE_SEMANTIC_INTERPRETATION_INTAKE`: `WP20_RUNTIME_SLICE_02D_READONLY_PLAN_BRIDGE_ISSUES_SUMMARY_INTAKE`
- Objective: intake exactly the `issues_summary` metadata surface from the existing WP-19 projection artifact into the existing read-only plan-bridge path
- Intake limited to:
    - `issues_summary.errors`
    - `issues_summary.warnings`
- Lane remains read-only
- Joined preview output remains deterministic
- Joined preview output must keep `mutation_authorized=false`
- Fail closed if required `issues_summary` fields are missing or malformed
- No apply behavior
- No Trio mutation behavior
- No socket mutation behavior
- Required validation: gate-OFF parity, carry-forward validation of existing slice-02A positives and negatives, carry-forward validation of existing slice-02B positives and negatives, carry-forward validation of existing slice-02C positives and negatives, deterministic repeat-run hash equality for issues-summary joined preview, issues-summary negative fail-closed checks, mutation boundary audit, and boundary check confirming no new Trio/socket/apply surfaces were introduced

- Validated/frozen step after the frozen current scope of `WP20_RUNTIME_SLICE_02D_READONLY_PLAN_BRIDGE_ISSUES_SUMMARY_INTAKE`: `WP20_RUNTIME_SLICE_02E_READONLY_PLAN_BRIDGE_RESOLUTION_PREVIEW`
- Objective: emit a deterministic read-only `resolution_preview` block inside the existing joined plan-bridge preview using only already-authorized projection metadata already consumed by slices 02B through 02D
- Boundary limited to:
    - `status_summary.status`
    - `semantic_interpretation_summary.scope_resolution`
    - `semantic_interpretation_summary.effective_scope`
    - `semantic_interpretation_summary.evidence_source`
- Lane remains read-only
- Joined preview output remains deterministic
- Joined preview output must keep `mutation_authorized=false`
- Fail closed if required resolution-preview inputs are absent after the current projection-intake validation path
- No new upstream projection-contract intake in this step
- No `rule_evaluation_summary` intake in this step
- No apply behavior
- No Trio mutation behavior
- No socket mutation behavior
- Validation status: PASS for gate-OFF parity, carry-forward validation of existing slice-02A through slice-02D positives and negatives, deterministic repeat-run hash equality for the resolution-preview joined preview, resolution-preview block presence and field-order checks, resolution-preview fail-closed checks for missing required inputs, mutation boundary audit, and the boundary check confirming no new Trio/socket/apply surfaces were introduced

- Validated/frozen step after the frozen current scope of `WP20_RUNTIME_SLICE_02E_READONLY_PLAN_BRIDGE_RESOLUTION_PREVIEW`: `WP20_RUNTIME_SLICE_02F_READONLY_PLAN_BRIDGE_RULE_EVALUATION_SUMMARY_INTAKE`
- Objective: intake exactly the `rule_evaluation_summary` metadata surface from the existing WP-19 projection artifact into the existing read-only plan-bridge path
- Intake limited to:
    - `rule_evaluation_summary.phase_order`
    - `rule_evaluation_summary.ordered_rules`
- Lane remains read-only
- Joined preview output remains deterministic
- Joined preview output must keep `mutation_authorized=false`
- Fail closed if required `rule_evaluation_summary` fields are missing or malformed
- No new upstream projection-contract intake in this step
- No apply behavior
- No Trio mutation behavior
- No socket mutation behavior
- No rule-evaluation execution behavior beyond bounded read-only summary intake
- No ordered-rules rendering expansion beyond the bounded read-only summary-intake scope
- Validation status: PASS for gate-OFF parity, carry-forward validation of existing slice-02A through slice-02E positives and negatives, deterministic repeat-run hash equality for the rule-evaluation-summary joined preview, phase-order and ordered-rules stability checks, rule-evaluation-summary negative fail-closed checks, mutation boundary audit, and the boundary check confirming no new Trio/socket/apply surfaces were introduced

- Validated/frozen step after the frozen current scope of `WP20_RUNTIME_SLICE_02F_READONLY_PLAN_BRIDGE_RULE_EVALUATION_SUMMARY_INTAKE`: `WP20_RUNTIME_SLICE_02G_READONLY_PLAN_BRIDGE_RULE_EVALUATION_TRACE_PREVIEW`
- Objective: emit a deterministic read-only `rule_evaluation_trace_preview` block inside the existing joined plan-bridge preview using only already-authorized traceability metadata already present in the existing WP-19 projection artifact and already loaded by slice-02B projection intake
- Boundary limited to:
    - `projection_contract`
    - `projection_kind`
    - `input_artifact`
    - `input_identity.artifact_path`
    - `input_identity.input_fingerprint_sha256`
    - `deterministic_identity_summary.normalized_plan_hash`
    - `deterministic_identity_summary.replay_identity`
    - `deterministic_identity_summary.validator_run_identity`
- Lane remains read-only
- Joined preview output remains deterministic
- Joined preview output must keep `mutation_authorized=false`
- Fail closed if required traceability fields are missing or empty after the existing projection-intake validation path
- No new upstream projection-contract intake or projection-contract expansion in this step
- No rule-execution behavior
- No rule-scoring logic
- No rule-ordering logic
- No rule-interpretation logic
- No rule-filtering logic
- No rule-result derivation or aggregation behavior
- No ordered-rules rendering expansion in this step
- No apply behavior
- No Trio mutation behavior
- No socket mutation behavior
- Validation status: PASS for gate-OFF parity, carry-forward validation of existing slice-02A through slice-02F positives and negatives, deterministic repeat-run hash equality for the rule-evaluation-trace joined preview, trace-preview block presence and field-order checks, traceability-field malformed-input fail-closed checks, mutation boundary audit, and the boundary check confirming no new Trio/socket/apply surfaces were introduced

- Distinct-slice check after the frozen current scope of `WP20_RUNTIME_SLICE_02G_READONLY_PLAN_BRIDGE_RULE_EVALUATION_TRACE_PREVIEW`: `WP20_RUNTIME_SLICE_02H_READONLY_PLAN_BRIDGE_DETERMINISTIC_IDENTITY_PREVIEW` is not a justified next slice under current repo truth
- Reason: `projection_metadata.deterministic_identity_summary` already exposes the bounded deterministic-identity metadata surface, and validated `rule_evaluation_trace_preview` already serializes that same `deterministic_identity_summary` surface together with the surrounding traceability metadata already defined by the existing WP-19 projection/adapter contracts
- Runtime lane posture after 02G: HOLD pending a future non-redundant bounded slice definition
- No new next-step runtime slice is added in this pass
- Runtime continuation marker: `RUNTIME_CONTINUATION_PASS_01` completed as a stabilization-only hardening pass on existing slice-01 through 02G outputs (deterministic ordering + fail-closed ordering-input handling), with no new slice, no schema change, and no mutation/apply authorization change
- Runtime continuation marker: `RUNTIME_CONTINUATION_PASS_02` completed as a stabilization-only hardening pass on existing slice-01 through 02G outputs (evidence completeness + structural consistency fail-closed enforcement), with no new slice, no schema change, and no mutation/apply authorization change

### Definition of Done

- runtime bridge contract documented with explicit guardrails
- bridge activation criteria tied to WP-15 through WP-19 stability evidence
- integration sequencing approved with deterministic regression requirements
- no unvalidated runtime coupling introduced

Strategic sequencing note:
- strategic target remains `Stat Query -> Deterministic Execution Plan`
- semantic viewer capability now exists and changes optimal milestone order
- runtime bridging is intentionally deferred until inspection/explainability/capture/validation are stable
- runtime bridge/execution sequencing is out-of-lane for current `feature/semantic-layer` work until separately approved and kicked off
- WP-18 closeout does not activate WP-20 and does not imply runtime bridge behavior
- WP-19 closeout does not activate WP-20 and does not imply runtime bridge behavior

---

# Long-Term (`v4.2+ / v5.0`)

- plan schema formalization
- internal model extraction for GUI control
- structured test harness automation
- config validation engine

---

Release discipline enforced starting 2026-02-28.
