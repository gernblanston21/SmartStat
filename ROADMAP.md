# ROADMAP - SmartStat v4 Lifecycle

## Current State
- v4.0.0_beta and v4.0.0_RC1 are frozen historical baselines.
- WP-10 through WP-14 are complete and retained as post-RC stabilization evidence.
- Active working lane in this repo is feature/semantic-layer (v4.2.0 architecture/tooling track).
- Current semantic posture: Phase 9 candidate-resolution scaffold is present in tools/semantic-source-view/.
- WP-17 now defines a versioned, deterministic plan-capture contract layer (docs/tests/tooling only).
- WP-18 is CLOSED as a validation-layer package (docs/tests/tooling only; runtime-independent).
- WP-19 is CLOSED as a read-only viewer-contract package over WP-17/WP-18 artifacts (docs/tests/tooling only; no UI/runtime behavior).
- WP-20 Target-06 authorization-packet fill/verification gate is defined as a separate runtime-bridge lane governance package; implementation is NOT STARTED and is not implied by WP-19 closeout.
## Lane Re-Baseline (2026-03-08)
- Single active lane: `feature/semantic-layer` for semantic architecture/tooling only (docs/tests/read-only tooling).
- Frozen runtime/core baseline: `v4.0.0_beta` and `v4.0.0_RC1` (no implicit runtime execution lane is active).
- WP-18 lane state: CLOSED (2026-03-09) with acceptance evidence under `tests/wp-18/artifacts/wp18_validator_runs/`.
- WP-19 lane state: CLOSED (2026-03-09) with acceptance evidence under `tests/wp-19/`.
- WP-20 lane state: TARGET-06 AUTHORIZATION-PACKET FILL/VERIFICATION GATE DEFINED (2026-03-10; implementation not started; explicit approval required).
- Branch boundary: runtime bridge/execution work requires explicit approval and should run on a separate dedicated branch when started.
- WP-15 through WP-17 closeout does not imply runtime bridge/apply integration.

## Determinism Doctrine (v4+)

SmartStat guarantees (within a given version + config):
- Identical input state + config ? identical output
- Identical STRICT harness run (same version/config) ? identical artifacts
- No implicit precedence via iteration order in resolver logic
- No filesystem-order-dependent behavior

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
- RC1_CHECKLIST.md fully executed
- Tag v4.0.0_RC1 created

---

# Post-RC Tracks (v4.1.0 -> v4.2.0)

## Guardrails (Post-RC)
- RC1 behavior is the baseline; changes must be intentional, scoped, and validated.
- No naming convention changes.
- No INI key reordering.
- Fail-closed semantics preserved (ambiguity + validation gates).
- Each WP must include: scope, DoD, regression plan, and validation artifacts.

## WP-10 (v4.1.0): Determinism Surface Stabilization (Explicit Ordering)

### Objective
Eliminate nondeterministic iteration surfaces that affect:
- Visible output stability
- Log emission stability
- Harness artifact reproducibility
- Resolver decision consistency
- INI traversal consistency

This is a controlled architectural stabilization pass.
No math, resolution logic, or precedence changes are permitted.

---

### Determinism Surface Classes (Mapped)

The following unordered iteration classes are in scope:

1. Dictionary key iteration (runtime collections)
2. Multi-source merge order (config overlays)
3. First-match resolver scans
4. Output_map emission ordering
5. Log/diagnostic dump ordering
6. INI section traversal
7. INI key traversal (if iterative)
8. Filesystem enumeration (if used for config load)
9. List parsing rehydration into unordered containers
10. Candidate set construction during heuristic scans

Each surface must be classified as:
- Behavior-affecting
- Presentation-only
- Non-impacting

HIGH-risk surfaces (behavioral) must be stabilized before emission-level sorting.

---

### Explicit Non-Scope

WP-10 must NOT:
- Change resolution precedence rules
- Alter ambiguity detection behavior
- Modify fail-closed gating
- Reorder INI files
- Refactor resolver algorithm design
- Introduce silent precedence changes
- Mask latent ambiguity bugs via sorting

---

### Implementation Phases

Phase 1 – Surface Audit (no code changes)
- Confirm actual presence of each determinism surface class
- Identify behavioral vs presentation-only cases
- Document implicit precedence dependencies

Phase 2 – Emission Stabilization
- Stabilize output_map ordering
- Stabilize ambiguity dump ordering
- Stabilize harness artifact serialization

Phase 3 – Behavioral Surface Hardening
- Stabilize dictionary iteration used in resolver decisions
- Stabilize overlay merge ordering explicitly
- Stabilize candidate evaluation order

### WP-10 Phase 3 — Behavioral Surface Hardening (Targets)

Completed:
- Target #1 — Commit ordering determinism (ApplyPlan.Keys sorted before non-atomic commit loop).
- Target #2 — Resolver tie determinism (two-pass tie detection; ties fail closed).
- Target #3 — TRANSFORMS_REGEX determinism (sorted load/apply; strict conflict fail-closed).
- Target #4 — AmbiguityContext lifecycle determinism + strict invariants (AMBIGUOUS_CONTEXT_INVALID; stable ambiguity emissions).
- Target #5 — TryCanonLookupFlexible: deterministic first-match behavior (normalize-collision handling).
- Target #6 — LoadIniSectionDictNormalized / LoadIni: deterministic normalized-key collision handling.
- Target #7 — SuggestQualifierMapping / ResolveFilterFragments: deterministic first-hit scanning (containment + fuzzy fallback).
- Target #8 — ResolveQualifierSmart candidate pool ordering determinism.
- Target #9 — ResolveCategorySmart alias/canonical merge determinism.
- Target #10 — Heuristic scanner input normalization

- Target #11 — Residual normalized lookup surfaces (if discovered) (Completed)

Phase-3 complete: All documented HIGH nondeterministic behavioral surfaces stabilized with ordering-only fixes. No ambiguity, scoring, resolver, or logging drift introduced.

Phase 4 – Regression Verification
- STRICT harness repeat-run validation
- Confirm identical resolution outcomes
- Confirm diffs are ordering-only
- Document validation artifact record

WP-10 CLOSED — Determinism surface stabilization and regression verification complete.

---

### Definition of Done (Expanded)

- All HIGH-risk determinism surfaces stabilized
- All MEDIUM-risk surfaces stabilized for artifact consistency
- STRICT harness repeat-run produces identical artifacts
- No new ambiguity leakage
- No resolution outcome changes
- Changelog + validation record committed

---

### Versioning

This is a minor version bump (v4.1.0) because:
- Output ordering will change
- Diff behavior changes
- Determinism guarantees are strengthened

### WP-10 Retrospective — Determinism Stabilization

WP-10 hardened SmartStat’s resolution engine to guarantee deterministic behavior across runs.
All previously identified HIGH-risk nondeterministic surfaces (dictionary iteration order, candidate pool construction, alias/canonical merge order, normalize-first lookup helpers, and heuristic scanner ingress) were stabilized without altering resolver math, scoring rules, or ambiguity policy.

The work focused exclusively on deterministic ordering and fail-closed ambiguity preservation so that identical input state and configuration now produce identical outputs every time.

Phase-4 regression verification validated repeat-run determinism in both STRICT and runtime harness modes, with archived evidence under `/tests/wp-10/phase-4/`.

This milestone establishes SmartStat’s first fully verified deterministic core and provides a stable foundation for future resolver enhancements and feature work.

No patch release permitted for this scope.

## WP-11 (v4.1.0): Harness regression pack framework
### Scope
- Define a repeatable “regression pack” set of templates / scenarios
- Standardize how artifacts are stored and compared

### Definition of Done
- Regression pack documented and runnable
- Artifact locations standardized
- Clear pass/fail criteria captured

WP-11 CLOSED — Harness regression pack framework implemented (tests-only).
Evidence: tests/wp-11/regression-pack/artifacts/compare/wp11_runA__wp11_runB/ (PACK_PASS=True)

## WP-12 (v4.1.0): Enhanced learn system validation
### Scope
- Verify learn file writes are correct, stable, and governed
- Ensure learn updates are validated (format + intent) before acceptance

### Definition of Done
- Learn write validation rules implemented
- Bad/partial writes are blocked or quarantined (fail-closed)
- Validation artifacts recorded

WP-12 CLOSED — Enhanced learn system validation implemented (tests-only).
Evidence: tests/wp-12/learn-validation/artifacts/wp12_runC/ (RUN_PASS=True)

## WP-13 CLOSED (v4.1.0): Resolver performance optimization
### Scope
- Optimize resolver hot paths without changing resolution outcomes
- Preserve determinism and logging semantics

### Definition of Done
- Performance improvement measured
- No behavioral diffs in STRICT harness regression pack

RC note:
- Resolver performance optimization validated via regression harness under `tests/wp-13/resolver-perf/`.
- Optimization itself modifies runtime code paths and is therefore **deferred until post-RC**.
- Implementation change is preserved in local stash `WP-13 perf micro-opt (post-RC)`.

Evidence: tests/wp-13/resolver-perf/ (STRICT regression harness validation)

## WP-14 (v4.1.0): TrayApp alignment preparation
### Scope
- Define and stabilize the contract between SmartStat core + TrayApp
- Ensure mappings/config expectations are explicit and version-safe

### Definition of Done
- Contract documented
- Validator suite implemented
- Harness execution passing
- No SmartStat core behavior changes

WP-14 implementation package (RC-safe, docs/tests/tooling only):
- Contract spec: docs/contracts/smartstat_trayapp_contract.md
- Validator suite: tests/wp-14/contract-validators/
- Harness runner: tests/wp-14/run_wp14.ps1
- Run command: pwsh -NoProfile -ExecutionPolicy Bypass -File tests/wp-14/run_wp14.ps1 -RunLabel <label>
- Evidence location: tests/wp-14/contract-validators/artifacts/<runLabel>/
- RC constraint: no SmartStat core VBScript behavior changes; no production INI schema/order mutations

WP-14 CLOSED - TrayApp contract + validator harness implemented (docs/tests/tooling only).
Evidence: tests/wp-14/contract-validators/artifacts/wp14_runB/ (historical closeout record) and tests/wp-14/contract-validators/artifacts/recovery_wp14_20260308/ (RUN_PASS=True, REPO_GATE_FAILURES=0; expanded sport-aware coverage).

## WP-15 (v4.2.0): Semantic Inspection Foundation
### Scope
- Establish semantic inspection as the first architecture baseline for post-RC planning work.
- Lock the existing Semantic Source View capability as the canonical read-only inspection surface.
- Preserve deterministic search/inspection contracts as reusable inputs for later explainability and planning work.

### Definition of Done
- Semantic inspection baseline documented (records, relationships, query paths, traceability, deterministic search).
- Deterministic inspection contracts are explicit and reusable.
- Read-only inspection workflow is reproducible on `feature/semantic-layer`.
- No SmartStat runtime behavior changes introduced.

WP-15 implementation package (read-only tooling):
- React Semantic Source View baseline: `tools/semantic-source-view/` (Phase 5).
- Deterministic search explainability/ranking: Phase 6.
- Deterministic search contract + lightweight tests: Phase 7.
- Checkpoint tag: `semantic-view-phase7` (`30d15b6`).

WP-15 CLOSED - Semantic inspection foundation accepted (read-only tooling only).

## WP-16 (v4.2.0): Resolution Explainability
### Scope
- Add a read-only explainability scaffold that models deterministic semantic resolution reasoning.
- Keep this as architecture/tooling only (no runtime apply behavior, no planner execution).
- Bridge semantic search inspection outputs to future plan-capture contracts.

### Definition of Done
- Resolution explainability data model/schema defined for semantic tooling.
- Deterministic explainability fixture/scaffold available for local inspection.
- Explainability boundaries documented against search explainability and plan execution.
- No SmartStat runtime behavior changes introduced.

WP-16 implementation package (read-only tooling):
- Explainability guide: `tools/semantic-source-view/RESOLUTION_EXPLAINABILITY_GUIDE.md`.
- Explainability schema: `tools/semantic-source-view/resolution-explainability.schema.json`.
- Deterministic fixture + read-only panel scaffold: `tools/semantic-source-view/src/data/resolutionExplainability.fixture.ts`.
- Scaffold checkpoint tag: `semantic-view-phase8-scaffold` (`55c56ef`).

WP-16 CLOSED - Phase 8 and Phase 9 read-only explainability/candidate-resolution scaffolds are accepted.
WP-16 handoff into WP-17 contract work is complete.

Validation evidence (2026-03-08):
- `npm run build` passed.
- `npm run test` passed (`3` files, `13` tests).
- Baseline UI render, Phase 8 scaffold render, Phase 9 candidate-resolution scaffold render, search-driven selection updates, browse-mode behavior, zero-normalized-result behavior, and deterministic visual repeatability all passed.
- Boundary checks passed: no runtime/apply behavior, no planner execution, and no runtime integration implied.

## WP-17 (v4.2.0): Plan Capture Contract Layer (Docs/Tests/Tooling Only)
### Scope
- Define a versioned captured-plan artifact contract and schema.
- Enforce contract shape via validator tooling and deterministic fixture harness.
- Keep capture layer read-only and explicitly separated from validation, viewer, and runtime execution.

### Definition of Done
- Contract doc + schema published.
- Validator + runner implemented under `tests/wp-17/`.
- Good/bad fixtures pass expected outcomes.
- Deterministic replay hash check passes in runner summary.
- No SmartStat runtime behavior changes introduced.

WP-17 implementation package (read-only docs/tests/tooling):
- Contract doc: `docs/onair/plan-capture-contract.md`
- Contract schema: `docs/onair/plan-capture.schema.json`
- Validator: `tests/wp-17/contract-validators/validate_plan_capture_contract.ps1`
- Harness runner: `tests/wp-17/run_wp17.ps1`
- Run command: `pwsh -NoProfile -ExecutionPolicy Bypass -File tests/wp-17/run_wp17.ps1 -RunLabel <label>`
- Evidence location: `tests/wp-17/contract-validators/artifacts/<runLabel>/`
- Constraint: no SmartStat runtime behavior changes; no planner execution; no runtime bridge/apply behavior.

WP-17 CLOSED - Standalone capture contract layer defined (docs/tests/tooling only).
Evidence: `tests/wp-17/contract-validators/artifacts/wp17_contract_20260308/` (`RUN_PASS=True`, `DETERMINISM_REPLAY_PASS=True`).

## WP-18 (v4.2.0): Plan Validation
Status: CLOSED (accepted on 2026-03-09; validation-layer-only, runtime-independent).

### Scope
- Define the Plan Validation Contract Layer between WP-17 capture and future WP-19/WP-20 layers.
- Validate captured plans using deterministic, read-only architecture rules.
- Keep validator work runtime-independent (no runtime execution, no Trio/apply behavior).

Kickoff governance artifacts:
- `docs/onair/plan-validation-contract.md`
- `docs/onair/wp18_kickoff_checklist.md`

WP-18 explicit boundaries:
- Allowed: captured-plan validation, deterministic rule evaluation, validator tooling, schema compatibility checks.
- Not allowed: runtime execution, Trio integration, SmartStat engine calls, applying stats to graphics, captured-plan mutation.

### Definition of Done
- Plan validator implemented.
- Fixture suite passes expected good/bad cases.
- Harness runner reports RUN_PASS=True for plan validation scope.
- Deterministic replay check is stable across repeated runs.
- No SmartStat runtime behavior changes introduced by validation tooling.

WP-18 implementation package (read-only docs/tests/tooling):
- Validator runner: tests/wp-18/validator/validator_runner.py
- Result-model doc: tests/wp-18/validator/validation_result_model.md
- Validation phase/readme doc: tests/wp-18/README.md
- Replay/harness tests:
  - tests/wp-18/replay/deterministic_replay_test.py
  - tests/wp-18/replay/determinism_rule_layer_test.py
  - tests/wp-18/replay/boundary_rule_layer_test.py
  - tests/wp-18/replay/result_model_hardening_test.py
- Acceptance note: tests/wp-18/ACCEPTANCE.md
- Evidence location: tests/wp-18/artifacts/wp18_validator_runs/target06/

WP-18 CLOSED - Plan validation layer accepted with:
- structural + semantic + determinism + boundary rules
- hardened result model
- deterministic semantic interpretation metadata for default stats scope (omitted scope => career in stats context)
- deterministic replay/harness evidence

WP-19 consumption boundary (allowed):
- Consume WP-18 validation_result payloads as read-only artifacts only.
- Consume rule evaluations, refusal diagnostics, deterministic identities, and semantic interpretation metadata.
- Must not imply runtime apply behavior or runtime bridge activation.
## WP-19 (v4.2.0): Plan Viewer
Status: CLOSED (accepted 2026-03-09; read-only viewer-contract package only).

### Scope
- Extend read-only inspection UX to include deterministic plan-view semantics.
- Provide explainable plan browsing/debugging without apply/runtime integration.
- Keep viewer contracts versioned and compatible with semantic + validation outputs.

### Kickoff Gate (Read-Only Consumer Layer)
WP-19 may consume the following immutable inputs:
- WP-17 captured-plan artifacts.
- WP-18 `validation_result` payloads.
- WP-18 `rule_evaluations`.
- WP-18 refusal diagnostics (`errors`, `warnings`, refusal codes/messages).
- WP-18 deterministic identities (`input_identity`, `normalized_plan_hash`, `replay_identity`, `validator_run_identity`).
- WP-18 `semantic_interpretation` metadata.

### Forbidden Behaviors (Must Not)
WP-19 must not:
- Execute runtime behavior.
- Trigger apply behavior.
- Introduce bridge behavior.
- Call SmartStat engine/apply surfaces.
- Mutate captured-plan artifacts.
- Mutate validation artifacts.

### Minimal Viewer Contract Surfaces (Pre-Implementation)
- Artifact intake surface: accepts WP-17/WP-18 artifacts as read-only inputs.
- Validation summary surface: exposes `status`, deterministic `errors`, and deterministic `warnings`.
- Rule evaluation surface: exposes stable phase/category/rule ordering from `rule_evaluations`.
- Deterministic identity surface: exposes artifact identity and replay-normalization identities for traceability.
- Semantic interpretation surface: exposes `semantic_interpretation` exactly as validation metadata (no runtime inference).

### Definition of Done
- Read-only plan viewer contract documented.
- Read-only projection and adapter contracts documented and validated with deterministic harness evidence.
- Compatibility rules documented across semantic/explainability/plan contracts.
- No runtime apply behavior changes introduced.

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
  - `projection_contract`, `projection_kind`, `input_artifact`, `input_identity`,
    `status_summary`, `issues_summary`, `rule_evaluation_summary`,
    `deterministic_identity_summary`, `semantic_interpretation_summary`
- WP-19 adapter contract surfaces:
  - adapter identity fields plus `view_model` sections:
    `status_view`, `issues_view`, `rules_view`, `semantic_view`, `trace_view`
- Deterministic ordering guarantees established by WP-19 harnesses.
- Explicitly not included:
  - runtime execution, apply behavior, bridge behavior, mutation, or runtime inference.


## WP-20 (v4.2.0+): Runtime Bridge
Status: TARGET-06 AUTHORIZATION-PACKET FILL/VERIFICATION GATE DEFINED (2026-03-10 governance/docs/tests only; implementation NOT STARTED).

### Scope
- Introduce controlled bridge points from validated plan artifacts toward runtime integration.
- Defer runtime bridge implementation until semantic inspection, explainability, plan capture, and plan validation are stable.
- Preserve fail-closed and deterministic discipline while defining bridge constraints.

### Kickoff Gate (Governance Only)
- WP-20 is explicitly separate from WP-17/WP-18/WP-19 read-only architecture packages.
- WP-20 kickoff is planning-only in this pass and does not authorize implementation.
- Gate checklist reference: docs/onair/wp20_kickoff_checklist.md
- Approval requirements reference: docs/onair/wp20_approval_requirements.md
- Target-01 evidence template reference: tests/wp-20/target-01/approval_evidence_template.md
- Target-02 lane charter reference: docs/onair/wp20_lane_charter.md
- Target-02 regression/evidence plan reference: docs/onair/wp20_regression_evidence_plan.md
- Target-03 rehearsal protocol reference: docs/onair/wp20_rehearsal_protocol.md
- Target-03 scaffold reference: tests/wp-20/target-03/README.md
- Target-04 rehearsal manifest template reference: docs/onair/wp20_rehearsal_manifest_template.md
- Target-04 gate review checklist reference: docs/onair/wp20_gate_review_checklist.md
- Target-04 scaffold reference: tests/wp-20/target-04/README.md
- Target-05 implementation-authorization record reference: docs/onair/wp20_implementation_authorization_record.md
- Target-05 branch-approval record reference: docs/onair/wp20_branch_approval_record.md
- Target-05 scaffold reference: tests/wp-20/target-05/README.md
- Target-06 authorization packet index template reference: docs/onair/wp20_authorization_packet_index_template.md
- Target-06 packet completeness checklist reference: docs/onair/wp20_packet_completeness_checklist.md
- Target-06 scaffold reference: tests/wp-20/target-06/README.md

### Allowed Upstream Inputs (Read-Only)
- WP-17 captured-plan artifacts and contract/schema evidence.
- WP-18 validation outputs: validation_result, rule_evaluations, refusal diagnostics, deterministic identities, and semantic_interpretation metadata.
- WP-19 projection/consumption/adapter contract surfaces and harness evidence.

### Forbidden Pre-Implementation Behaviors
- No runtime bridge code paths.
- No apply behavior implementation.
- No Trio integration.
- No SmartStat engine/apply calls.
- No mutation of captured/validated/projected/adapter artifacts.
- No runtime side-effect inference beyond read-only artifacts.

### Required Approval Conditions Before Implementation
- Explicit WP-20 implementation approval recorded in governance docs.
- Dedicated branch/lane confirmed for runtime bridge risk isolation.
- Pre-implementation regression/evidence plan approved (determinism, fail-closed behavior, rollback plan).
- Runtime bridge risk-class review approved before any code mutation.

### Target-01 through Target-06 Governance Outcomes
- Approval prerequisites are defined before any WP-20 code changes.
- Required evidence categories are defined for a future implementation lane.
- Rollback criteria and abort/fail-closed triggers are defined.
- Minimum acceptance gates are defined for any future runtime bridge prototype.
- Concrete lane charter is defined for branch/scope isolation.
- Pre-approved regression/evidence execution structure is defined.
- Pre-implementation rehearsal protocol is defined for dry-run flow validation.
- Rehearsal artifact pack layout, naming rules, checkpoint order, and fail/abort handling are defined.
- Formal rehearsal manifest template is defined for deterministic evidence intake.
- Formal sign-off gate checklist is defined with required roles and review surfaces.
- Gate outcomes are defined: implementation_ready, hold, and fail.
- Minimum evidence requirements are defined before implementation authorization.
- Formal implementation-authorization record template is defined with required decision inputs and outcomes.
- Formal implementation branch-approval record template is defined with branch isolation and ownership fields.
- Authorization and revocation ownership expectations are explicitly defined.
- Gate outcomes are defined: authorized_to_start_implementation, hold, and denied.
- Formal authorization packet index template is defined to bind required pre-implementation inputs.
- Formal packet completeness checklist is defined with verification outcomes: packet_complete, packet_incomplete, and packet_invalid.
- Packet verification is explicitly defined as non-authorizing; separate authorization decision still governs implementation start.
- This pass does not start runtime bridge implementation.
### Branch / Lane Rule
- Runtime bridge/execution implementation must run on a separate dedicated lane from the current read-only semantic lane.
- feature/semantic-layer remains governance/docs/tests/tooling for kickoff gating until explicit implementation approval is granted.

### Definition of Done
- Runtime bridge contract documented with explicit guardrails.
- Bridge activation criteria tied to WP-15 through WP-19 stability evidence.
- Integration sequencing approved with deterministic regression requirements.
- No unvalidated runtime coupling introduced.

Strategic sequencing note:
- Strategic target remains Stat Query -> Deterministic Execution Plan.
- Semantic viewer capability now exists and changes optimal milestone order.
- Runtime bridging is intentionally deferred until inspection/explainability/capture/validation are stable.
- Runtime bridge/execution sequencing is out-of-lane for current feature/semantic-layer work until separately approved and kicked off.
- WP-18 closeout does not activate WP-20 and does not imply runtime bridge behavior.
- WP-19 closeout does not activate WP-20 and does not imply runtime bridge behavior.
---
# Long-Term (v4.2+ / v5.0)

- Plan schema formalization
- Internal model extraction for GUI control
- Structured test harness automation
- Config validation engine

---

Release discipline enforced starting 2026-02-28.
