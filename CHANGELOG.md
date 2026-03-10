# Changelog

All notable changes to the SmartStat Core Engine are documented in this file.
Runtime versioning follows script identifiers (`SmartStat_v*.vbs`), while `VERSION.txt` remains VIZOR UI display metadata.
## [feature/semantic-layer] - 2026-03-10 (WP-20 Target-05 Implementation-Authorization Decision Gate)
### Summary
- Defined the final WP-20 implementation-authorization decision gate artifacts without starting runtime-bridge implementation.
- Added a formal implementation-authorization record template and formal implementation branch-approval record template with required decision inputs, outcomes, and ownership/revocation expectations.
- Reaffirmed this pass is governance/docs/tests only and WP-20 implementation remains NOT STARTED.

### Added
- `docs/onair/wp20_implementation_authorization_record.md`
- `docs/onair/wp20_branch_approval_record.md`
- `tests/wp-20/target-05/README.md`
- `tests/wp-20/target-05/artifacts/.gitkeep`

### Changed
- `ROADMAP.md`
- `SESSION.md`
- `AGENTS.md`
- `tests/wp-20/README.md`
- `CHANGELOG.md`

### Validation
- Governance/docs/tests pass only.
- Protected files remained unchanged (`SmartStat_v4.0.0_beta.vbs`, production INIs, capture schema, validation contract).
## [feature/semantic-layer] - 2026-03-10 (WP-20 Target-04 Pre-Implementation Sign-Off Gate)
### Summary
- Defined the formal WP-20 pre-implementation sign-off gate for transitioning from rehearsal-ready to implementation-ready without starting runtime-bridge code.
- Added a formal rehearsal manifest template and a formal gate-review checklist with required roles, review surfaces, and gate outcomes (`implementation_ready`, `hold`, `fail`).
- Reaffirmed WP-20 implementation remains NOT STARTED in this pass.

### Added
- `docs/onair/wp20_rehearsal_manifest_template.md`
- `docs/onair/wp20_gate_review_checklist.md`
- `tests/wp-20/target-04/README.md`
- `tests/wp-20/target-04/artifacts/.gitkeep`

### Changed
- `ROADMAP.md`
- `SESSION.md`
- `AGENTS.md`
- `tests/wp-20/README.md`
- `CHANGELOG.md`

### Validation
- Governance/docs/tests pass only.
- Protected files remained unchanged (`SmartStat_v4.0.0_beta.vbs`, production INIs, capture schema, validation contract).
## [feature/semantic-layer] - 2026-03-10 (WP-20 Target-03 Pre-Implementation Rehearsal Gate)
### Summary
- Defined the WP-20 pre-implementation rehearsal protocol for future prototype dry-runs without starting runtime-bridge code.
- Defined checkpoint-by-checkpoint rehearsal flow, required artifact pack structure, evidence naming/location rules, and fail/abort handling.
- Reaffirmed WP-20 implementation remains NOT STARTED in this pass.

### Added
- `docs/onair/wp20_rehearsal_protocol.md`
- `tests/wp-20/target-03/README.md`
- `tests/wp-20/target-03/artifacts/.gitkeep`

### Changed
- `ROADMAP.md`
- `SESSION.md`
- `AGENTS.md`
- `tests/wp-20/README.md`
- `CHANGELOG.md`

### Validation
- Governance/docs/tests pass only.
- Protected files remained unchanged (`SmartStat_v4.0.0_beta.vbs`, production INIs, capture schema, validation contract).


## [feature/semantic-layer] - 2026-03-10 (WP-20 Target-02 Governance Gate)
### Summary
- Defined the concrete WP-20 runtime-bridge lane charter and pre-approved regression/evidence execution plan for a future implementation lane.
- Documented exact allowed and forbidden implementation surface categories, rollback evidence requirements, and prototype merge-readiness criteria.
- Reaffirmed WP-20 implementation remains NOT STARTED in this pass.

### Added
- `docs/onair/wp20_lane_charter.md`
- `docs/onair/wp20_regression_evidence_plan.md`
- `tests/wp-20/target-02/README.md`
- `tests/wp-20/target-02/artifacts/.gitkeep`

### Changed
- `ROADMAP.md`
- `SESSION.md`
- `AGENTS.md`
- `tests/wp-20/README.md`
- `CHANGELOG.md`

### Validation
- Governance/docs/tests pass only.
- Protected files remained unchanged (`SmartStat_v4.0.0_beta.vbs`, production INIs, capture schema, validation contract).


## [feature/semantic-layer] - 2026-03-10 (WP-20 Target-01 Approval-Evidence)
### Summary
- Defined the WP-20 Target-01 approval-evidence package required before any runtime-bridge implementation pass may begin.
- Documented approval prerequisites, required evidence categories, rollback criteria, abort criteria, and minimum prototype acceptance gates.
- Reaffirmed this pass is governance/docs/tests only and does not start WP-20 implementation.

### Added
- `docs/onair/wp20_approval_requirements.md`
- `tests/wp-20/README.md`
- `tests/wp-20/target-01/README.md`
- `tests/wp-20/target-01/approval_evidence_template.md`
- `tests/wp-20/target-01/artifacts/.gitkeep`

### Changed
- `ROADMAP.md`
- `SESSION.md`
- `AGENTS.md`
- `CHANGELOG.md`

### Validation
- Governance/docs/tests pass only.
- Protected files remained unchanged (`SmartStat_v4.0.0_beta.vbs`, production INIs, capture schema, validation contract).


## [feature/semantic-layer] - 2026-03-10 (WP-20 Kickoff Planning Gate)
### Summary
- Defined the WP-20 kickoff gate as governance-only and explicitly separate from WP-17/WP-18/WP-19 read-only architecture work.
- Documented allowed upstream read-only inputs, forbidden pre-implementation behaviors, and explicit approval conditions required before any WP-20 implementation can begin.
- Reaffirmed that WP-19 closeout does not activate WP-20 implementation.

### Added
- `docs/onair/wp20_kickoff_checklist.md`

### Changed
- `ROADMAP.md`
- `SESSION.md`
- `AGENTS.md`
- `CHANGELOG.md`

### Validation
- Governance-only pass: no runtime bridge/apply/viewer implementation changes.
- Protected files remained unchanged (`SmartStat_v4.0.0_beta.vbs`, production INIs, capture schema, validation contract).


## [feature/semantic-layer] - 2026-03-09 (WP-19 Closeout / Acceptance)
### Summary
- Closed WP-19 as a complete read-only viewer-contract package over WP-17/WP-18 artifacts.
- Consolidated final acceptance truth for projection contract, projection summary hardening, consumption contract, and projection-to-view-model adapter contract surfaces.
- Reaffirmed WP-20 remains deferred and not implied by WP-19 closeout.

### Added
- `tests/wp-19/ACCEPTANCE.md`

### Changed
- `ROADMAP.md`
- `SESSION.md`
- `CHANGELOG.md`
- `docs/onair/plan-viewer-contract.md`
- `tests/wp-19/README.md`

### Validation
- Ran `python tests/wp-19/harness/read_only_intake_contract_test.py`.
- Ran `python tests/wp-19/harness/viewer_projection_contract_test.py`.
- Ran `python tests/wp-19/harness/projection_consumption_contract_test.py`.
- Ran `python tests/wp-19/harness/projection_to_view_model_adapter_contract_test.py`.

## [feature/semantic-layer] - 2026-03-09 (WP-19 Target-06 Read-Only Projection-to-View-Model Adapter Contract)
### Summary
- Defined a strict read-only projection-to-view-model adapter contract surface for future WP-19 viewer implementation handoff.
- Added deterministic mapping rules that separate projection fields, adapter rules, and future view-model sections.
- Added adapter contract harness assertions for stable shape, deterministic mapping, and read-only non-goal boundaries.

### Added
- `tests/wp-19/harness/projection_to_view_model_adapter_contract_test.py`
- `tests/wp-19/target-06/README.md`
- `tests/wp-19/target-06/artifacts/.gitkeep`
- `tests/wp-19/target-06/artifacts/projection_to_view_model_adapter_contract_test_output.txt`

### Changed
- `docs/onair/plan-viewer-contract.md`
- `tests/wp-19/README.md`
- `tests/wp-19/harness/README.md`

### Validation
- Ran `python tests/wp-19/harness/read_only_intake_contract_test.py`.
- Ran `python tests/wp-19/harness/viewer_projection_contract_test.py`.
- Ran `python tests/wp-19/harness/projection_consumption_contract_test.py`.
- Ran `python tests/wp-19/harness/projection_to_view_model_adapter_contract_test.py`.

## [feature/semantic-layer] - 2026-03-09 (WP-19 Target-05 Projection Consumption Contract Consolidation)
### Summary
- Consolidated WP-19 projection consumption contract expectations into a clear read-only implementation handoff surface.
- Defined authoritative projection fields and explicit display/summary versus traceability surface boundaries.
- Added a Target-05 handoff harness asserting stable contract shape, deterministic ordering, and explicit non-goal boundaries.

### Added
- `tests/wp-19/harness/projection_consumption_contract_test.py`
- `tests/wp-19/target-05/README.md`
- `tests/wp-19/target-05/artifacts/.gitkeep`

### Changed
- `docs/onair/plan-viewer-contract.md`
- `tests/wp-19/README.md`
- `tests/wp-19/harness/README.md`

### Validation
- Ran `python tests/wp-19/harness/read_only_intake_contract_test.py`.
- Ran `python tests/wp-19/harness/viewer_projection_contract_test.py`.
- Ran `python tests/wp-19/harness/projection_consumption_contract_test.py`.

## [feature/semantic-layer] - 2026-03-09 (WP-19 Target-04 Projection Summary Contract Hardening)
### Summary
- Hardened WP-19 projection summary contract shape with explicit top-level and nested key-order assertions.
- Strengthened viewer projection harness checks for deterministic projection serialization, summary section consistency, and read-only non-mutation boundaries.
- Added Target-04 docs/tests scaffold for projection summary contract evidence.

### Added
- `tests/wp-19/target-04/README.md`
- `tests/wp-19/target-04/artifacts/.gitkeep`
- `tests/wp-19/target-04/artifacts/viewer_projection_contract_test_output.txt`
- `tests/wp-19/target-04/artifacts/read_only_intake_contract_test_output.txt`

### Changed
- `docs/onair/plan-viewer-contract.md`
- `tests/wp-19/harness/viewer_projection_contract_test.py`
- `tests/wp-19/target-03/fixtures/projection_pass_case.json`
- `tests/wp-19/target-03/fixtures/projection_refuse_case.json`
- `tests/wp-19/target-03/fixtures/projection_stats_implicit_default_scope_case.json`
- `tests/wp-19/target-03/fixtures/projection_stats_explicit_scope_case.json`
- `tests/wp-19/README.md`
- `tests/wp-19/harness/README.md`

### Validation
- Ran `python tests/wp-19/harness/viewer_projection_contract_test.py`.
- Ran `python tests/wp-19/harness/read_only_intake_contract_test.py`.

## [feature/semantic-layer] - 2026-03-09 (WP-19 Target-03 Viewer Projection Contract Fixtures)
### Summary
- Added WP-19 Target-03 read-only viewer projection contract fixtures for PASS, REFUSE, implicit-default scope, and explicit scope cases.
- Added deterministic projection harness checks for stable shape, ordering preservation, read-only non-mutation, and no runtime/apply/bridge fields.

### Added
- `tests/wp-19/harness/viewer_projection_contract_test.py`
- `tests/wp-19/target-03/README.md`
- `tests/wp-19/target-03/fixtures/projection_pass_case.json`
- `tests/wp-19/target-03/fixtures/projection_refuse_case.json`
- `tests/wp-19/target-03/fixtures/projection_stats_implicit_default_scope_case.json`
- `tests/wp-19/target-03/fixtures/projection_stats_explicit_scope_case.json`
- `tests/wp-19/target-03/artifacts/.gitkeep`
- `tests/wp-19/target-03/artifacts/viewer_projection_contract_test_output.txt`

### Changed
- `tests/wp-19/README.md`
- `tests/wp-19/harness/README.md`
- `docs/onair/plan-viewer-contract.md`

### Validation
- Ran `python tests/wp-19/harness/viewer_projection_contract_test.py`.
- Re-ran `python tests/wp-19/harness/read_only_intake_contract_test.py`.
## [feature/semantic-layer] - 2026-03-09 (WP-19 Target-02 Read-Only Intake Harness)
### Summary
- Added WP-19 read-only intake contract tests/harness for WP-17/WP-18 artifact consumption.
- Added contract-shape assertions, deterministic ordering preservation checks, and no-mutation assertions.
- Added boundary assertions that reject runtime/apply/bridge call surfaces in WP-19 harness scope.

### Added
- `tests/wp-19/harness/read_only_intake_contract_test.py`
- `tests/wp-19/target-02/README.md`
- `tests/wp-19/target-02/artifacts/.gitkeep`

### Changed
- `tests/wp-19/README.md`
- `tests/wp-19/harness/README.md`

### Validation
- Ran `python tests/wp-19/harness/read_only_intake_contract_test.py` (read-only contract test harness).
## [feature/semantic-layer] - 2026-03-09 (WP-19 Target-01 Contract Scaffolding)
### Summary
- Added WP-19 read-only plan-viewer contract scaffolding (docs/tests/tooling only).
- Defined accepted WP-17/WP-18 inputs and minimal read-only view-model expectations.
- Reconfirmed forbidden behavior boundaries (no runtime/apply/bridge behavior and no artifact mutation).

### Added
- `docs/onair/plan-viewer-contract.md`
- `tests/wp-19/README.md`
- `tests/wp-19/harness/README.md`
- `tests/wp-19/target-01/README.md`
- `tests/wp-19/artifacts/.gitkeep`
- `tests/wp-19/target-01/artifacts/.gitkeep`

### Validation
- Governance/docs/tests scaffolding pass only; no SmartStat runtime, production INI, WP-17 schema, or WP-18 validator logic changes.
## [feature/semantic-layer] - 2026-03-09 (WP-19 Kickoff Planning)
### Summary
- Defined WP-19 kickoff gate as a read-only consumer layer over WP-17/WP-18 artifacts (governance/docs only).
- Specified allowed WP-18 inputs for WP-19 planning: validation_result payloads, rule evaluations, refusal diagnostics, deterministic identities, and semantic interpretation metadata.
- Specified forbidden WP-19 behavior: no runtime/apply/bridge behavior, no SmartStat engine calls, no artifact mutation.
- Reaffirmed WP-20 remains deferred and is not implied by WP-19 kickoff planning.

### Changed
- Updated governance truth in `ROADMAP.md`, `SESSION.md`, and `AGENTS.md` for WP-19 kickoff boundaries and status.

### Validation
- Governance-only pass: no validator logic changes, no runtime script changes, no production INI changes, no schema changes.
## [feature/semantic-layer] - 2026-03-09 (WP-18 Closeout / Acceptance)
### Summary
- Closed WP-18 as a complete validation-layer package (docs/tests/tooling only).
- Consolidated acceptance evidence for structural, semantic, determinism, and boundary rule layers.
- Hardened the WP-18 result model with deterministic identity fields and semantic interpretation metadata.
- Preserved strict runtime independence (no SmartStat runtime/apply behavior changes).

### Changed
- Governance truth alignment updates in roadmap/session docs to mark WP-18 CLOSED and keep WP-19/WP-20 deferred.
- WP-18 acceptance documentation added/updated:
  - tests/wp-18/README.md
  - tests/wp-18/validator/validation_result_model.md
  - tests/wp-18/ACCEPTANCE.md
- WP-18 closeout evidence consolidated under tests/wp-18/artifacts/wp18_validator_runs/target06/.

### Validation
- validator_runner.py good fixtures: pass.
- validator_runner.py bad fixtures: refuse as expected.
- deterministic_replay_test.py: pass.
- determinism_rule_layer_test.py: pass.
- boundary_rule_layer_test.py: pass.
- result_model_hardening_test.py: pass.
- Canonical scope verification pass, including:
  - omitted stats scope interpreted as implicit/default career
  - explicit career and explicit season retained as explicit interpretation metadata.
## [feature/semantic-layer] - 2026-03-08 (Governance Re-Baseline)
### Summary
- Re-baselined roadmap/branch governance to lock the repo into one active lane before WP-18 kickoff.
- Clarified active vs frozen vs deferred scope boundaries across roadmap/session/governance docs.
### Changed
- Active lane explicitly locked to `feature/semantic-layer` semantic architecture/tooling work only.
- Runtime/core baseline explicitly frozen (`v4.0.0_beta`, `v4.0.0_RC1`) with no implied runtime bridge/apply activation.
- WP-18, WP-19, and WP-20 explicitly deferred until explicit kickoff.
- Branch policy now states runtime bridge/execution work requires explicit approval and may require a separate dedicated branch.
### Validation
- Governance-only pass: no SmartStat runtime script, production INI, validator, schema, or fixture mutations.
## [feature/semantic-layer] - 2026-03-08 (WP-17)

### Summary
- Implemented WP-17 as a standalone plan-capture contract layer (docs/tests/tooling only).
- Added versioned contract + schema for deterministic captured-plan artifacts.
- Added WP-17 validator harness with fail-closed refusal semantics and deterministic replay hash checks.
- Preserved runtime boundaries: no SmartStat core behavior changes, no planner execution, no runtime bridge/apply behavior.

### Added
- `docs/onair/plan-capture-contract.md`
- `docs/onair/plan-capture.schema.json`
- `tests/wp-17/contract-validators/common_plan_capture_validator.ps1`
- `tests/wp-17/contract-validators/validate_plan_capture_contract.ps1`
- `tests/wp-17/run_wp17.ps1`
- `tests/wp-17/README.md`
- WP-17 good/bad deterministic fixtures under `tests/wp-17/fixtures/`

### Changed
- OnAir architecture docs now explicitly route to WP-17 enforceable contract surface.
- Roadmap/session truth now marks WP-17 contract-layer closeout and WP-18+ deferrals.

### Validation
- `tests/wp-17/run_wp17.ps1` passed (`RUN_PASS=True`, `DETERMINISM_REPLAY_PASS=True`).
- Existing semantic source-view test/build baselines rerun and passed.
- Confirmed no SmartStat runtime script changes in this pass.
## [feature/semantic-layer] - 2026-03-08

### Summary
- Recovery pass for contract/tooling/governance alignment without SmartStat runtime behavior changes.
- Restored WP-14 contract validator reliability against current repo TemplateConfig.
- Extended WP-14 default repo coverage to include sport-aware mappings used by runtime selection logic.
- Updated roadmap/session/doc truth to reflect Phase 9 semantic-source-view state and current test counts.

### Changed
- WP-14 TemplateConfig validator token policy now accepts runtime-evidenced numeric tab tokens (for example `0500`, `1101`) alongside alpha-prefixed tokens.
- WP-14 runner now validates:
  - `SmartStat_Mappings.ini`
  - `SmartStat_MappingsNBA.ini`
  - `SmartStat_MappingsNHL.ini`
  - `SmartStat_Mappings.learn.ini`
  - `SmartStat_MappingsNBA.learn.ini`
  - `SmartStat_MappingsNHL.learn.ini`
- Contract docs now align with validator/runtime token policy and sport-aware mapping coverage.
- Viz Trio docs now explicitly track unsupported-but-used runtime command surfaces and observed socket payload risk notes.

### Validation
- `tests/wp-14/run_wp14.ps1` rerun with recovery label and passed (`RUN_PASS=True`, `REPO_GATE_FAILURES=0`).
- Required deterministic regression packs rerun (`wp-11`, `wp-10 phase-4`) and passed with no core runtime behavior changes introduced.

## [feature/semantic-layer] - 2026-03-07

### Summary
- Completed Semantic Source View Phase 5 through Phase 8 as read-only semantic tooling.
- Established deterministic search explainability baseline and locked search-ranking behavior as a reusable contract with tests.
- Added a validated semantic resolution explainability scaffold as the Phase 8 bridge from semantic inspection to future plan capture.
- Kept scope tooling-only with no SmartStat runtime apply behavior changes.

### Added
- Phase 5: local React Semantic Source View for deterministic inspection of semantic index data (`3101c8d`).
- Phase 6: deterministic search explainability, ranking, debug narratives, and source-hint empty-result guidance (`9ea9977`).
- Phase 7: extracted deterministic search contract module + lightweight regression tests (`30d15b6`).
- Phase 8: read-only semantic resolution explainability scaffold (`55c56ef`).
- Semantic checkpoint tags:
  - `semantic-view-phase7` at `30d15b6`.
  - `semantic-view-phase8-scaffold` at `55c56ef`.

### Changed
- Post-RC architecture sequencing is now semantic-inspection-first in planning docs (WP-15 through WP-20 ordering).
- Runtime bridge sequencing remains intentionally deferred until semantic inspection/explainability/plan capture/plan validation are stable.
- Phase 8 acceptance validated (`npm run build`, `npm run test` with `2` files / `8` tests, plus deterministic render/boundary checks).
## [v4_Dev] - 2026-02-26

### Summary
- Comprehensive hardening pass on top of `v4.0.0_beta`, focused on phase consistency, ambiguity transparency, harness safety checks, output-map reliability, and static-override auditability.
- Source window covered: commits after `v4.0.0_beta` up through `922c583` on `v4_Dev`.

### Added
- Phase pipeline helpers:
  - `Phase_Begin`
  - `Phase_EndOk`
  - `Phase_Fail`
  - `Phase_EarlyExit`
- Ordered phase tracking with diagnostics warnings (`PHASE_ORDER_WARN`) when execution order jumps unexpectedly.
- Harness/integrity framework extensions:
  - `HARNESS`, `HARNESS_COMMIT`, `HARNESS_CAPTURE`, and `HARNESS_STRICT` modes
  - pre-run snapshot capture
  - fixture export (`fixture_*.ini`)
  - integrity diff output (`diff_*.txt`)
  - strict-mode commit block when diff is non-empty
- Output-map completion logic:
  - `BuildEffectiveOutputMap` inference path
  - grouped output candidate matching by prefix/hundred group
  - explicit map entries preserved while missing rows/columns are inferred where safe
- Static override diagnostics:
  - `StaticOverride_GetPrevValue`
  - `Diag_LogOverrideApplied`
  - skip logging path when overrides INI is unavailable (`OVERRIDE_APPLY_SKIP`)
- Tooling scripts for extracting/copying latest script snapshots under `.tools/`.

### Changed
- `ExecuteTemplatePipeline` now uses standardized phase boundaries and failure/early-exit logging across:
  - template classification
  - qualifier/filter detection
  - output-map build
  - syntax build
  - static override apply
- Fail-closed exit points are now explicit and operator-readable:
  - `TEMPLATE.EMPTY`
  - `QUALIFIER.UNRESOLVED`
  - `OUTMAP.EMPTY`
- `ProcessQualifier` normalization now prevents duplicate season-prefix chaining in resolved paths.
- `Stage_ValidatePlan` now has tighter unknown-failure recording and type checks around ambiguity/learn dictionaries.
- Environment/config validation now emits specific missing/unreadable filename diagnostics during startup checks.
- WP-10 Phase-3 Target #1: `Stage_CommitTransaction` now sorts `ApplyPlan` keys before commit using `vbTextCompare` with `vbBinaryCompare` tie-break, eliminating implicit dictionary-order precedence in non-atomic abort paths while preserving write/verify semantics and successful-commit outcomes.
- WP-10 Phase-3 Target #2: resolver tie handling now uses pass-1 winner preservation plus pass-2 top-distance tie detection; tied best candidates fail closed through existing ambiguity paths (`ResolveQualifierSmart`, `ResolveCategorySmart`, and runtime fallback `SuggestQualifierMapping`) while strict non-tie winners remain unchanged.
- WP-10 Phase-3 Target #3: `TRANSFORMS_REGEX` load/apply order is now deterministic (`vbTextCompare` + `vbBinaryCompare` tie-break); exact-pattern conflicting duplicates emit deterministic non-STRICT warnings (`sorted source-key order; later wins`) and fail closed in `HARNESS_STRICT`.
- WP-10 Phase-3 Target #4: AmbiguityContext lifecycle is now deterministic and strict-invariant checked; ambiguity recording always initializes context or fails closed in HARNESS_STRICT; ambiguity summaries are stable-sorted.
- WP-10 Phase-3 Target #5: TryCanonLookupFlexible canonical lookup is now deterministic under normalize-collisions; strict mode fails closed on ambiguous normalize matches.
- WP-10 Phase-3 Target #6: INI normalized-key collisions are now handled deterministically; strict mode fails closed on collisions.
- WP-10 Phase-3 Target #7: unresolved-filter scanning is now deterministic; strict mode fails closed on multi-hit containment matches; non-strict uses stable sorted-first selection with deterministic collision logging.
- WP-10 Phase-3 Target #8: ResolveQualifierSmart now feeds deterministically sorted candidate pools into fuzzy qualifier resolution; outcomes are stable across runs without math/policy changes.
- WP-10 Phase-3 Target #9: ResolveCategorySmart now feeds deterministically sorted candidate pools (and deterministic alias/canon merge where applicable); outcomes are stable across runs without math/policy changes.
- WP-10 Phase-3 Target #10: Heuristic/fuzzy scanner helpers now normalize non-array enumerable candidate ingress to deterministic TEXT_BINARY-sorted arrays before evaluation (ordering-only; duplicates preserved). Learn-only SuggestCanonKey now uses MergeKeysSortedTextBinary for deterministic candidate merge order. No scoring, threshold, tie-rule, ambiguity-policy, or resolver algorithm changes.
- WP-10 Phase-3 Target #11: LoadIniSectionDictNormalized now uses deterministic TEXT_BINARY-sorted section-key ingress for normalize-first population (ordering-only; strict fail-closed unchanged).
- WP-10 Phase-3 complete — Behavioral surface determinism stabilization finalized.
- WP-10 Phase-4: STRICT + non-strict regression verification passed (repeat-run determinism and parity confirmed).
- WP-10 CLOSED.

### Fixed
- Ambiguity context initialization/assignment path in `Main` and `Ambiguity_AddEx` to avoid invalid object-type states.
- `SmartStat_MappingsNBA.learn.ini` malformed section header corrected to `[ALIASES_REGEX]`.
- Duplicate `POINTS/GM` alias entry removed from `SmartStat_MappingsNHL.learn.ini`.
- Additional ambiguity diagnostics stability improvements for top-candidate reporting and summary emission.

### Removed
- Legacy `v3.92` script snapshots from `Main_TrioScript`.
- Legacy `DiagScripts/*` artifacts no longer used by the active `v4` runtime path.

### Compatibility Notes
- Fail-closed ambiguity gating remains strict by default (`allow_ambiguous_apply=false`).
- Viz Trio naming/tabfield conventions are unchanged.
- No blocking UI prompts were introduced; operator messaging remains log/socket based.
- SmartStatTrayApp/socket consumers should continue tolerating multiline ambiguity context details (`AMBIGUITY:` / summary blocks).

### Validation Focus
- Verify harness behavior for all control modes on a known template.
- Verify `OUTMAP.EMPTY` gating still blocks unsafe applies.
- Verify unresolved qualifier paths still hard-stop without partial writes.
- Verify static override logs include source section + old/new values only when values changed.

## [v4.0.0_beta] - 2026-02-24

### Summary
- Initial v4 core engine release with staged transaction writes, strict plan validation, ambiguity capture, and structured diagnostics.

### Added
- Compiler-context state objects:
  - `CompilerContext`
  - `ApplyPlan`
  - `PlanValidationErrors`
- Transactional write path:
  - staged writes via `Tx_SetCustomProp`
  - centralized commit via `Stage_CommitTransaction`
- Structural transaction gate (`Stage_ValidatePlan`) with hard-stop checks for:
  - unresolved qualifier chains
  - ambiguity hits (unless explicitly allowed)
  - empty apply plans
  - unbalanced moustache braces
- Ambiguity subsystem:
  - `Ambiguity_Add` / `Ambiguity_AddEx`
  - top-2 fuzzy-candidate tracking
  - operator-visible ambiguity context in diagnostics/socket messaging
- Sport-aware mappings resolution:
  - `SmartStat_Mappings{SPORT}.ini` and `.learn.ini`
  - fallback support for `Mappings{SPORT}.ini` naming
- Dynamic usage resolver for pitch contexts (`USAGE` -> arsenal/pitch-category percentage paths).

### Changed
- Default apply behavior moved to transactional mode (`TRANSACTION_MODE=True`).
- Qualifier handling became strict fail-safe:
  - blank qualifier defaults to `season`
  - unresolved non-blank qualifiers block apply
- Ambiguity handling became broadcast-safe by default:
  - commits are blocked unless `[LEARN] allow_ambiguous_apply=true`
- League-aware syntax token adjustment applied for NBA/NHL preferred-name behavior.

### Fixed
- UTF-8 BOM protection in INI parse path to avoid first-key corruption.
- Mapping-load failure handling made deterministic so finalize/refresh still runs on fail paths.
- Fuzzy resolution resiliency improved with doubled-letter and adjacent-swap recovery.

### Configuration Notes
- Baseline config family:
  - `SmartStat_Mappings*.ini`
  - `SmartStat_Mappings*.learn.ini`
  - `SmartStat_StaticOverrides.ini`
  - `SmartStat_TemplateConfig.ini`
- Security signature marker currently resides in `SmartStat_TemplateConfig.ini`:
  - `[SECURITY]`
  - `signature=MSSG_FANDUEL_SECURE`

### Release Identity
- Runtime source: `SmartStat_v4.0.0_beta.vbs` (`SMARTSTAT_VERSION="4.0.0_beta"`).
- `VERSION.txt` (`4.0.0`) remains VIZOR UI metadata only.

## [v3.92] - 2025-12-23

### Summary
- Last pre-v4 production line before staged transaction architecture and ambiguity-gating overhaul.
