# Changelog

All notable changes to the SmartStat Core Engine are documented in this file.
Runtime versioning follows script identifiers (`SmartStat_v*.vbs`), while `VERSION.txt` remains VIZOR UI display metadata.

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
