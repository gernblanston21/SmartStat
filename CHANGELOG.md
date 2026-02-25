# Changelog

All notable changes to the SmartStat Core Engine are documented in this file.
SmartStat runtime versioning follows script identifiers (for this release: `v4.0.0_beta`), while `VERSION.txt` remains VIZOR UI display metadata.

## [v4.0.0_beta]

### Summary
- Introduces a v4 compiler-context execution model with staged transaction apply, structured diagnostics, ambiguity capture, and stricter validation gates.
- Expands multi-sport behavior with sport-aware mappings resolution (`MLB` default, `NBA`/`NHL` variants) and league-aware syntax token adjustments.
- Retains template-driven output-map behavior and static overrides while hardening fail-safe conditions before writes.

### Added
- Compiler-context state objects:
  - `CompilerContext`
  - `ApplyPlan`
  - `PlanValidationErrors`
- Transactional write path:
  - staged writes through `Tx_SetCustomProp`
  - centralized commit phase via `Stage_CommitTransaction`
- Structural validation gate (`Stage_ValidatePlan`) with hard-stop checks for:
  - unresolved qualifiers
  - ambiguous mappings (default block)
  - empty apply plans
  - unbalanced moustache braces in staged syntax
- Ambiguity subsystem:
  - `Ambiguity_Add` / `Ambiguity_AddEx`
  - top-2 fuzzy candidate scoring (`HeuristicPickWithAlt`, `FuzzyResolveTop2Advanced`)
  - operator-visible ambiguity section appended to socket message context (`AMBIGUITY:` block)
- Diagnostics framework:
  - phased markers (`00.BOOT` through `99.DONE`)
  - environment checks (WSH, write access, ADO registry signal)
  - config readability assertions with operator-facing fail messaging
  - bounded log growth via size trimming
- Dynamic `USAGE` resolver for pitch contexts:
  - `pitch_type(...)` -> `arsenal_<pitch_plural>_percentage`
  - `pitch_category(...)` -> `pitch_category_<group>_percentage`
- Sport-aware mapping discovery:
  - resolves `SmartStat_Mappings{SPORT}.ini` and `.learn.ini`
  - supports fallback filename conventions (`Mappings{SPORT}.ini`)

### Changed
- Default apply mode is transactional (`TRANSACTION_MODE = True`) rather than immediate-write behavior.
- Qualifier handling is now strict fail-safe:
  - blank qualifier still defaults to `season`
  - non-blank qualifier must fully resolve; leftovers trigger `QUALIFIER_UNRESOLVED` and block apply
- Ambiguity handling is now broadcast-safe by default:
  - apply is blocked unless `allow_ambiguous_apply=true` in `[LEARN]`
- Fuzzy resolution now tracks alternate near-matches and promotes ambiguity rather than silent guessing when candidate scores are too close.
- League-aware token adjustment now rewrites `{{info.player.preferred_name}}` to `{{info.player.first_name}}` for `NBA` and `NHL`.
- Mapping path resolution now supports league/global-variable and environment-driven selection before generic fallbacks.

### Fixed
- UTF-8 BOM guard in INI parsing to prevent first-line key corruption.
- Added fallback handling on mappings INI load failure so pipeline finalization/validation flow remains deterministic.
- Reduced false positives/negatives in text resolution with:
  - doubled-letter collapse rescue
  - adjacent transposition tolerance for short tokens
  - singularization fallback for qualifier tokens
- Added explicit ambiguity gates to prevent unsafe commits when multiple high-confidence matches exist.

### Configuration and Data
- Baseline config set includes:
  - `SmartStat_Mappings.ini`
  - `SmartStat_MappingsNBA.ini`
  - `SmartStat_MappingsNHL.ini`
  - `SmartStat_Mappings.learn.ini`
  - `SmartStat_MappingsNBA.learn.ini`
  - `SmartStat_MappingsNHL.learn.ini`
  - `SmartStat_StaticOverrides.ini`
  - `SmartStat_TemplateConfig.ini`
- `SmartStat_TemplateConfig.ini` includes 15 template blocks, including `config_id` variants for shared template names.
- `SmartStat_Mappings.ini` includes a security marker section:
  - `[SECURITY]`
  - `signature=MSSG_FANDUEL_SECURE`
- Static override behavior remains entity-specific and template-aware, including pitcher substitution of `{{info.player.primary_position}}` to `{{info.player.pitcher_hand}}` in player-pitcher context.

### Compatibility Notes
- Viz Trio tabfield naming/pattern conventions are unchanged.
- SmartStatTrayApp and other socket-context consumers should tolerate multiline `message_context` values containing appended `AMBIGUITY:` diagnostics.
- Operator workflows that previously relied on permissive fuzzy fallback may now see intentional hard blocks until qualifier/category ambiguity is resolved.

### Known Issues
- `SmartStat_MappingsNBA.learn.ini` contains a malformed section header at line 41:
  - `LIASES_REGEX]`
  - expected bracketed format (e.g., `[ALIASES_REGEX]`)
  This can prevent intended regex-alias parsing for that section.
- `SESSION.md` and `ROADMAP.md` currently describe ambiguity support as missing, which does not match the implemented `v4.0.0_beta` script state.
- Default script/log paths are still anchored to `E:\EDRIVE\UNIVERSAL\SmartStat\...`; non-standard deployment paths depend on fallback discovery and environment alignment.

### Validation Checklist
- Confirm transaction validation blocks apply when:
  - qualifier is unresolved
  - ambiguity exists and `allow_ambiguous_apply=false`
  - staged plan is empty
  - moustache braces are unbalanced
- Validate end-to-end syntax generation on representative templates across MLB, NBA, and NHL mapping profiles.
- Confirm static overrides are applied after transaction commit and that pitcher-hand substitution remains correct.
- Confirm socket refresh payload includes expected tabfield set and ambiguity context when ambiguity is triggered.

### Release Identity
- SmartStat runtime release source: `SmartStat_v4.0.0_beta.vbs` (`SMARTSTAT_VERSION = "4.0.0_beta"`).
- `VERSION.txt` value (`4.0.0`) is VIZOR UI metadata and not runtime version gating.