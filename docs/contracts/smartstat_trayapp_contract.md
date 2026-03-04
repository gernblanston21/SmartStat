# SmartStat Core <-> SmartStatTrayApp Contract

Contract version: `1.0.0`  
Status: `RC-safe preparation (docs/tests/tooling only)`  
Last updated: `2026-03-04`

## Purpose

This document defines the versioned integration contract between:

- SmartStat core engine (VBScript + INI inputs)
- SmartStatTrayApp (operator-side tooling that reads config and applies outputs)

RC rule: this contract package does not change SmartStat runtime behavior. It only formalizes expectations and adds validators.

## Grounding (Viz Trio docs)

This contract is constrained by:

- `docs/viz-trio/environment_constraints.md` -> "What live-safe means", "Do not break the operator rules"
- `docs/viz-trio/command_reference.md` -> page get/set patterns and tabfield custom property command patterns
- `docs/viz-trio/commands_full_index.md` -> command category grounding (`Page`, `Tabfield`, `Gui`, `Script`)
- `docs/viz-trio/tabfields.md` -> tabfield naming heuristics and custom property usage
- `docs/viz-trio/page_list.md` -> operator read/take workflow constraints
- `docs/viz-trio/page_editor.md` -> editor workflow and save/read behavior constraints
- `docs/viz-trio/show_control.md` -> show-level operational context

If behavior is unsupported by those docs, fail closed.

## Versioning And Compatibility

SemVer policy:

- Major (`X.0.0`): breaking contract changes (TrayApp update required)
- Minor (`0.X.0`): backward-compatible additive changes
- Patch (`0.0.X`): clarifications/fixes with no semantic change

Compatibility rules:

- TrayApp must declare the contract version it implements.
- SmartStat + TrayApp are compatible when major versions match and TrayApp supports the contract minor/patch in use.
- Unknown contract major version -> TrayApp must refuse apply.

## Contract Inputs (TrayApp reads)

Required:

- `SmartStat_TemplateConfig.ini`
- `SmartStat_Mappings.ini`
- `SmartStat_StaticOverrides.ini`

Optional (only when learn-aware workflows are enabled):

- `SmartStat_Mappings.learn.ini`
- `SmartStat_MappingsNBA.learn.ini`
- `SmartStat_MappingsNHL.learn.ini`

Input mode is read-only. TrayApp must not mutate these files during validation/apply.

## Contract Outputs (TrayApp writes/applies)

TrayApp may apply only explicit, deterministic outputs:

- Resolved `config_id` selection for the active `[TEMPLATE:<name>]` contract block
- Computed `output_map` preview for operator confirmation/logging
- Final tabfield value set operations derived from `output_map`
- Optional tabfield custom properties for traceability (only when supported by integration command set), for example:
  - `smartstat.config_id`
  - `smartstat.output_map_preview`
  - `smartstat.contract_version`

No implicit write targets are allowed.

## Template Block Contract

Section header format:

- `[TEMPLATE:<name>]`
- `<name>` must be a non-empty token of letters, digits, and underscore

Required keys in exact order for each template block:

1. `config_id`
2. `qualifier`
3. `filter_tabfields`
4. `category_tabfields`
5. `row_limit`
6. `output_map`

No unknown keys are permitted inside `[TEMPLATE:<name>]` blocks for this contract version.

### Key Definitions

`config_id`

- Type: string token
- Allowed: non-empty `[A-Za-z0-9_.-]+` (examples: `default`, `alt`)
- Purpose: identifies a selectable contract variant for same template name

`qualifier`

- Type: tabfield token or sentinel
- Allowed: `none` or tabfield id like `B0200`, `H0020`
- Purpose: the qualifier input field used by resolver/filter logic

`filter_tabfields`

- Type: comma list or sentinel
- Allowed: `none` or comma-separated tabfield ids (`H0100,H0200`)
- Purpose: filter operand source fields

`category_tabfields`

- Type: comma list
- Allowed: comma-separated tabfield ids (`H1101,H1201,...`)
- Purpose: category selector source fields

`row_limit`

- Type: two-token tuple
- Allowed format: `<token>,<positive-int>`
- `<token>` may be:
  - `none`
  - tabfield id (`C0000`)
  - tabfield id with suffix (`C0000-NumRows`)
- Examples: `none,1`, `C0000,7`, `C0000-NumRows,8`

`output_map`

- Type: comma list of mapping entries
- Allowed:
  - empty value (explicitly allowed)
  - or entries formatted as `<tabfield>:<column>:<row>`
- Example: `H1110:1:1,H1120:1:2,H1210:2:1`

## Parsing Rules

- Comments: full-line comments beginning with `;`, `#`, or `'` are ignored.
- Whitespace: leading/trailing whitespace around keys/values is ignored.
- BOM: UTF-8 BOM is allowed; no-BOM is also allowed.
- Newlines: CRLF is preferred; LF is tolerated; mixed line endings should be flagged as hygiene warnings.
- Duplicate keys: forbidden within a section.
- Empty values:
  - Allowed only for `output_map` in template sections.
  - Any other required template key with empty value is invalid.

## Fail-Closed Rules (TrayApp Must Refuse Apply)

TrayApp must refuse apply when any of these occur:

- Missing required key in a template block
- Wrong required-key order
- Duplicate key in relevant section
- Malformed section header or parse failure
- Empty value in non-empty-required template keys
- Unknown key in `[TEMPLATE:<name>]`
- Duplicate `(template_name, config_id)` pair
- Contract version mismatch (unsupported major)

On refusal:

- Do not write tabfield values.
- Do not emit partial applies.
- Emit clear operator-facing reason and log details.

## Explicit Invariants

- No silent precedence: when multiple candidates conflict, behavior must be explicit and deterministic.
- No implicit ordering: iteration/file load order must not define winner unless contract specifies it.
- Determinism required: identical inputs + same contract version => identical validation/apply decision and output plan.
- Live-safe behavior: non-blocking UX and fail-closed behavior are mandatory.

## TrayApp Dev Checklist

- Read contract version and assert compatibility before parsing.
- Parse INI files read-only with duplicate/format checks enabled.
- Validate each `[TEMPLATE:<name>]` block for required keys and exact order.
- Reject unknown template keys for contract `1.0.0`.
- Enforce empty-value policy (`output_map` only).
- Ensure `(template_name, config_id)` uniqueness.
- Build deterministic output plan before any apply.
- On any contract error, fail closed and do not write.
- Log contract version, template, config_id, and refusal reason for traceability.

## Worked Examples

### Example A (Simple, empty output_map allowed)

```ini
[TEMPLATE:CG_TICKER_PLAYERSTATS]
config_id=default
qualifier=B0300
filter_tabfields=none
category_tabfields=H0100,H0200,H0300,H0400
row_limit=none,1
output_map=
```

### Example B (Multi-field with populated output_map)

```ini
[TEMPLATE:SL_PLYRSTATS3COL]
config_id=default
qualifier=B0210
filter_tabfields=H0110,H0120,H0130
category_tabfields=H1101,H1201,H1301,H1401,H1501,H1601,H1701
row_limit=C0000,7
output_map=H1110:1:1,H1120:1:2,H1130:1:3,H1210:2:1,H1220:2:2,H1230:2:3,H1310:3:1,H1320:3:2,H1330:3:3
```

## Change Control

Any behavior-affecting contract update must:

- bump contract version appropriately
- update validators under `tests/wp-14/contract-validators/`
- include regression evidence under `tests/wp-14/.../artifacts/`
