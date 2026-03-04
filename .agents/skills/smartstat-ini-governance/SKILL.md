---
name: smartstat-ini-governance
description: Use this skill when validating or editing SmartStat INI contracts (TemplateConfig, Mappings, StaticOverrides). Enforce key ordering rules, detect duplicates, and normalize mapping keys (spaces->underscores) as required by SmartStat.
---

# SmartStat INI Governance

## When to use
- Any change to `SmartStat_TemplateConfig.ini`, `SmartStat_Mappings*.ini`, `SmartStat_StaticOverrides.ini`
- Any request like: "validate INI ordering", "check duplicates", "normalize keys"

## Hard rules
- Preserve key order within existing sections.
- Do not reorder existing keys.
- Detect duplicate keys inside a section (fail closed).
- TemplateConfig `[TEMPLATE:<name>]` sections must include keys in this order when present:
  1) config_id
  2) entity
  3) qualifier
  4) filter_tabfields
  5) category_tabfields
  6) row_limit
  7) output_map

## Scripts (PowerShell)
- `scripts/validate_templateconfig_order.ps1`
- `scripts/validate_ini_no_duplicate_keys.ps1`
- `scripts/normalize_mapping_key.ps1`

## Expected outputs
- A clear pass/fail report.
- If failing: list the section and the exact offending lines/keys.
- No automatic rewrites unless explicitly asked.
