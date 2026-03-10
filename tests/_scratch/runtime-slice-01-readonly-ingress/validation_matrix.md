# Runtime Slice-1 Validation Matrix

Scope: `WP20_RUNTIME_SLICE_01_READONLY_INGRESS` validation-only pass  
Runtime line: `SmartStat_v4.1.0.vbs`  
Boundary: read-only ingress only

## Fixture Results

| Fixture | Expected Result | Actual Result | Pass/Fail | Evidence |
|---|---|---|---|---|
| `ingress_case01.fixture` (`pos_ingress_run1`) | `status=success`, empty `error_code` | `status=success`, empty `error_code` | PASS | `runs/pos_ingress_run1.json` |
| `ingress_case01.fixture` (`pos_ingress_run2`) | `status=success`, empty `error_code` | `status=success`, empty `error_code` | PASS | `runs/pos_ingress_run2.json` |
| `neg_missing_pagename.fixture` | `status=fail_closed`, `error_code=REQUIRED_INPUT_MISSING` | `status=fail_closed`, `error_code=REQUIRED_INPUT_MISSING` | PASS | `runs/neg_missing_pagename.json` |
| `neg_missing_template.fixture` | `status=fail_closed`, `error_code=REQUIRED_INPUT_MISSING` | `status=fail_closed`, `error_code=REQUIRED_INPUT_MISSING` | PASS | `runs/neg_missing_template.json` |
| `neg_missing_tabfield_list.fixture` | `status=fail_closed`, `error_code=REQUIRED_INPUT_MISSING` | `status=fail_closed`, `error_code=REQUIRED_INPUT_MISSING` | PASS | `runs/neg_missing_tabfield_list.json` |
| `neg_duplicate_tabfields.fixture` | `status=fail_closed`, `error_code=AMBIGUOUS_TABFIELD_LIST` | `status=fail_closed`, `error_code=AMBIGUOUS_TABFIELD_LIST` | PASS | `runs/neg_duplicate_tabfields.json` |
| `neg_unsupported_surface.fixture` | `status=fail_closed`, `error_code=FIXTURE_LOAD_FAILED` | `status=fail_closed`, `error_code=FIXTURE_LOAD_FAILED` | PASS | `runs/neg_unsupported_surface.json` |
| `neg_missing_required_command.fixture` | `status=fail_closed`, `error_code=FIXTURE_COMMAND_MISSING` | `status=fail_closed`, `error_code=FIXTURE_COMMAND_MISSING` | PASS | `runs/neg_missing_required_command.json` |

## Positive Determinism Check

- `runs/pos_ingress_run1.json` SHA256: `E7250A68C63D9A0164BEE40E530FCEE8C0DF10D6176BCA8CEF0887CE559ED7B7`
- `runs/pos_ingress_run2.json` SHA256: `E7250A68C63D9A0164BEE40E530FCEE8C0DF10D6176BCA8CEF0887CE559ED7B7`
- Result: `PASS` (`hash_match=True`)

## Notes

- This matrix is local validation evidence only.
- No live Viz Trio proof is claimed in this file.
