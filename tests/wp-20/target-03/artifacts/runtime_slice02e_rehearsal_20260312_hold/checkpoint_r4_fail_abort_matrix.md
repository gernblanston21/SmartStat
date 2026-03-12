# Checkpoint R4 - Fail/Abort Matrix

| Trigger | Expected Action | Abort Required |
|---|---|---|
| Scope drift outside read-only resolution-preview | Stop rehearsal and mark fail | yes |
| Protected surface diff on frozen/runtime/config/schema files | Stop rehearsal and mark fail | yes |
| Missing required checkpoint output | Stop rehearsal and mark fail | yes |
| Naming/location rule violation | Stop rehearsal and mark fail | yes |
| Any implied apply/Trio/socket mutation | Stop rehearsal and mark fail | yes |
| Any implied `rule_evaluation_summary` intake | Stop rehearsal and mark fail | yes |
| Any implied new upstream projection-contract intake | Stop rehearsal and mark fail | yes |

## Checkpoint Result

- Checkpoint status: `pass`
- Notes: `Fail/abort decision paths are explicitly defined and remain fail-closed.`
