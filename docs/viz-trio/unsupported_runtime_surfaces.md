# Viz Trio - Unsupported Runtime Surfaces (SmartStat Evidence Notes)

## Purpose

This file documents SmartStat runtime command surfaces that are currently used in repo code but are not explicitly grounded by the current `docs/viz-trio/` command references.

Use this file to keep governance fail-closed and to prevent undocumented assumptions from becoming implicit contract behavior.

## Repo-Observed Commands Without Current In-Repo Command Grounding

Observed in:

- `SmartStat_v4.0.0_beta.vbs:580`
- `SmartStat_v4.0.0_beta.vbs:5110`
- `SmartStat_v4.0.0_beta.vbs:5125`
- `SmartStat_v4.0.0_beta.vbs:5126`

Commands:

- `trio:get_global_variable league`
- `page:getpagedescription`
- `sock:socket_is_connected`
- `sock:send_socket_data`

Current status in `docs/viz-trio/`:

- `unsupported-by-current-repo-docs` (no explicit command signatures found in current local docs)

Governance handling:

- Do not infer semantics beyond observed usage in core script.
- Treat behavior as potentially environment-specific.
- Fail closed if command availability/behavior is uncertain.

## Socket Payload Surface (Observed)

Observed emitter:

- `SmartStat_v4.0.0_beta.vbs:5079` (`SmartStat_RefreshSocketData`)

Observed payload pattern:

```text
sock:send_socket_data on_air_get message_number=<page_name> query=<on_air_tabs> message_context=<page_desc>\r\n
```

Observed field notes:

- `message_number` is sourced from `page:getpagename`.
- `query` is a serialized list of tab/custom-property pairs.
- `message_context` is page description with embedded ambiguity summary, linebreak-normalized to `\n`.

Risk notes:

- Payload shape is runtime-observed, not a versioned external schema in this repo.
- Consumers must treat this as best-effort telemetry unless a versioned contract is added.
- This surface must not override fail-closed runtime safety rules.

## Related Grounded Commands

These commands are explicitly documented in current repo Viz Trio docs and remain grounded:

- `page:get_property`
- `page:set_property`
- `page:get_tabfield_names`
- `tabfield:get_custom_property`
- `page:getpagename`
- `page:getpagetemplate`

References:

- `docs/viz-trio/command_reference.md`
- `docs/viz-trio/commands_full_index.md`
