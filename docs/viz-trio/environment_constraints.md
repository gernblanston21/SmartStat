# Viz Trio 3.2 – Environment Constraints (Live Production)

This file is the “rules of the road” for any tooling that runs during a live show.

---

## What “live-safe” means

### Must-haves
- **Non-blocking UX**: never pause the operator
- **Fail closed**: when uncertain, do nothing (don’t spray bad data)
- **Fast**: predictable execution time
- **Clear recovery path**: log + tell operator what to do next

### Avoid
- Modal dialogs (MsgBox, input prompts that block)
- Heavy file IO loops
- Network calls on the hot path
- Large dynamic allocations (string concatenation storms)

---

## Error handling expectations

- Log errors to a known file path
- Surface a short “error occurred” message in UI
- Include a run-id or timestamp so logs correlate to operator actions

---

## “Do not break the operator” rules

1. Never change tabfield naming conventions.
2. Never reorder template config keys in INI blocks.
3. Never write outputs if the template/config is not confidently resolved.
4. Never overwrite user-entered values outside SmartStat-owned outputs.
5. Prefer idempotent writes (only set values if changed).

---

## SmartStat-specific constraints

- Ambiguity gating is required for safety.
- Output_map must be deterministic and validated before writing.
- Normalize mapping keys consistently (spaces → underscores, etc.).
