# Viz Trio 3.2 – Scripting Notes (SmartStat-oriented)

---

## Runtime reality

Viz Trio supports scripting + macro commands for customization.
In typical broadcast environments you will encounter:
- **VBScript** for operator macros/tools
- Command strings executed through the Trio command layer

SmartStat’s “engine” scripts should assume:
- single-threaded execution
- no async/await
- errors must be handled defensively
- operators need clear, non-blocking feedback

---

## Non-blocking operator UX

In live production, avoid:
- modal dialogs (anything that pauses the show operator)
- long-running operations without progress
- excessive logging spam

Prefer:
- log-to-file
- brief UI error messaging (non-blocking)
- early exits with clear reason

---

## Coding patterns that work well

### 1) Snapshot first
Read everything you need up front:
- template name
- relevant tabfields
- relevant custom props

Then compute, then apply writes in one pass.

### 2) Fail closed
If:
- template isn’t recognized
- config block missing/invalid
- ambiguity exists that cannot be resolved safely

Then do not write outputs. Log + notify.

### 3) Deterministic outputs
Generated syntax should be stable:
- normalized keys
- predictable ordering
- no random whitespace changes

This improves operator trust and reduces churn in the page.

---

## ClearScript JavaScript caveat (if used)
Some Trio JS environments are older/limited compared to modern browsers.
Avoid relying on newer array helpers unless you’ve validated them in your environment.

---

## SmartStat-specific scripting checklist

- [ ] No MsgBox / modal UI
- [ ] Writes only after validation + ambiguity checks
- [ ] All errors logged to file and surfaced via non-blocking message
- [ ] Never renames tabfields or changes conventions
- [ ] Preserves tabfield values you do not own
