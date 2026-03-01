# Viz Trio 3.2 – Command Reference (SmartStat-oriented)

Primary source (Commands page):
- https://docs.vizrt.com/viz-trio-guide/3.2/Commands.html

---

## Page property commands

> Naming varies by environment and integration style. Below are the command patterns SmartStat-style tools typically rely on.

### Read a tabfield value
```vb
val = TrioCmd("page:get_property H0010")
```

### Write a tabfield value
```vb
TrioCmd("page:set_property H0010 " & value)
```

**Operational guidance**
- Always sanitize/escape values that may contain spaces/quotes if your environment requires it.
- Prefer “read-all, compute, write-all” to reduce mid-run operator changes.

---

## Custom property commands (tabfield metadata)

Viz Trio exposes commands to set/get custom properties on tabfields (often via the “tabfield:” namespace in the commands list).

### Set a custom property
Conceptual pattern:
```text
tabfield:set_custom_property <tabfield> <value>
```

### Get a custom property
Conceptual pattern:
```text
tabfield:get_custom_property <tabfield>
```

**SmartStat usage ideas**
- store resolved mapping key
- store resolved entity/qualifier snapshot
- store last-generated syntax (for diffing / “no-op” detection)

---

## UI messaging

### Non-blocking error message
```vb
TrioCmd("gui:error_message " & msg)
```

Use this sparingly:
- only on actionable failures
- keep message short; details go to logs

---

## Notes for building a command glossary

When you expand this reference further, categorize by:
- Page operations
- Tabfield operations
- UI operations
- Engine/control operations
- Rundown/show operations

This makes retrieval much cleaner than one long list.
