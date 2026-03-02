# Viz Trio 3.2 – Tabfields (SmartStat-oriented)

---

## What is a tabfield?

A **tabfield** is an operator-editable field exposed by a template/page.
Typical types include:
- Text fields
- Image references / paths
- Toggles / switches
- Numeric controls
- Database/scroll editors (template-dependent)

In SmartStat, tabfields are your **inputs** (filters/categories/controls) and **outputs** (generated syntax, mapped fields).

---

## Naming patterns

Many shows use a consistent naming pattern:
- `Letter + 4 digits` (example: `H0010`)

### Practical implication
SmartStat should **not** assume numeric adjacency means semantic adjacency.
Example pattern you’ve seen in the wild:
- category: `H0100`
- output columns: `H0110`, `H0120`, `H0130` (not `H0101`, `H0102`)

So: map by **pattern heuristics**, not “next number” math.

---

## Prefix conventions (broadcast practice)

These are *common* conventions (not guaranteed by Viz Trio itself):

| Prefix | Typical meaning | SmartStat usage |
|---|---|---|
| `A####` | toggles / switches | usually ignore/deprioritize for mappings |
| `B####` | logos, qualifiers, filters, toggles | high value for qualifier/filter inference |
| `C####` | row/column and layout controls | infer row_limit / column count / spacing |
| `E####` | sponsors | usually ignore |
| `H####`–`Y####` | player/team/stat data | primary category + output_map targets |

**Design rule for SmartStat tooling:** treat prefix inference as a *soft signal*, not a hard rule.

---

## Custom properties on tabfields

Tabfields can carry metadata via **custom properties**.
SmartStat uses these to store:
- resolved mapping keys
- resolved syntax strings
- diagnostic hints (optional)
- stable operator configuration without changing visible values

When building tools:
- Keep custom property values short and deterministic
- Prefer “namespaced” property keys if you create new ones (avoid collisions)

---

## Common tabfield workflows

### Read → decide → write
1. Read filter/category controls
2. Resolve mapping keys + templates
3. Write generated syntax back into output tabfields
4. Optionally store intermediate data in custom props

### Use row/column controls safely
C-fields often drive layout:
- number of rows
- row limit
- column spacing
- visibility switches

SmartStat should avoid fighting layout logic:
- treat C-fields as “authoritative UI controls”
- write only what you own (stat outputs), not layout controls, unless explicitly required

---

## Failure modes to guard against

- Empty tabfields where logic expects values
- Non-standard naming (no letter+digits)
- Template-specific editors (scroll/db) where raw strings behave differently
- Operators changing values mid-run (race conditions)

Mitigations:
- snapshot reads first, then compute, then apply writes
- fail closed if the template config doesn’t match reality
