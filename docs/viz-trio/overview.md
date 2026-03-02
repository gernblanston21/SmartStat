# Viz Trio 3.2 – Overview (SmartStat-oriented)

**Purpose of this folder:** Provide retrieval-ready, *developer-focused* notes about Viz Trio 3.2 for SmartStat/Codex work.

**Primary source (HTML manual index):**
- https://documentation.vizrt.com/viz-trio-guide/3.2/Viz_Trio_User_Guide.html
- Commands reference: https://docs.vizrt.com/viz-trio-guide/3.2/Commands.html

> Note: These docs are paraphrased and structured for engineering use (not a verbatim copy of the manual).

---

## Mental model

Viz Trio is a client UI that:
- Loads **templates** (imported from Viz Artist scenes)
- Creates **pages** (instances of templates)
- Lets operators edit page fields (tabfields) and trigger actions
- Communicates changes to Viz Engine where the scene logic responds

For SmartStat, Viz Trio is best treated as:
- A **tabfield container** (values + custom props)
- A **command bridge** (macro/command layer)
- A **live operator environment** (non-blocking, fast, fail-safe)

---

## Key objects

### Templates
A template is the “blueprint” for graphics. It defines:
- Which tabfields exist
- What editors appear (text, image, database, etc.)
- What scene logic runs in Viz Engine

### Pages
A page is a **single filled-out instance** of a template.
- Same template can produce unlimited unique pages
- Each page has its own identity and values
- In most SmartStat workflows: you operate on “current page”

---

## What matters most for SmartStat

### 1) Tabfields
- Names commonly look like `H0010`, `B0100`, `C0200` (letter + digits)
- Meaning is often **convention-based** per show/package
- Custom properties are frequently used as metadata stores

### 2) Command layer
SmartStat-style tools typically:
- Read tabfields
- Write tabfields
- Read/write tabfield custom properties
- Optionally surface operator errors (non-blocking)

### 3) Live constraints
- Never block the operator (no modal UI)
- Fail closed when uncertain
- Log and report errors safely

---

## Suggested reading order in this folder

1. `environment_constraints.md`
2. `command_reference.md`
3. `tabfields.md`
4. `scripting.md`
5. (Back here) `overview.md`
