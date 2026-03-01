---
source: "Viz Trio User Guide"
source_version: "3.2"
doc_page: "Commands.html"
url: "https://docs.vizrt.com/viz-trio-guide/3.2/Commands.html"
doc_type: "manual_paraphrase"
smartstat_relevance: "high"
signals:
  command_layer: true
  tabfield_logic: true
  ui_only: false
---

# Commands (Viz Trio 3.2) – Categorized Index (SmartStat-oriented)

This file is a **retrieval-first** command index:
- grouped by Viz Trio’s command categories
- highlights the commands most useful for SmartStat tooling
- avoids copying the full command manual verbatim

Tags: `[SMARTSTAT_RELEVANT] [COMMAND_LAYER]`

---

## Category list (as documented)

Viz Trio groups commands into (non-exhaustive list shown here):
- Channelcontrol, dblink, Gui, Macro, Main, Page, Playlist, Proxy, Rundown, Script, Scroll2,
  Settings, Show, Sock, Tabfield, Table, Text, Trio, Util, Viz, vtwtemplate

---

## 1) Page (highest SmartStat value)

Use **Page** commands to:
- read pages / navigate views
- get/set tabfield values
- enumerate tabfields / properties
- take/takeout pages

### Key commands for SmartStat
SIGNATURES:
- page:get_property(string name)
- page:set_property(string name, restString value)
- page:get_tabfield_names(optional pageName)
- page:get_tabfield_count
- page:get_property_keys(optional pageName)
- page:getpagename
- page:getpagetemplate
- page:read(pageNr)
- page:save
- page:saveas(pageNr)
- page:take(pageName, channelName)
- page:takeout(pageName, channelName)
- page:cue(pageName, channelName)

### Notes that matter in live ops
- Many take/cut/cue commands accept an optional **channel** argument.
- Some commands behave differently for scene-based vs standalone pages (layer queries can return empty for standalone scenes).

---

## 2) Tabfield (editing/selection helpers)

Use **Tabfield** commands to:
- change editor mode for a tabfield (clock, browse file, etc.)
- edit a specific property within a tabfield
- manipulate current tabfield state

### Representative commands (SmartStat-adjacent)
SIGNATURES:
- tabfield:active(active:string)
- tabfield:browse_file
- tabfield:clock
- tabfield:downcase_tabfield
- tabfield:edit_property(string propertyName)

**SmartStat guidance:** most SmartStat automation should prefer `page:get_property` / `page:set_property` so you aren’t dependent on “current tabfield” focus.

---

## 3) Script (VBScript execution + show scripts)

Use **Script** commands to:
- evaluate expressions
- run snippets
- call into show script functions

### Key commands
SIGNATURES:
- script:eval(restString scriptExpression)
- script:run_macro_script(restString scriptCode)
- script:run_script(string scriptFunction, restString argumentList)
- script:run_showscript(string scriptFunction, restString argumentList)
- script:get_showscript_name
- script:import_script(string scriptName, restString scriptFilename)

**Important nuance:** argument lists are separated by commas or whitespace; arguments containing commas/whitespace must be quoted.

---

## 4) Gui (operator UI control)

Use **Gui** commands for UI interactions such as:
- opening views/panels
- showing messages
- switching modes (varies by command set / policy)

### SmartStat-relevant patterns
- `gui:error_message <text>` is commonly used for *non-blocking* operator errors.

**SmartStat guidance:** keep UI messaging short; log details to file.

---

## 5) Show (show-level operations)

Show commands generally relate to:
- show lifecycle operations (open/import/export)
- show configuration behaviors
- housekeeping actions

SmartStat usually touches show indirectly, but show changes can impact:
- template availability
- database connections
- transition effects paths

---

## 6) Playlist / Rundown (automation around lists)

Playlist and Rundown commands cover:
- playlist filters, carousel start/stop, looping
- rundown story navigation and related controls

**SmartStat impact:** only relevant if SmartStat ever becomes “playlist-aware” (e.g., batch-filling graphics in a playlist).

---

## 7) dblink (database linking)

dblink is used to link template/page properties to a database connection and map columns.
This can matter if:
- your templates rely on database-driven fields
- you need to understand why a tabfield value “snaps back” or refreshes

---

## 8) Proxy / Sock (external control channels)

- **Proxy**: enabling Trio as a Viz Proxy, setting host/port
- **Sock**: socket configuration + sending UTF-8 data

**SmartStat note:** these are useful for external automation and control-plane designs, but many productions lock these down.

---

## SmartStat “top-10” commands (quick retrieval)

1. `page:get_property`
2. `page:set_property`
3. `page:get_tabfield_names`
4. `page:get_tabfield_count`
5. `page:get_property_keys`
6. `page:getpagename`
7. `page:getpagetemplate`
8. `page:read`
9. `page:save`
10. `gui:error_message`

---

## What to add next (if you want deeper coverage)

If you want this index to be “close to exhaustive” without copying the manual:
- add per-category tables listing command name → 1-line paraphrase
- prioritize Page/Tabfield/Script/Gui first
- keep each category in its own chunked file to improve retrieval:
  - `commands_page.md`, `commands_tabfield.md`, `commands_script.md`, `commands_gui.md`, ...
