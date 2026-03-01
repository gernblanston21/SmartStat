---
source: "Viz Trio User Guide"
source_version: "3.2"
doc_page: "Page_Editor.html"
url: "https://docs.vizrt.com/viz-trio-guide/3.2/Page_Editor.html"
doc_type: "manual_paraphrase"
smartstat_relevance: "high"
signals:
  tabfield_logic: true
  command_layer: true
  ui_only: false
---

# Page Editor (Viz Trio 3.2)

The **Page Editor** is where pages/templates are edited. It can host multiple specialized editors depending on the template (text, database, scroll, tables, etc.).

## Retrieval signals
SIGNATURES:
- page:read <pageNr>
- page:save
- page:saveas <pageNr>
- page:get_property <TABFIELD>
- page:set_property <TABFIELD> <VALUE>
- tabfield:edit_property <propertyName>

Tags: `[SMARTSTAT_RELEVANT] [TABFIELD_LOGIC] [EDITOR_WORKFLOW]`

---

## Where it lives + how operators navigate

- Located upper-right of the UI; typically visible by default.
- Operators can navigate editable elements via:
  - **TAB** to step between elements
  - clicking elements in the preview window to focus/select

**SmartStat impact:** scripts should not rely on UI focus state. Always address tabfields explicitly via commands.

---

## Editing a page (typical operator flow)

1. **Read** a page (double-click or Read)
2. Modify properties through the relevant editor
3. **Take** to air (optional; saving is not required before taking a page)
4. **Save** / **Save As** to persist changes

Important nuance:
- A **template** cannot be taken on-air until it has been saved as a **page** at least once.

---

## Controls (core buttons + modes)

The Page Editor includes controls for:
- playout
- script
- save / save as
- transition effect selection/application
- refresh
- variant selection

### Callup Code field
- Stores the callup code assigned to a page.
- Can be set on first save, or when saving as a new page.

### Field Linking (templates only)
When editing **templates** (not pages), a **Field Linking** panel can appear:
- define an external feed URI
- map tabfield properties to external feed properties
- optionally enable “structured content” mapping (VDF / ATOM-feed based)
- optionally auto-refresh values on read

**SmartStat impact:** if a page uses linked data, your tooling should consider whether values may change after read/refresh.

---

## Playout buttons (behavior notes)

Common meanings:
- **Take**: takes the currently displayed page on-air (if in On Air mode)
- **Continue**: continues animation if there are stop points
- **Take Out**: takes out the page in preview (or takes out elements in the same layer if the preview page is not on-air)
- **Take + Read Next**: takes current page and reads the next item in the active view
- **Cue**: cues a specified page (or selected page if no parameter provided)

**SmartStat impact:** Take Out behavior can be layer-dependent; avoid automation that issues Take Out unless you’re confident about layers.

---

## Save behavior (template vs page)

### When working with a template
- **Save template** persists template modifications (defaults, feed linking, layout, etc.)
- **Save as** creates a new page based on the template
- Default callup code typically becomes “highest existing callup code + 1”

### When working with a page
- **Save page** persists changes without changing callup code
- **Save as** creates a new page (with default “highest + 1” behavior)

**SmartStat impact:** if your workflow depends on stable page IDs, do not “Save As” casually.

---

## Editors you may encounter

Depending on template design:
- Text editor (kerning, style, etc.)
- Database linking editors
- Image property editor
- Transformation properties
- Tables (multi-column value editing)
- Clock editor
- Maps editor

**SmartStat impact:** table/scroll editors can pack multiple values into one control; for robust automation, prefer direct tabfield addressing via `page:get_property` / `page:set_property` when possible.

---

## SmartStat-specific takeaways

- The Page Editor is where values are authored, but automation should treat it as **a view**, not a dependency.
- Your safest model is:
  1) snapshot tabfields
  2) compute outputs
  3) apply writes
  4) optionally store custom properties for traceability
