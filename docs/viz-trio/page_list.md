---
source: "Viz Trio User Guide"
source_version: "3.2"
doc_page: "Page_List.html"
url: "https://docs.vizrt.com/viz-trio-guide/3.2/Page_List.html"
doc_type: "manual_paraphrase"
smartstat_relevance: "high"
signals:
  tabfield_logic: true
  command_layer: true
  ui_only: false
---

# Page List (Viz Trio 3.2)

The **Page List** is the operator’s main “inventory” of **pages** (instances created from templates). Pages can be played out directly from the Page List, or used inside playlists.

## Retrieval signals
SIGNATURES:
- page:read <pageNr>
- page:take <pageNr> [channel]
- page:takeout <pageNr> [channel]
- page:get_property <TABFIELD>
- page:set_property <TABFIELD> <VALUE>

Tags: `[SMARTSTAT_RELEVANT] [PAGE_WORKFLOW] [TABFIELD_LOGIC]`

---

## Concept: template vs page

- **Template** = blueprint (usually imported from Viz Artist scene)
- **Page** = a saved, unique instance created from a template, with its own **page ID / callup code**
- You can create unlimited pages from one template; each page is distinct.

**SmartStat impact:** your scripts usually operate on the “currently read” page, so your logic must be robust when operators read different pages quickly.

---

## Page content filling (core workflow)

A typical create → fill workflow looks like:

1. **Create / open a show** (Show Directories via File > Open Show or Change Show)
2. **Import scenes** into the show (so templates exist in Template List)
3. **Convert a template into a page**:
   - Open template in Page Editor
   - Use **Save As** (first save) to create a page (callup code assigned)
4. **Fill page content**:
   - Drag/drop media (images/video) into editors (for example from Viz One, if integrated)
5. **Use it**:
   - Take on-air from Page List **or**
   - Add to a playlist

Note: “composite elements” may not be displayable from the Page List (template/design dependent).

---

## Context menu (right-click)

Right-click a page to access actions and settings. Typical actions include:
- **Read** the page (load into preview + Page Editor)
- **Direct Take / Direct Cut / Direct Continue / Direct Take Out**
- **Delete**
- Variant selection (via icon / dropdown) if variants are configured

**SmartStat impact:** “Direct” operations can bypass the read/edit step; tooling should not assume a page is always “read” before being taken.

---

## Columns (what the operator can expose)

Operators can right-click column headers to toggle which columns appear.

Examples of column meanings (high-level):
- **Available**: availability percentage of an element
- **Channel**: playout channel selection
- **Modified**: timestamp for sorting (sorting must be enabled to use effectively)
- **Page**: page/callup code (often editable inline)
- **Loop**: loop behavior for scenes or full-screen video clips (see notes below)

### Loop notes (important nuance)
Looping behavior depends on:
- element type (video vs graphics)
- whether certain config settings are enabled/disabled
- relationship to timeline editor looping

**SmartStat impact:** avoid any automation that toggles loop unless you fully control the show’s configuration policy.

---

## Common operator procedures

### Reading a page
- Select the page in the list → double-click, or context menu **Read**, or keyboard **Read** key.
- Reading loads it into preview and opens it in the Page Editor for editing.

### Taking pages on-air
Typical safe sequence:
1. Select page
2. Read page
3. Switch to **On Air mode**
4. Take (and Continue if stop points exist)
5. Take Out

### Deleting pages
- Context menu → delete, or Delete key.
- Multi-select supported (Shift/Ctrl).

### Export selected pages
- File → Export Selected Pages Archive…
- Export is written as an XML-based archive (used for portability/backup).

### Variant selection
- Right-click the Scene Icon column → choose a variant from the dropdown (if variants are configured).

### Adding transition effects (via Page List)
- Show the Effect columns → choose effect via ellipsis → apply to items.

---

## SmartStat-specific takeaways

- Page List is your **entry point** to what is “currently live-editable”.
- Automation should not assume:
  - sorting is enabled
  - columns are visible
  - operators always read before take
- Always snapshot the “current page state” before generating outputs.
