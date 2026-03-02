---
source: "Viz Trio User Guide"
source_version: "3.2"
doc_page: "Show_Control.html"
url: "https://docs.vizrt.com/viz-trio-guide/3.2/Show_Control.html"
doc_type: "manual_paraphrase"
smartstat_relevance: "medium"
signals:
  tabfield_logic: false
  command_layer: true
  ui_only: false
---

# Show Control (Viz Trio 3.2)

**Show Control** is how operators manage a show before/during/after taking it on-air.
It’s also where you access:
- previous shows
- Viz Pilot / newsroom playlists (if integrated)
- show playlists
- show-level properties (like database connections and transition paths)

Tags: `[SHOW_WORKFLOW] [SMARTSTAT_RELEVANT]`

---

## Top controls (what the buttons do)

Show Control includes buttons that open key workflows:
- **Change Show**: open Show Directories (choose/open/import/create shows)
- **Add Page View**: create filtered Page List views (callup code ranges)
- **Create Playlist**: create playlists inside the show
- **Show Properties**: show-wide configuration and connections
- **Cleanup Renderers**: clears renderer data (cleanup channels)
- **Initialize**: initializes channels for the show
- **Show Concept**: displays show concept (e.g., Sport/News)
- **Callup Code** + **Read Page**: read the page matching the next callup code

---

## Show Directories

“Show logic” is the recommended organizational model for pages.

Key points:
- Shows and playlists tabs are visible by default.
- There is also a **Viz Directories** view for compatibility with older workflows.
- Viz Directories can be disabled by default and enabled via Configuration → User Restrictions.

### Importing an exported show
High-level flow:
1. Change Show
2. Import Show
3. Select a `.trioshow` file
4. Choose which archived elements to import
5. Optionally merge into an existing show

---

## Playlists (newsroom / Viz Pilot integration)

If integrated, the Playlists tab can show newsroom-provided playlists:
- name
- start time
- duration
- host
- channel ID
- status
- active indicator

Operational nuance:
- newsroom playlists are often **read-only** by default,
  but configuration can allow editing depending on system policy.

---

## Add Page List View (filtered views)

Creates additional Page List “views” without changing the underlying list.

- You specify a **callup code range** (example: 1000–2000).
- A new filtered view is created showing only items in that range.
- No practical limit to number of views.
- Original views remain unchanged.

**SmartStat impact:** If your tooling assumes “the page list” is singular, be careful—operators may have multiple filtered views active.

---

## Show Properties

Show Properties can include:
- transition effects path (where effects are stored in Viz)
- show-specific folder associations (package dependent)
- show-specific colors (character colors, etc.)
- database connections (OLE DB / Excel)

### Setting the transition effects path
- Open Show Properties → select Transition Path → browse to effects folder → OK.

### Creating a new database connection
- Open Database Config via Show Properties
- Add database → choose connection type (OLE DB or Excel)
- Configure connection string / provider, test connection

**SmartStat impact:** Database connections and effects paths influence what editors and lookups are available inside Page Editor. If you are troubleshooting “why does this template behave differently on this machine,” Show Properties is a prime suspect.

---

## Cleanup vs Initialize (conceptual)

- **Cleanup Renderers**: clears data on renderers/channels (a “reset”)
- **Initialize**: prepares the show on renderers (a “setup/prime”)

**SmartStat impact:** tools that assume persistent state on engines should tolerate cleanup/initialize being invoked at any time.
