# Content Library — SPFx Web Part

A modern, fully-featured SharePoint Framework (SPFx) web part that displays content from any SharePoint list or document library with multiple display styles, optional keyword search, category filtering with colour coding, per-item icon overrides, resizable columns, and a detailed item panel — all without leaving the page.

---

## Previews

<table>
<tr>
<td align="center"><strong>Tile view with category colours</strong></td>
<td align="center"><strong>Card grid with filters</strong></td>
</tr>
<tr>
<td><img src="docs/preview-tiles.png" width="480" alt="Tile view" /></td>
<td><img src="docs/preview-cards.png" width="480" alt="Card grid view" /></td>
</tr>
</table>

<p align="center">
  <strong>Table / List view with horizontal scroll and resizable columns</strong><br/>
  <img src="docs/preview-list.png" width="960" alt="Table view" />
</p>

---

## Features at a Glance

| Feature | Description |
|---|---|
| **Multiple display styles** | Table/list, compact card grid, large tile grid, dashboard |
| **Document libraries & lists** | Works with both; clicking a document opens the file, clicking a list item opens a details panel |
| **Keyword search** | Live, debounced search across configurable fields; multiple bar styles |
| **Category filters** | Horizontal pills, vertical rail, card selector, or compact buttons |
| **Category colour coding** | Assign colours to each category; cards/tiles tint to match |
| **Resizable columns** | Drag the right edge of any column header in table view to resize |
| **Show/hide columns** | "Columns" button in table view reveals a panel to toggle visibility per column |
| **Per-item icon overrides** | Edit-mode hover button lets authors swap the icon for any individual item |
| **Font & icon size controls** | Sliders to scale item name text and icon independently |
| **Item detail panel** | Clicking a list item opens a clean side panel showing all view columns |
| **Property pane organisation** | Two-page property pane: Data & Display on page 1, Columns & Advanced on page 2 |

---

## Supported SPFx Version

![SPFx version](https://img.shields.io/badge/SPFx-1.20.0-green.svg)
![Node.js](https://img.shields.io/badge/Node.js-18.x-brightgreen.svg)
![SharePoint Online](https://img.shields.io/badge/SharePoint-Online-blue.svg)

---

## Prerequisites

- A Microsoft 365 tenant with SharePoint Online
- Node.js 18.x
- SharePoint Framework development toolchain (`@microsoft/generator-sharepoint`)
- An existing SharePoint list or document library to connect to

---

## Installation (deploy to your tenant)

1. Download `content-library.sppkg` from the `sharepoint/solution/` folder (or build it yourself — see below).
2. Upload it to your **tenant App Catalog** (`https://<tenant>.sharepoint.com/sites/appcatalog/AppCatalog`).
3. Click **Deploy** when prompted. Choose **Make this solution available to all sites** if you want tenant-wide availability.
4. On any modern SharePoint page, edit the page, click **+**, search for **Content Library**, and add it.

---

## Building from Source

```bash
# 1. Clone the repository
git clone https://github.com/<your-org>/content-library.git
cd content-library

# 2. Install dependencies
npm install

# 3. Bundle for production
gulp bundle --ship

# 4. Package
gulp package-solution --ship
# → Output: sharepoint/solution/content-library.sppkg

# 5. (Optional) Local workbench
gulp serve
```

---

## Configuration — Property Pane

### Page 1 — Data Source & Display Style

#### 📋 Data Source

| Setting | Description |
|---|---|
| **Web part title** | Optional heading shown above the web part |
| **Show title** | Toggle the title on/off |
| **Site URL** | Leave blank to use the current site, or enter the full URL of another site in your tenant |
| **List or document library** | Choose the list/library to display (loaded dynamically after Site URL is set) |
| **View to display** | Choose which saved SharePoint view drives the visible columns, sort order, and data |
| **Maximum items to load** | Number slider (1–500); limits how many items are fetched |

> **Tip:** The selected view controls exactly which columns are shown in all display modes.  
> Views can be created and managed directly in the SharePoint list or library.

#### 🎨 Display Style

| Setting | Description |
|---|---|
| **Item display style** | Table/list · Card grid · Tile grid · Dashboard |
| **Spacing density** | Compact · Normal · Comfortable |
| **Card corner radius** | Rounds card/tile corners (0–20 px) |
| **Shadow intensity** | None · Subtle · Medium · Strong |
| **Show file type icon** | Show/hide the icon next to each item name |
| **Show description** | Show/hide any Description or Note field on cards/tiles |
| **Show column headers** | Show/hide header row in table mode |
| **Columns (card/tile modes)** | Number of columns in the grid (1–6) |
| **Item font size** | Scales the item name/title text only (does not affect metadata lines) |
| **Icon size** | Scales the item icon |

#### 🔍 Search Settings

| Setting | Description |
|---|---|
| **Enable search** | Toggle the search bar on/off |
| **Search placeholder** | Custom placeholder text for the search input |
| **Search bar style** | Minimal · Elevated card · Toolbar-integrated |
| **Search bar position** | Top full-width · Top right · Integrated with toolbar |
| **Debounce (ms)** | Delay after the user stops typing before results update (default 300 ms) |
| **Fields to search** | Which internal field names are searched (defaults to Title/FileLeafRef) |

#### 🏷 Filter / Category Settings

| Setting | Description |
|---|---|
| **Enable filters** | Toggle the filter bar on/off |
| **Filter field** | The internal name of the column whose values become filter tabs (Choice, Text, Yes/No, etc.) |
| **Filter style** | Horizontal pills · Vertical side rail · Card/tile selector · Compact buttons |
| **Filter position** | Top · Left · Right |
| **Show "All" option** | Prepend an "All" tab that shows everything |
| **"All" label** | Custom text for the All tab (default: `All`) |
| **Show counts** | Display item count next to each category |
| **Sort categories** | Alphabetical or by item count |
| **Max visible categories** | Truncate after N categories |

> **Recommended filter fields:** Choice columns (e.g. Status, Category, Department) give the most reliable results.  
> Single-line text columns work but values must be consistent.

#### 🌈 Category Colour Coding

When **Enable category colours** is on, each category can have its own colour. Cards and tiles tint to match their item's category colour, and text/icon contrast adjusts automatically.

| Control | Description |
|---|---|
| **Colour swatch** | Click to open a colour picker (theme swatches + full HSV picker) |
| **Text toggle (A / ☀ / 🌙)** | `A` = auto-contrast · `☀` = force white text · `🌙` = force dark text |

The **All** category always renders in light grey and is not colour-customisable.

---

### Page 2 — Visible Fields & Advanced

#### 📄 Card / Tile Detail Lines

These two dropdowns control which metadata appears below the item name on cards and tiles.

| Setting | Description |
|---|---|
| **Detail line 1** | First metadata line (default: Modified date) |
| **Detail line 2** | Second metadata line (default: Modified by) |

Built-in options (Modified date, Created date, Modified by, Created by) are always available.  
Custom columns appear in the dropdowns once a list and view are selected on page 1.

> **Note:** A list and view must be selected on page 1 before custom columns appear in these dropdowns.

#### ⚙️ Advanced

| Setting | Description |
|---|---|
| **Open documents in** | Same tab or new tab (document libraries only) |
| **Empty state message** | Custom message shown when no items match the current filters/search |
| **Custom CSS class** | Append an extra CSS class to the web part root for custom styling |

---

## Runtime Controls (on the web part itself)

These controls are available directly on the rendered web part without opening the property pane.

### Table / List view — Columns button

A **Columns** button appears in the top-right corner of the table view. Clicking it opens a side panel listing all columns with checkboxes. Unchecking a column hides it immediately. The selection resets when you switch list or view.

### Table / List view — Column resizing

Hover over the right edge of any column header to reveal a blue resize handle. Drag left or right to adjust the column width. All columns have a minimum width of 60 px. Widths reset when you switch list or view.

### Edit mode — Per-item icon override

In SharePoint page **Edit** mode, hovering over any item shows a small paint-brush button. Clicking it opens a panel where you can:

- Pick any Fluent UI icon from the searchable icon grid
- Choose a custom colour for that icon

The override is saved to the web part properties and persists across page loads. Overrides can be cleared by selecting the default icon.

---

## Display Modes

### Table / List view

- Closest to the native SharePoint list experience
- Column headers with optional user-sortable click targets
- Horizontal scroll with a styled scrollbar when columns overflow
- Resizable columns (drag right edge of header)
- Show/hide individual columns at runtime

### Card Grid view

- Responsive grid of modern rounded cards
- Shows file/item icon, name, and two configurable metadata lines
- Category colour tinting with accent bar at top of each card
- Hover states and keyboard navigation

### Tile Grid view

- Square/rectangular tiles with prominent icon
- Ideal for icon-heavy document portals or quick-access grids
- Supports category colour tinting

### Dashboard view

- Mixed layout with featured sections
- Starred/favourited items functionality
- Best for document-centre or portal pages

---

## Item Interaction

| Source type | Click behaviour |
|---|---|
| **Document library** | Opens the file in the same or new tab (configured in Advanced settings) |
| **List** | Opens a modern side panel showing all columns visible in the selected view |

The list item detail panel includes:
- A hero header with the item icon and title
- A row for each view column with a value
- A footer showing the created date and author
- Light-dismiss (click outside to close) and a close button

---

## Technical Details

| Technology | Details |
|---|---|
| **Framework** | SharePoint Framework (SPFx) 1.20 |
| **UI library** | Fluent UI v8 (`@fluentui/react`) |
| **Data access** | PnPjs v3 (`@pnp/sp`) |
| **Language** | TypeScript 4.7 |
| **Styling** | CSS Modules (SCSS) |
| **React** | 17.x with hooks |
| **Toolchain** | gulp 4, webpack 5 |

### Architecture

```
src/webparts/contentLibrary/
├── ContentLibraryWebPart.ts        # Web part class, property pane
├── components/
│   ├── ContentLibrary.tsx          # Main orchestrator (state, data, filtering)
│   ├── DocumentTableView/          # Table/list display mode
│   ├── DocumentCardGrid/           # Card grid display mode
│   ├── DocumentTileGrid/           # Tile grid display mode
│   ├── DashboardView/              # Dashboard display mode
│   ├── SearchBar/                  # Search input component
│   ├── FilterBar/                  # Category filter bar
│   ├── SortBar/                    # Standalone sort controls
│   ├── ItemIconEditor/             # Per-item icon override panel
│   └── CategoryColorPicker/        # Custom colour picker for categories
├── services/
│   ├── SharePointDataService.ts    # PnPjs data fetching
│   └── ViewMapper.ts               # Maps view field names to column definitions
├── helpers/
│   ├── categoryExtraction.ts       # Extracts unique category values from items
│   ├── colorUtils.ts               # Colour math (HSV, contrast, tinting)
│   ├── fieldFormatting.ts          # Formats field values for display
│   ├── fileIconMapping.ts          # Maps file extensions to Fluent UI icons
│   └── searchUtils.ts              # Client-side search and sort
├── models/
│   ├── IListItem.ts                # SharePoint item interfaces
│   └── IWebPartConfig.ts           # All property pane settings + defaults
└── styles/
    └── ContentLibrary.module.scss  # Scoped SCSS for all components
```

---

## Version History

| Version | Date | Notes |
|---|---|---|
| 1.0.29 | April 2026 | Column resizing, column show/hide picker |
| 1.0.28 | April 2026 | Column show/hide panel, horizontal table scroll |
| 1.0.26 | April 2026 | Document library items loading fix, correct file URL construction |
| 1.0.24 | April 2026 | View field fetching fix (expand ViewFields), close button on detail panel |
| 1.0.22 | April 2026 | Modern custom detail panel replacing iframe |
| 1.0.21 | April 2026 | Card meta fields re-fetch on dropdown change |
| 1.0.19 | March 2026 | Detail line note in property pane, card meta custom columns |
| 1.0.10 | March 2026 | Missing icons fix (`initializeIcons`) |
| 1.0.7 | March 2026 | Category colour coding, font/icon size sliders |
| 1.0.5 | March 2026 | App catalog icon, toolbox section, web part rename |
| 1.0.3 | March 2026 | Per-item icon override editor |
| 1.0.1 | March 2026 | Search position, debounce, filter fixes |
| 1.0.0 | March 2026 | Initial release |

---

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
