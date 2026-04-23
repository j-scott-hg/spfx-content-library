# Content Library (SPFx)

SharePoint Framework web part that connects to any SharePoint **list** or **document library** and presents items with modern layouts: card grid, preview thumbnails, tiles, table, and dashboard views. Includes search, optional category filters with colour coding, per-item icon and thumbnail overrides, and a **details** modal for list items.

## Requirements

- Node.js **18.17.1** (SPFx 1.20.x compatible range; see `package.json` `engines`)
- SharePoint Online

## Build and package

```bash
npm install
gulp bundle --ship
gulp package-solution --ship
```

The `.sppkg` is written under `sharepoint/solution/` (see `config/package-solution.json`).

## Details window thumbnail

For **list** items (not document libraries), opening an item shows a details modal. Thumbnail behaviour:

- **Show thumbnail in details window when available** (property pane → Visible Fields): when **On**, the modal uses the same image source as Preview cards (`customThumbnailUrl` from per-item overrides, or SharePoint preview when `fileRef` exists). When **Off**, the header shows the **item icon** (override icon/colour if set, otherwise file-type icon).
- **Thumbnail layout**: **Left of title** (compact 16∶9 thumbnail beside the title) or **Above title** (image centered above a centered title with spacing).
- If the image fails to load, the header falls back to the icon layout.
- Icon in the header respects **icon overrides** from the icon editor when the thumbnail is not shown or after image error.

## Repository

If GitHub reports a moved repository, update the remote:

```bash
git remote set-url origin https://github.com/j-scott-hg/spfx-content-library.git
```

## Licence

See project licence file if present; otherwise follow your organisation’s policy for this codebase.
