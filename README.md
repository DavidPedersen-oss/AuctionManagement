# CSULB Auction Console

A single-page console for tracking surplus items flagged for auction, pulled
straight from your **Master Survey** workbook. No server, no Python, no install.
Built in the CSULB Property Management identity — black, gold, and red.

## Run it

**GitHub Pages (shareable link)**
1. Put `index.html`, `app.js`, and `xlsx.full.min.js` in a repo.
2. Settings → Pages → deploy from `main`, root folder.
3. Open the published URL.

**Local (no internet needed for data)**
- Open `index.html` in any browser. Keep the three files together.
- Fonts (Oswald/Inter) load from Google Fonts when online; offline they fall
  back to system fonts gracefully.

## Working with it

**Import** — drop your `Master_-_Survey.xlsx`. It reads the survey sheet, finds
the header row wherever it sits, and pulls every row where **Disposal Action =
AUCTION**. Multiple line items under one survey collapse into a single entry
(the individual items are kept in that entry's notes). Re-importing keeps your
edits and only adds new surveys.

**Excel-style column filters** — every column header has its own control:
- Survey #, Item, Tag → live text filter (type to narrow).
- Department, Cond., Status, Platforms → click the **ALL ▼** button for an
  autofilter checklist with value counts, search, and Select all / Clear.
- Click the header label itself to sort.
- Active filters show as removable pills under the search bar.

**Change status without opening an item** — the Status column is an inline
dropdown on every row. Pick a value and it saves instantly, recoloring to match.

**Batch editing** — tick rows, then use the bar to set status, mark a platform,
log a release (+1), or delete. Select-all respects the current filters.

**Detailed editing** — click any row (or Edit) for the full side panel: every
field plus per-platform releases, list price, sold price, and listing URL.

**Export / Restore** — JSON backup (full fidelity) or CSV (for Excel). Restore
reloads a JSON backup.

## Statuses

Prep · Live · Sold · Unsold · Relisted · Paid · Picked up · Closed · Review

## Files

    index.html           the console (UI, styles, embedded logos)
    app.js               all logic (parse, filter, edit, persist)
    xlsx.full.min.js     SheetJS, vendored so import works offline / on Pages

## Adding live scraping later

Keep this as the system of record. A scraper can later match listings back to
these entries by survey #, tag #, or asset number and update each item's release
count and sold price — those fields already exist in the data model. Because the
app is static, scraping runs as a separate step that writes an updated JSON you
re-import.

## Backup

Your working copy lives in this browser. Export JSON regularly — that file is the
real backup and restores everything exactly.
