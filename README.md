# CSULB Auction Console — v6

A single-page console for tracking surplus items flagged for auction across
**Public Surplus** and **GovDeals**, pulled from your Master Survey workbook.
No server, no Python, no install. Built in the CSULB Property Management
identity — black, gold, and red.

## Run it

**Local (no internet needed for your data)**
- Open `index.html` in any browser. Keep all four files together:
  `index.html`, `app.js`, `features.js`, `xlsx.full.min.js`.

**GitHub Pages (shareable link)**
1. Put the four files in a repo.
2. Settings → Pages → deploy from `main`, root folder.
3. Open the published URL.

## What's new in v6

**Date Assigned column** — shown on the dashboard, sortable, editable in the
detail drawer, and read from the Master Survey on import (recognizes
"Date Assigned", "Survey Date", "Assigned Date"; handles Excel date cells).

**"Current Bid" replaces "Est. Value" on the dashboard** — the column now
shows the live price an item is at right now (the highest current bid across
its listed platforms). The original at-purchase value is still kept — it just
lives in the drawer now as "Est. value at purchase," since it isn't useful at
a glance.

**Paste a listing to import it** — `Pull Listings` → paste the description
HTML from a Public Surplus *or* GovDeals listing. It reads the title,
condition, Survey#, Tag#, and bullet details, then either updates the matching
item (matched by Survey# → Tag# → title) or adds it as new.

**Built-in HTML converter** — `Convert HTML` turns a Public Surplus listing
block into GovDeals format or vice-versa, rebuilding the markup in each
platform's own conventions (PS uses underlined/bold headings and sibling
`<ul>` nesting; GD uses `<b><u>` headings, `Condition: FAIR`, and nested
`<ul>` inside `<li>`). You can also create a tracked item straight from a
pasted block.

**Weekly recap generator** — `Recap` produces a recap in your email format
(Sold / Collected / Prepping → Relist, New, Released), grouped by department,
each item as "Description – Asking Price $X." Copy it or download as .txt.

**Pull live listings (best-effort)** — `Pull Listings` → `Fetch from web`
tries to read the two public search pages through a CORS proxy and update
matching items with live bids.

### A note on web scraping from a static page

A page opened from disk (or GitHub Pages) **cannot reliably fetch** GovDeals
and Public Surplus directly: those sites block automated requests, and the
browser blocks cross-site reads. Out of the box, `Fetch from web` routes
through a public CORS proxy as a convenience, but it is best-effort — it can
be rate-limited or stop working without warning.

**The reliable path is the small `auction-scraper.js` Cloudflare Worker
included in this repo.** It fetches the public search pages server-side from
Cloudflare's edge (no CORS, no per-user-IP rate limits) and returns the raw
HTML to the console, where the existing parsers do the extraction. Deploy
once and paste the Worker URL into the Pull Listings modal — no code change
required to switch over.

#### Deploying the scraper Worker (one-time, ~2 min)

```
npm install -g wrangler
wrangler login
cd <this repo>
wrangler deploy
```

The deploy prints a URL like
`https://csulb-auction-scraper.<your-subdomain>.workers.dev`.

Then in the console, open **Pull Listings** → paste that URL into the
**Scraper Worker URL** field → it saves automatically to your browser.
`Fetch from web` now uses the Worker (the result line says `via scraper
Worker`). Clearing the field falls back to the legacy CORS proxy.

No KV, no secrets — the Worker reads only two public web pages and has a
60-second edge cache so it stays friendly to the origin.

**The always-works fallback** is paste-the-HTML: open a listing, copy its
description HTML, paste it below in the Pull Listings modal. This works
offline and never depends on a third party.

## Working with it

**Import** — drop your `Master_-_Survey.xlsx`. Reads the survey sheet, finds
the header row wherever it sits, and pulls every row where **Disposal Action =
AUCTION**. Multiple line items under one survey collapse into one entry.
Re-importing keeps your edits and only adds new surveys.

**Athletics category** — surveys with a **letter-prefixed Survey #** (`J123`,
`Z456`, etc.) are recognized automatically as athletics / non-standard items,
separate from the numeric auction surveys. They're **hidden from the main view
by default** and left out of the stat-strip counts. The toolbar toggle cycles
`Athletics hidden → Athletics only → Athletics shown`, and any that appear carry
an amber **Athletics** badge. This is distinct from Archive — no manual
selection needed; it's driven purely by the survey number.

**Archive** — other items you don't actively track can be selected in the table
and bulk-Archived via the bulk action bar, or
archived per-item from the detail drawer's footer. Archived items are hidden
from the main view by default and excluded from the stat strip counts; click
the `Show archived (N)` toggle in the toolbar to bring them back. They stay in
your data and are exported with the JSON backup and CSV for full round-trip.

**Excel-style column filters, inline status, batch editing, detail drawer,
export/restore** — unchanged from v5; see the in-app footer hints.

## Statuses

Prep · Live · Sold · Unsold · Relisted · Paid · Picked up · Closed · Review

The recap maps statuses like this: Sold/Paid → **Sold**, Picked up →
**Collected**, Relisted/Unsold → **Relist**, Prep → **New**, Live →
**Released**.

## Files

    index.html           the console (UI, styles, embedded logos)
    app.js               core logic (parse, filter, edit, persist)
    features.js          v6 add-ons (listing parser, converter, scraper, recap)
    xlsx.full.min.js     SheetJS, vendored so import works offline / on Pages
    auction-scraper.js  optional Cloudflare Worker for reliable automatic pulls
    wrangler.toml        Wrangler config for the scraper Worker

## Backup

Your working copy lives in this browser. Export JSON regularly — that file is
the real backup and restores everything exactly.
