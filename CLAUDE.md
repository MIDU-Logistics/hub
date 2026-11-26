# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

MIDU Logistics Hub: a single-page, installable web app (warehouse GRN/MDS/MRS, assets, utilities back charge, procurement/finance, HR attendance, maintenance) hosted as static files on GitHub Pages. There is **no build, no package manager, no test suite and no linter** — files are edited and uploaded as-is. Every library lives in `vendor/` (MSAL, SheetJS, html2pdf, jsPDF, JsBarcode, QR code, logo/stamp images as JS) so nothing loads from a CDN.

To run locally, serve the folder over HTTP (e.g. `python3 -m http.server 8000`) and open `http://localhost:8000/`. With no OneDrive folder selected and `clientId` empty, it runs as the offline demo with seeded sample data in `localStorage`.

## Three storage modes, one interface

`index.html` holds all UI and business logic. Screens never talk to storage directly; they call a single object kept under its historical Google Apps Script name **`GAS`**:

```js
GAS.withSuccessHandler(ok).withFailureHandler(ko).<fn>(...)
// fns: getAll, upsertMany(payload), savePdf(html,name,path), exportSheets(path,title,sheets), importAttendance()
```

At startup (`index.html`, `M365.init().then(...)` near `startData`), the first available backend wins:

1. **Microsoft 365** — `m365.js`, active only when `M365_CONFIG.clientId` is set. MSAL sign-in, SharePoint lists via Microsoft Graph. `M365_LISTS` maps each hub collection to a list and its columns (`[hubField, column, kind]`, kinds `t/n/b/d/j/p`). Writes only changed records and PATCHes only changed columns, so columns used by the existing Power Apps are never touched. Dates are written as `yyyy-MM-ddT12:00:00Z`. Choice columns only accept values already in SharePoint.
2. **Shared-folder mode** — `local.js` (currently the mode in real use, "Plan B" until IT registers the app). Uses the File System Access API (Edge/Chrome) on the OneDrive-synced folder `MIDU Hub Data`. Each collection is a JSON file `{v, collection, saved, records:{key:record}, meta}` under `<area folder>/_data/` (`AREAS` / `COLL_AREA` map collections to area folders; unmapped collections fall back to `00 System`). Each save re-reads the file and merges only changed records so concurrent users don't overwrite each other; OneDrive conflict copies are merged on load. Users/roles come from `hub-users` in the folder; `LOCAL.readOnly` (viewer) blocks saving.
3. **Offline demo** — `GAS` stays `null`; `D` is persisted to `localStorage` under `midu_hub3_hub`.

Many screens check `GAS` (any backend) vs `LOCAL.connected` (folder mode) vs `M365.signedIn` to show or hide features (e.g. Excel export, "Save PDF to OneDrive").

## Data model and saving (`index.html`)

- `D` is the whole in-memory database: one array per collection (`items, stock, emps, assets, docs, prs, cqs, pos, invs, ut, dns, punches, mcl, mreq, areq, …`; `budget` is an object). `seed()` builds the demo data; `emptyD()` is the shape used when a backend loads.
- `KEYS` gives each collection's record-key function. `snapshot()` stores a JSON copy of every record in `SNAP`; `persist()` diffs `D` against `SNAP` and sends only changed records through `upsertMany`. **After mutating `D`, call `persist()`** — that is the only save path.
- Adding a new collection: add `KEYS.<coll>=r=>…` and `D.<coll>=D.<coll>||[]` (see Update 50 `pca`, Update 54 `cost`), add it to `COLL_AREA` in `local.js` if it belongs to an area, and to `M365_LISTS` in `m365.js` for SharePoint.
- `loadServer()` reloads everything via `getAll()`; navigating after 5 minutes triggers a reload automatically.

## How `index.html` is organised

- One large inline `<script>`. Screens are `VIEWS.<id>=()=>htmlString` with event wiring in `BIND.<id>=()=>{…}`; `render()` draws `VIEWS[view]` then runs `BIND[view]`. `go(v)` navigates; the view id is mirrored in `location.hash`.
- The file grows by appending numbered **"Update NN"** blocks at the end (`/* ===== Update 54: … ===== */`). Later blocks freely **redefine** earlier `VIEWS`/`BIND` entries and **wrap** `NAV` (`{const N0=NAV;NAV=()=>{const n=N0(); n.splice(…); return n;}}`) to insert menu items. When changing a screen, search for the *last* definition of `VIEWS.<id>` / `BIND.<id>` — earlier ones may be dead. Follow the same pattern (new numbered block at the end) for new features.
- Bilingual UI: every user-visible string goes through `tr(en, ar)`. Data entry stays English-only (`class="en"` fields). `L` is the current language; `applyLang()` re-renders.
- Access control: `M365_ROLES` (in `m365.js`) lists which view ids each role may open; `M365.can(view)` is used by both backends. New screens must be added to the relevant roles or they will be hidden for non-admins.
- Helpers used everywhere: `$` (querySelector), `esc` (HTML-escape — always use it on data), `toast`, `head`, `fld`, `tag`, `today()`, `iso()`.
- Printable documents (GRN/MDS/MRS, PO, debit notes, reports) are HTML built with `.doc*` CSS classes, shown in the modal, and turned into PDFs by html2pdf; `pdfFolder()` picks the destination folder.

## Deploying changes

- After editing `m365.js` or `local.js`, bump the `?v=` query string on their `<script>` tags in `index.html` so browsers fetch the new version.
- `sw.js` caches only app files (network-first, cache fallback). If you add a new file that must work offline, add it to `FILES` there and bump `CACHE` (`midu-hub-vN`).
- `forms/` holds static blank PDF forms linked from the "Blank forms" screen.

## Config

Only the block at the top of `m365.js` is meant to be edited for environment changes: `clientId` (empty ⇒ M365 mode off), `tenantId`, `site`, `docRoot`, plus `M365_WH_MAP` (Asset Register location → warehouse) and `M365_DEFAULT_RATES` (back-charge rates). Warehouses/locations used by the screens are in `WH_PLAN` in `index.html`.
