# MIDU Logistics Hub — Microsoft 365 version

The approved hub interface (`MIDU_Unified_Hub.html`) with the Google layer replaced by Microsoft 365:
sign-in with MIDU Microsoft accounts (MSAL.js), data in SharePoint lists, PDFs and Excel files in OneDrive, all through Microsoft Graph with delegated permissions `User.Read` + `Sites.ReadWrite.All`. Zero cost: GitHub Pages hosting, no premium connectors.

## Files

| File | What it is |
|---|---|
| `index.html` | The hub. Every screen, label, print template, the AR/EN toggle and English-only entry are kept. Only the storage layer changed, plus the Utilities screens (now one reading per item per day, like the Power App). |
| `m365.js` | The Microsoft 365 layer: sign-in, list mapping, saving, PDF, Excel, attendance import. **The config block at the top is the only part to edit.** |
| `manifest.webmanifest`, `sw.js`, `icon-*.png` | Makes the hub installable as an app (phone / desktop). The service worker caches only the app files, never data. |
| `vendor/` | MSAL.js 2.39, SheetJS 0.18.5 and html2pdf.js 0.10.2 (copied here so the hub does not depend on outside CDNs). |

## Config (top of `m365.js`)

```js
clientId: "",                                     // from IT after the app registration
tenantId: "464e5e59-26fe-4715-ac16-9efb6ccf733f",
site: "<company SharePoint site>", // set when IT gives the client ID
docRoot: "MIDU Logistics (Claude)/Hub Documents"  // folder for PDFs and Excel exports
```

While `clientId` is empty the hub runs as the offline demo (sample data in the browser only).

## Deploy on GitHub Pages (step 2)

1. Create a GitHub account (free), then a public repository named `midu-hub`.
2. Upload all files of this folder (keep the `vendor` folder).
3. Repository → Settings → Pages → Source: *Deploy from a branch* → `main` / root → Save.
4. The address will be `https://<github-user>.github.io/midu-hub/`.
5. Send that address to IT. They add it as the **Single-page application redirect URI** of the app registration.
6. Put the client ID from IT into `m365.js`, upload it again.

## How saving works

- On sign-in the hub reads every list once. After that it saves **only the records that changed**, and inside a record **only the columns that changed** (PATCH). Columns the hub does not use (rental, operator, warranty, GL…) are never touched, so the Power Apps keep working.
- The top bar shows ✔ Saved to SharePoint / ✖ Not saved.
- Choice columns are filled only with existing values (loaded from SharePoint at sign-in).
- Dates are written as `yyyy-MM-ddT12:00:00Z`, the same as Utility Readings today.

## Existing lists reused

| Hub | SharePoint list | Notes |
|---|---|---|
| assets | MIDU Asset Register | Employee ID = employee number; Status: issue → In Use, return → In Stock / Damaged, open maintenance request → Under Maintenance |
| emps | MIDU Employees | Device ID used for ZKTeco matching once the column is added |
| assetTypes | MIDU Asset Types | read only; Asset Class / GL / Useful Life copied to new assets (blank when Rented) |
| uitems | Utility Items | read only; Category, Location, Charge To |
| ut | Utility Readings | one row per item per day, key DayKey; MonthKey and WeekStart (Saturday) filled on save |

## Not like the Google version

- Attendance files in “05 HR/01 Attendance Inbox” are imported when someone presses **Import now** (there is no server to run a 15-minute timer at zero cost). Uploading a file from the hub works as before and also accepts .xlsx.
- Excel exports have plain headers (the free Excel library cannot colour cells).
- Back charge uses the **Charge To** of each utility item and one rate per category. Generator rental (Byrne + 10 %) and shared-item splits are not in the hub yet.
