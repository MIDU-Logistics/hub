MIDU Hub – update 59: item photos (28 Nov 2026)

Upload BOTH files to github.com/MIDU-Team/hub (Add file > Upload files > Commit):
  index.html
  local.js   (index.html loads it as local.js?v=202611282100)

What changes
- When you pick an item that has a photo (PR lines, GRN / MDS / MRS lines, Item history),
  a small photo shows next to the item box. Click it to see the large photo with code, unit and price.
- Photos are read from MIDU Hub Data/00 System/Item Photos/<item code>.png
  (the "photo" field in items.json). Items without a photo look the same as before.
- Fix: a NEW item code created in the Hub now counts every code already in the item list,
  so it can never repeat an existing code (e.g. next stationery code = MIDU-ADM-STN-019).

Built on the live version from GitHub (local.js v202611282000). Tested in Playwright: PR, GRN, Item history, no errors.
