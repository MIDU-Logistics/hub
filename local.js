/* ==========================================================================
   MIDU Logistics Hub — shared-folder mode (local.js)  ·  Plan B until IT registers the app
   --------------------------------------------------------------------------
   The hub saves its data as JSON files in the OneDrive folder "MIDU Hub Data"
   (shared by Azzam and Turki). Edge / Chrome write to that folder directly
   (File System Access API) and OneDrive syncs it. No Microsoft sign-in, no IT.

   Folder layout
     MIDU Hub Data/data/<collection>.json    one file per collection
         {v:1, collection, saved, records:{key:record}, meta:{key:{ts,by}}}
     MIDU Hub Data/data/choices.json         dropdown values copied from SharePoint
     MIDU Hub Data/Documents/...             PDFs and Excel exports (same folder names as Code.gs)
     MIDU Hub Data/Documents/05 HR/01 Attendance Inbox   ZKTeco files to import

   Two users: every save re-reads the file, merges only the changed records and
   writes it back, so the other person's work is kept. OneDrive conflict copies
   ("assets-PCNAME.json") are merged automatically on the next load.
   ========================================================================== */
const LOCAL = { dir: null, connected: false, user: "", supported: typeof window !== "undefined" && "showDirectoryPicker" in window };

const LOCAL_DB = "midu-hub-folder", LOCAL_STORE = "handles";
function idb(mode, fn) {
  return new Promise((ok, ko) => {
    const rq = indexedDB.open(LOCAL_DB, 1);
    rq.onupgradeneeded = () => rq.result.createObjectStore(LOCAL_STORE);
    rq.onerror = () => ko(rq.error);
    rq.onsuccess = () => { const tx = rq.result.transaction(LOCAL_STORE, mode); const r = fn(tx.objectStore(LOCAL_STORE)); tx.oncomplete = () => ok(r && r.result); tx.onerror = () => ko(tx.error); };
  });
}
async function savedHandle() { try { return await idb("readonly", s => s.get("dir")); } catch (e) { return null; } }

/* ---------- connect ---------- */
LOCAL.userName = () => { try { return localStorage.getItem("midu-hub-user") || ""; } catch (e) { return ""; } };
LOCAL.askUser = () => {
  let n = LOCAL.userName();
  while (!n) { n = (prompt("Your name (shown on the records you save), e.g. Azzam El Masri") || "").trim(); if (!n) break; }
  try { if (n) localStorage.setItem("midu-hub-user", n); } catch (e) {}
  LOCAL.user = n || "Hub user";
};
async function checkFolder(dir) {
  const data = await dir.getDirectoryHandle("data", { create: false }).catch(() => null);
  if (!data) { const names = []; for await (const [n] of dir.entries()) names.push(n); throw new Error("The folder \"" + dir.name + "\" has no \"data\" folder inside (it contains: " + (names.slice(0, 8).join(", ") || "nothing - OneDrive may still be downloading it") + "). Pick the folder named \"MIDU Hub Data\"."); }
  return data;
}
/** Silent reconnect on load (works when the browser kept the permission). */
LOCAL.init = async function () {
  if (!LOCAL.supported) return false;
  const h = await savedHandle(); if (!h) return false;
  try { if ((await h.queryPermission({ mode: "readwrite" })) !== "granted") { LOCAL.pending = h; return false; } await checkFolder(h); }
  catch (e) { return false; }
  LOCAL.dir = h; LOCAL.connected = true; LOCAL.user = LOCAL.userName() || "Hub user"; return true;
};
/** Button: pick the folder (or re-allow the one picked before). Needs a click. */
LOCAL.connect = async function () {
  const step = t => { LOCAL.step = t; try { console.log("[hub] " + t); } catch (e) {} };
  let h = LOCAL.pending || null;
  step("re-allow saved folder");
  if (h) { if ((await h.requestPermission({ mode: "readwrite" })) !== "granted") h = null; }
  if (!h) {
    step("folder window");
    h = await window.showDirectoryPicker({ id: "midu-hub-data" });
    LOCAL.picked = true;
    step("ask to edit files");
    let p = "unknown";
    try { p = await h.requestPermission({ mode: "readwrite" }); } catch (e) { throw new Error("Chrome would not allow editing files in \"" + h.name + "\" (" + (e.name || "") + ": " + e.message + ")."); }
    if (p !== "granted") throw new Error("Editing files in \"" + h.name + "\" was not allowed (answer: " + p + "). Click the button again and choose Edit files / Allow.");
  }
  step("check data folder");
  await checkFolder(h);
  step("remember folder");
  try { await idb("readwrite", s => s.put(h, "dir")); } catch (e) { console.warn(e); }
  LOCAL.dir = h; LOCAL.connected = true; LOCAL.pending = null;
  step("your name"); LOCAL.askUser();
  step("done");
  return true;
};
LOCAL.disconnect = async function () { try { await idb("readwrite", s => s.delete("dir")); } catch (e) {} location.reload(); };

/* ---------- files ---------- */
async function subDir(path, create) {
  let d = LOCAL.dir;
  for (const p of String(path).split("/").map(s => s.trim()).filter(Boolean)) d = await d.getDirectoryHandle(p, { create: !!create });
  return d;
}
async function readJson(dir, name) {
  try { const f = await (await dir.getFileHandle(name)).getFile(); const t = await f.text(); return t ? JSON.parse(t) : null; }
  catch (e) { if (e.name === "NotFoundError") return null; throw new Error(name + ": " + e.message); }
}
async function writeFileTo(dir, name, data) {
  const fh = await dir.getFileHandle(name, { create: true }); const w = await fh.createWritable();
  await w.write(data); await w.close();
}
const emptyColl = name => ({ v: 1, collection: name, saved: "", records: {}, meta: {} });
const strip = r => JSON.stringify(r);

/* Merge OneDrive conflict copies such as "assets-DESKTOP-1A2B.json" or "assets (1).json" into the main file. */
async function mergeConflicts(dataDir, name, main) {
  const copies = [];
  for await (const [fname, h] of dataDir.entries()) {
    if (h.kind !== "file" || fname === name + ".json") continue;
    if (new RegExp("^" + name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "(-.+| \\(\\d+\\))\\.json$").test(fname)) copies.push(fname);
  }
  if (!copies.length) return { main, merged: 0 };
  const arch = await dataDir.getDirectoryHandle("_conflicts", { create: true });
  let merged = 0;
  for (const fname of copies) {
    const c = await readJson(dataDir, fname); if (!c || !c.records) continue;
    Object.keys(c.records).forEach(k => {
      const mt = (main.meta[k] || {}).ts || "", ct = (c.meta && c.meta[k] || {}).ts || "";
      if (!(k in main.records) || ct > mt) { main.records[k] = c.records[k]; main.meta[k] = (c.meta || {})[k] || { ts: ct, by: "" }; merged++; }
    });
    await writeFileTo(arch, new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-") + " " + fname, JSON.stringify(c));
    await dataDir.removeEntry(fname);
  }
  return { main, merged };
}

/* ---------- API used by the hub (same as m365.js) ---------- */
const LOCAL_COLLS = ["emps", "assetTypes", "assets", "uitems", "ut", "items", "stock", "docs", "prs", "cqs", "pos", "invs", "budget", "rates", "dns", "punches", "areq", "mcl", "mreq"];
async function localGetAll() {
  const dataDir = await subDir("data");
  const ch = await readJson(dataDir, "choices.json");
  M365.choices = (ch && ch.choices) || {};
  const out = {}; let merged = 0;
  for (const n of LOCAL_COLLS) {
    let c = (await readJson(dataDir, n + ".json")) || emptyColl(n);
    const r = await mergeConflicts(dataDir, n, c); c = r.main;
    if (r.merged) { merged += r.merged; c.saved = new Date().toISOString(); await writeFileTo(dataDir, n + ".json", JSON.stringify(c)); }
    out[n] = Object.values(c.records || {});
  }
  if (merged) setTimeout(() => toast("✔ " + merged + " " + tr("records merged from a OneDrive conflict copy", "سجل تم دمجه من نسخة تعارض في OneDrive")), 500);
  out._missing = [];
  return out;
}
async function localUpsertMany(payload) {
  const dataDir = await subDir("data"), now = new Date().toISOString(), result = {}, others = [];
  for (const name of Object.keys(payload || {})) {
    if (!KEYS[name] || name === "uitems" || name === "assetTypes") continue;          // reference lists are read-only here too
    const file = (await readJson(dataDir, name + ".json")) || emptyColl(name);
    file.records = file.records || {}; file.meta = file.meta || {};
    SNAP[name] = SNAP[name] || {};
    payload[name].forEach(rec => {
      const k = KEYS[name](rec), old = SNAP[name][k], cur = file.records[k];
      if (cur && old && strip(cur) !== old && strip(cur) !== strip(rec)) others.push((file.meta[k] || {}).by || "another user");
      file.records[k] = rec; file.meta[k] = { ts: now, by: LOCAL.user };
    });
    file.saved = now;
    await writeFileTo(dataDir, name + ".json", JSON.stringify(file));
    payload[name].forEach(rec => SNAP[name][KEYS[name](rec)] = JSON.stringify(rec));
    result[name] = payload[name].length;
  }
  if (others.length) toast("⚠ " + tr("Saved. Note: ", "تم الحفظ. ملاحظة: ") + others.length + " " + tr("record(s) had also been changed by", "سجل عدّله أيضاً") + " " + [...new Set(others)].join(", "));
  return result;
}
async function localSaveBlob(path, name, blob) {
  const d = await subDir("Documents/" + path, true);
  await writeFileTo(d, name, blob);
  return { url: URL.createObjectURL(blob), path: "MIDU Hub Data/Documents/" + path + "/" + name };
}
async function localSavePdf(html, name, path) {
  const body = /<body>([\s\S]*)<\/body>/.exec(html), css = /<style>([\s\S]*?)<\/style>/.exec(html);
  const src = `<div style="width:190mm;background:#fff"><style>${css ? css[1].replace(/@page\{[^}]*\}/g, "") : ""}</style>${body ? body[1] : html}</div>`;
  const blob = await html2pdf().set({ margin: 10, filename: name + ".pdf", image: { type: "jpeg", quality: 0.95 }, html2canvas: { scale: 2, backgroundColor: "#ffffff", scrollY: 0 },
    jsPDF: { unit: "mm", format: "a4", orientation: "portrait" }, pagebreak: { mode: ["css", "legacy"] } }).from(src, "string").outputPdf("blob");
  return localSaveBlob(path, name + ".pdf", new Blob([blob], { type: "application/pdf" }));
}
async function localExportSheets(path, title, sheets) {
  const wb = XLSX.utils.book_new();
  sheets.forEach(s => {
    const ws = XLSX.utils.aoa_to_sheet(s.rows || []); const width = Math.max(1, ...(s.rows || []).map(r => r.length));
    ws["!cols"] = Array.from({ length: width }, (_, c) => ({ wch: Math.min(48, Math.max(10, ...(s.rows || []).map(r => String(r[c] === undefined ? "" : r[c]).length + 2))) }));
    for (const ref in ws) if (ref[0] !== "!" && typeof ws[ref].v === "number") ws[ref].z = "#,##0.00";
    XLSX.utils.book_append_sheet(wb, ws, String(s.name).replace(/[\\\/\?\*\[\]:]/g, " ").substring(0, 31));
  });
  const buf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const r = await localSaveBlob(path, title + ".xlsx", new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  return { url: r.url, name: r.path };
}
async function localImportAttendance() {
  const inbox = await subDir("Documents/05 HR/01 Attendance Inbox", true), done = await subDir("Documents/05 HR/02 Attendance Processed", true), errs = await subDir("Documents/05 HR/03 Attendance Errors", true);
  const res = { files: 0, imported: 0, duplicates: 0, unmatched: 0, errors: 0 }, unmatched = [], before = D.punches.length, names = [];
  for await (const [fname, h] of inbox.entries()) if (h.kind === "file") names.push(fname);
  const st = new Date().toISOString().slice(0, 16).replace(/[-:T]/g, "");
  for (const fname of names) {
    res.files++;
    const fh = await inbox.getFileHandle(fname), file = await fh.getFile(), buf = await file.arrayBuffer();
    try {
      const m = matchPunches(readPunchFile(buf, fname), fname);
      res.imported += m.imported; res.duplicates += m.duplicates; res.unmatched += m.unmatched; unmatched.push(...m.unmatchedRows);
      await writeFileTo(done, st + " " + fname, buf);
    } catch (e) { res.errors++; await writeFileTo(errs, st + " " + fname, buf); }
    await inbox.removeEntry(fname);
  }
  if (unmatched.length) await writeFileTo(errs, "Attendance Unmatched " + st + ".csv", "Device user ID,Punch time,Device,File\r\n" + unmatched.map(r => r.map(c => '"' + String(c).replace(/"/g, '""') + '"').join(",")).join("\r\n"));
  const added = D.punches.slice(before);
  if (added.length) await localUpsertMany({ punches: added });
  return res;
}
async function localUploadDoc(path, name, blob) { return (await localSaveBlob(path, name, blob)).path; }

(function () {
  const fns = { getAll: localGetAll, upsertMany: localUpsertMany, savePdf: localSavePdf, exportSheets: localExportSheets, importAttendance: localImportAttendance };
  const chain = (ok, ko) => {
    const o = { withSuccessHandler: f => chain(f, ko), withFailureHandler: f => chain(ok, f) };
    Object.keys(fns).forEach(k => o[k] = (...a) => { fns[k](...a).then(r => ok && ok(r)).catch(e => { console.error(e); ko && ko(e); }); });
    return o;
  };
  LOCAL.run = chain(null, null);
  LOCAL.uploadDoc = localUploadDoc;
})();
