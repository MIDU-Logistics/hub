/* ==========================================================================
   MIDU Logistics Hub — shared-folder mode (local.js)  ·  Plan B until IT registers the app
   --------------------------------------------------------------------------
   The hub saves its data as JSON files in the OneDrive folder "MIDU Hub Data"
   (shared by Azzam and Turki). Edge / Chrome write to that folder directly
   (File System Access API) and OneDrive syncs it. No Microsoft sign-in, no IT.

   Folder layout (update 18 = "area folders", so each area can be shared separately)
     MIDU Hub Data/00 System/_data/            hub-users, choices, whref, emps, assetTypes, items
     MIDU Hub Data/01 Warehouse & Assets/_data assets, stock, docs        (+ the area's PDFs beside it)
     MIDU Hub Data/02 Utilities/_data          ut, uitems, rates, dns
     MIDU Hub Data/03 Procurement/_data        prs, cqs, pos, invs, budget
     MIDU Hub Data/05 HR/_data                 punches, areq
     MIDU Hub Data/06 Maintenance/_data        mcl, mreq
     MIDU Hub Data/Backups/                    weekly + manual copies of every _data folder (admin only)
     <area>/_data/_history/<yyyy-mm> <user>.json   change history (PR, CQS, PO, invoices, debit notes)
   A person who is not given an area folder in OneDrive simply does not see that area's screens.
   Old layout (update 17 and before), still read until the admin organises it:
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
/* ---------- areas ---------- */
const AREAS = { sys: "00 System", wh: "01 Warehouse & Assets", ut: "02 Utilities", pr: "03 Procurement", hr: "05 HR", mt: "06 Maintenance" };
const AREA_NAMES = { sys: "System (users, item list, employees)", wh: "Warehouse & Assets", ut: "Utilities & debit notes", pr: "Procurement & Finance", hr: "HR attendance", mt: "Maintenance" };
const COLL_AREA = { emps: "sys", assetTypes: "sys", items: "sys", assets: "wh", stock: "wh", docs: "wh", uitems: "ut", ut: "ut", rates: "ut", dns: "ut",
  prs: "pr", cqs: "pr", pos: "pr", invs: "pr", budget: "pr", punches: "hr", areq: "hr", mcl: "mt", mreq: "mt" };
const VIEW_AREA = { dashWH: "wh", grn: "wh", mds: "wh", mrs: "wh", asset: "wh", assetReg: "wh", labels: "wh", count: "wh",
  utDaily: "ut", utRep: "ut", utBC: "ut", utReg: "ut", dashPF: "pr", pr: "pr", cqs: "pr", po: "pr", fin: "pr", capex: "pr",
  dashHR: "hr", att: "hr", dashMT: "mt", mcl: "mt", mreq: "mt" };
LOCAL.layout = 0; LOCAL.areaOK = {};
const dataPath = a => LOCAL.layout === 2 ? AREAS[a] + "/_data" : "data";
const docPath = p => (LOCAL.layout === 2 ? "" : "Documents/") + String(p || "").replace(/^\/+/, "");
LOCAL.isAdmin = () => (LOCAL.roles || []).indexOf("admin") > -1;

async function checkFolder(dir) {
  const sys = await dir.getDirectoryHandle(AREAS.sys, { create: false }).then(d => d.getDirectoryHandle("_data", { create: false })).catch(() => null);
  if (sys) { LOCAL.layout = 2; return sys; }
  const data = await dir.getDirectoryHandle("data", { create: false }).catch(() => null);
  if (data) { LOCAL.layout = 1; return data; }
  const names = []; for await (const [n] of dir.entries()) names.push(n);
  throw new Error("The folder \"" + dir.name + "\" has no Hub data inside (it contains: " + (names.slice(0, 8).join(", ") || "nothing - OneDrive may still be downloading it") + "). Pick the folder named \"MIDU Hub Data\". If you do see it but not \"" + AREAS.sys + "\" inside, ask Logistics to share that folder with you.");
}
/** Silent reconnect on load (works when the browser kept the permission). */
LOCAL.init = async function () {
  if (!LOCAL.supported) return false;
  const h = await savedHandle(); if (!h) return false;
  try { if ((await h.queryPermission({ mode: "readwrite" })) !== "granted") { LOCAL.pending = h; return false; } await checkFolder(h); }
  catch (e) { return false; }
  LOCAL.dir = h; LOCAL.connected = true; LOCAL.user = LOCAL.userName() || "Hub user";
  await LOCAL.signIn(); await LOCAL.offerMigration(); return true;
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
  step("your name"); await LOCAL.signIn();
  step("area folders"); await LOCAL.offerMigration();
  step("done");
  return true;
};
LOCAL.disconnect = async function () { try { await idb("readwrite", s => s.delete("dir")); } catch (e) {} location.reload(); };


/* ---------- Hub users (folder mode): name + role + PIN in data/hub-users.json ----------
   A simple screen lock until IT enables Microsoft sign-in: it keeps each person on the
   screens of their role. It is not strong security — the shared folder still holds all files. */
LOCAL.users = []; LOCAL.roles = null; LOCAL.readOnly = false;
async function pinHash(name, pin) {
  const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode("midu-hub|" + String(name).trim().toLowerCase() + "|" + pin));
  return [...new Uint8Array(buf)].map(b => b.toString(16).padStart(2, "0")).join("");
}
LOCAL.readUsers = async function () {
  const d = await subDir(dataPath("sys"), true); const j = await readJson(d, "hub-users.json");
  LOCAL.users = (j && Array.isArray(j.users)) ? j.users : [];
  LOCAL.users.forEach(u => { if (/\blogistics\b/.test(u.role || "")) u.role = [...new Set(String(u.role).replace(/\blogistics\b/, "utilities,assets").split(",").map(x => x.trim()).filter(Boolean))].join(","); });
  return LOCAL.users;
};
LOCAL.writeUsers = async function (users) {
  const d = await subDir(dataPath("sys"), true); const j = (await readJson(d, "hub-users.json")) || {};
  const byName = {}; (j.users || []).forEach(u => byName[u.name.toLowerCase()] = u); users.forEach(u => byName[u.name.toLowerCase()] = u);
  LOCAL.users = Object.values(byName).sort((a, b) => a.name.localeCompare(b.name));
  await writeFileTo(d, "hub-users.json", JSON.stringify({ v: 1, saved: new Date().toISOString(), by: LOCAL.user, users: LOCAL.users }, null, 1));
};
function lockBox(inner) {
  let o = document.getElementById("hubLock"); if (!o) { o = document.createElement("div"); o.id = "hubLock"; document.body.appendChild(o); }
  o.style.cssText = "position:fixed;inset:0;z-index:90;background:rgba(15,23,42,.72);display:flex;align-items:center;justify-content:center";
  o.innerHTML = `<div style="background:#fff;border-radius:16px;padding:24px 26px;width:min(420px,92vw);box-shadow:0 20px 50px rgba(0,0,0,.3);font-size:14px">${inner}</div>`;
  return o;
}
const T = (en, ar) => (typeof tr === "function" ? tr(en, ar) : en);
const lockInput = "width:100%;padding:9px 10px;border:1px solid #cfd6e3;border-radius:8px;font-size:15px;margin:4px 0 12px;box-sizing:border-box";
const lockBtn = "background:#1B273D;color:#fff;border:0;border-radius:8px;padding:10px 16px;font-weight:700;cursor:pointer;width:100%";
LOCAL.signIn = async function (force) {
  await LOCAL.readUsers();
  let sess = ""; try { sess = force ? "" : (sessionStorage.getItem("midu-hub-session") || ""); } catch (e) {}
  const active = LOCAL.users.filter(u => u.active !== false);
  const done = u => {
    LOCAL.user = u.name; LOCAL.roles = String(u.role || "").toLowerCase().split(/[,;]/).map(s => s.trim()).filter(Boolean);
    LOCAL.readOnly = LOCAL.roles.length > 0 && LOCAL.roles.every(r => r === "viewer");
    try { sessionStorage.setItem("midu-hub-session", u.name); localStorage.setItem("midu-hub-user", u.name); } catch (e) {}
    const o = document.getElementById("hubLock"); if (o) o.remove();
  };
  if (sess) { const u = active.find(x => x.name === sess); if (u) { done(u); return; } }
  if (!active.length) {            // first time: the person connecting becomes the admin
    return new Promise(res => {
      const o = lockBox(`<h3 style="margin:0 0 6px">${T("Set up Hub users", "إعداد مستخدمي المنصة")}</h3>
        <p style="color:#555;margin:0 0 12px">${T("No users yet. You will be the administrator. Choose a 4-digit PIN.", "لا يوجد مستخدمون بعد. ستكون المسؤول. اختر رمزاً من 4 أرقام.")}</p>
        <label>${T("Your name", "اسمك")}</label><input id="luN" lang="en" style="${lockInput}" value="${(LOCAL.userName() || "").replace(/"/g, "")}">
        <label>PIN</label><input id="luP" type="password" inputmode="numeric" maxlength="8" style="${lockInput}">
        <label>${T("Repeat PIN", "أعد الرمز")}</label><input id="luP2" type="password" inputmode="numeric" maxlength="8" style="${lockInput}">
        <div id="luE" style="color:#B42318;min-height:18px"></div><button id="luOk" style="${lockBtn}">${T("Save", "حفظ")}</button>`);
      o.querySelector("#luOk").onclick = async () => {
        const n = o.querySelector("#luN").value.trim(), p = o.querySelector("#luP").value, p2 = o.querySelector("#luP2").value, e = o.querySelector("#luE");
        if (!n) return e.textContent = T("Enter your name.", "أدخل اسمك.");
        if (!/^\d{4,8}$/.test(p)) return e.textContent = T("The PIN must be 4 to 8 digits.", "يجب أن يكون الرمز من 4 إلى 8 أرقام.");
        if (p !== p2) return e.textContent = T("The two PINs are different.", "الرمزان غير متطابقين.");
        const u = { name: n, role: "admin", pinHash: await pinHash(n, p), active: true };
        LOCAL.user = n; await LOCAL.writeUsers([u]); done(u); res();
      };
    });
  }
  return new Promise(res => {
    const last = LOCAL.userName();
    const o = lockBox(`<h3 style="margin:0 0 4px">MIDU Logistics Hub</h3><p style="color:#555;margin:0 0 12px">${T("Choose your name and enter your PIN.", "اختر اسمك وأدخل الرمز.")}</p>
      <label>${T("Name", "الاسم")}</label><select id="luN" style="${lockInput}">${active.map(u => `<option ${u.name === last ? "selected" : ""}>${u.name.replace(/</g, "&lt;")}</option>`).join("")}</select>
      <div id="luSet"></div><label>PIN</label><input id="luP" type="password" inputmode="numeric" maxlength="8" style="${lockInput}" autofocus>
      <div id="luE" style="color:#B42318;min-height:18px"></div><button id="luOk" style="${lockBtn}">${T("Open the Hub", "فتح المنصة")}</button>`);
    const sel = o.querySelector("#luN"), setBox = o.querySelector("#luSet");
    const upd = () => { const u = active.find(x => x.name === sel.value); setBox.innerHTML = u && !u.pinHash ? `<p style="color:#0E7490;margin:0 0 8px">${T("First sign-in: choose your own PIN (4–8 digits) and repeat it.", "أول دخول: اختر رمزك (4–8 أرقام) وأعده.")}</p><label>${T("Repeat PIN", "أعد الرمز")}</label><input id="luP2" type="password" inputmode="numeric" maxlength="8" style="${lockInput}">` : ""; };
    sel.onchange = upd; upd();
    const go = async () => {
      const u = active.find(x => x.name === sel.value), p = o.querySelector("#luP").value, e = o.querySelector("#luE");
      if (!u) return;
      if (!u.pinHash) {
        const p2 = (o.querySelector("#luP2") || {}).value;
        if (!/^\d{4,8}$/.test(p)) return e.textContent = T("The PIN must be 4 to 8 digits.", "يجب أن يكون الرمز من 4 إلى 8 أرقام.");
        if (p !== p2) return e.textContent = T("The two PINs are different.", "الرمزان غير متطابقين.");
        u.pinHash = await pinHash(u.name, p); LOCAL.user = u.name; await LOCAL.writeUsers([u]); done(u); return res();
      }
      if (await pinHash(u.name, p) !== u.pinHash) { e.textContent = T("Wrong PIN.", "رمز غير صحيح."); o.querySelector("#luP").value = ""; return; }
      done(u); res();
    };
    o.querySelector("#luOk").onclick = go; o.querySelector("#luP").onkeydown = ev => { if (ev.key === "Enter") go(); };
  });
};
LOCAL.changeUser = async function () { try { sessionStorage.removeItem("midu-hub-session"); } catch (e) {} await LOCAL.signIn(true); if (typeof applyLang === "function") applyLang(); };

/* ---------- files ---------- */
async function subDir(path, create) {
  let d = LOCAL.dir; const parts = String(path).split("/").map(s => s.trim()).filter(Boolean);
  for (let i = 0; i < parts.length; i++) {
    const top = i === 0 && LOCAL.layout === 2 && create !== "force" && Object.values(AREAS).indexOf(parts[0]) > -1;
    try { d = await d.getDirectoryHandle(parts[i], { create: top ? false : !!create }); }
    catch (e) { if (top && e.name === "NotFoundError") throw new Error(T("You do not have access to the folder", "ليس لديك صلاحية على المجلد") + " \"" + parts[0] + "\" " + T("in MIDU Hub Data. Ask Logistics to share it with you.", "في MIDU Hub Data. اطلب من اللوجستيات مشاركته معك.")); throw e; }
  }
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
  LOCAL.areaOK = {}; const dirs = {};
  for (const a of Object.keys(AREAS)) {
    try { dirs[a] = await subDir(dataPath(a)); LOCAL.areaOK[a] = true; }
    catch (e) { if (a === "sys") throw new Error(T("The folder \"" + AREAS.sys + "\" in MIDU Hub Data is not available on this PC. Ask Logistics to share it with you, then wait for OneDrive to sync.", "المجلد \"" + AREAS.sys + "\" غير متاح على هذا الجهاز. اطلب من اللوجستيات مشاركته معك.")); LOCAL.areaOK[a] = false; }
  }
  const sysDir = dirs.sys;
  const ch = await readJson(sysDir, "choices.json");
  try { window.WHREF = await readJson(sysDir, "whref.json"); } catch (e) { window.WHREF = null; }
  M365.choices = (ch && ch.choices) || {};
  const out = {}; let merged = 0;
  for (const n of LOCAL_COLLS) {
    const dataDir = dirs[COLL_AREA[n]];
    if (!dataDir) { out[n] = []; continue; }               // area not shared with this person
    let c = (await readJson(dataDir, n + ".json")) || emptyColl(n);
    const r = await mergeConflicts(dataDir, n, c); c = r.main;
    if (r.merged) { merged += r.merged; c.saved = new Date().toISOString(); await writeFileTo(dataDir, n + ".json", JSON.stringify(c)); }
    out[n] = Object.values(c.records || {});
    if (AUDIT.indexOf(n) > -1) LOCAL.meta[n] = c.meta || {};
  }
  await histLoad(dirs);
  if (merged) setTimeout(() => toast("✔ " + merged + " " + tr("records merged from a OneDrive conflict copy", "سجل تم دمجه من نسخة تعارض في OneDrive")), 500);
  if (!LOCAL.autoChecked) { LOCAL.autoChecked = true; setTimeout(() => LOCAL.autoBackup().catch(e => console.warn("backup", e)), 2500); }
  out._missing = [];
  return out;
}
async function localUpsertMany(payload) {
  if (LOCAL.readOnly) throw new Error("View only — this user cannot save changes.");
  const now = new Date().toISOString(), result = {}, others = [], hist = {};
  for (const name of Object.keys(payload || {})) {
    if (!KEYS[name] || name === "uitems" || name === "assetTypes") continue;          // reference lists are read-only here too
    const area = COLL_AREA[name] || "sys";
    if (LOCAL.layout === 2 && LOCAL.areaOK[area] === false) throw new Error(T("You do not have access to", "ليس لديك صلاحية على") + " \"" + AREAS[area] + "\"");
    let dataDir;
    try { dataDir = await subDir(dataPath(area)); }
    catch (e) { if (LOCAL.layout === 1 && e.name === "NotFoundError") throw new Error(T("The Hub data was reorganised into area folders. Press F5 to reload the Hub.", "تمت إعادة تنظيم بيانات المنصة. اضغط F5 لإعادة التحميل.")); throw e; }
    const file = (await readJson(dataDir, name + ".json")) || emptyColl(name);
    file.records = file.records || {}; file.meta = file.meta || {};
    SNAP[name] = SNAP[name] || {};
    payload[name].forEach(rec => {
      const k = KEYS[name](rec), old = SNAP[name][k], cur = file.records[k];
      if (cur && old && strip(cur) !== old && strip(cur) !== strip(rec)) others.push((file.meta[k] || {}).by || "another user");
      if (AUDIT.indexOf(name) > -1 && (!cur || strip(cur) !== strip(rec))) {
        const e = histEntry(name, k, cur, rec, now); if (e) (hist[area] = hist[area] || []).push(e);
      }
      file.records[k] = rec; file.meta[k] = { ts: now, by: LOCAL.user };
      if (LOCAL.meta[name]) LOCAL.meta[name][k] = file.meta[k];
    });
    file.saved = now;
    await writeFileTo(dataDir, name + ".json", JSON.stringify(file));
    payload[name].forEach(rec => SNAP[name][KEYS[name](rec)] = JSON.stringify(rec));
    result[name] = payload[name].length;
  }
  for (const a of Object.keys(hist)) { try { await histAppend(a, hist[a]); } catch (e) { console.warn("history", e); toast("⚠ " + tr("Saved, but the change history could not be written: ", "تم الحفظ لكن تعذّر تسجيل سجل التغييرات: ") + e.message); } }
  if (others.length) toast("⚠ " + tr("Saved. Note: ", "تم الحفظ. ملاحظة: ") + others.length + " " + tr("record(s) had also been changed by", "سجل عدّله أيضاً") + " " + [...new Set(others)].join(", "));
  return result;
}

/* ---------- change history (who / what / when) for PR, CQS, PO, invoices and debit notes ----------
   Each person writes only their own monthly file (<area>/_data/_history/2026-11 Azzam El Masri.json),
   so two people saving at the same time never overwrite each other's history.                       */
const AUDIT = ["prs", "cqs", "pos", "invs", "dns"];
LOCAL.hist = []; LOCAL.meta = {};
const histShort = v => { if (v === undefined || v === null || v === "") return ""; const t = typeof v === "object" ? JSON.stringify(v) : String(v); return t.length > 160 ? t.slice(0, 157) + "…" : t; };
function histEntry(coll, key, a, b, ts) {
  const ch = [];
  new Set([...Object.keys(a || {}), ...Object.keys(b || {})]).forEach(f => {
    const x = a ? a[f] : undefined, y = b[f];
    if (JSON.stringify(x === undefined ? null : x) !== JSON.stringify(y === undefined ? null : y)) ch.push({ f, from: a ? histShort(x) : "", to: histShort(y) });
  });
  if (a && !ch.length) return null;
  let act = !a ? "Created" : "Changed";
  if (a && b.status !== a.status && b.status !== undefined) act = "Status: " + (a.status || "—") + " → " + b.status;
  else if (a && coll === "prs" && b.pd && !a.pd) act = "PD approved";
  else if (a && coll === "pos" && b.fin !== a.fin) act = "Finance: " + (a.fin || "—") + " → " + (b.fin || "—");
  else if (a && coll === "invs" && b.paid && !a.paid) act = "Marked paid";
  return { ts, by: LOCAL.user, coll, key, act, ch: !a ? ch.filter(c => ["status", "date", "sup", "contractor", "period", "sub", "cc", "req", "amt"].indexOf(c.f) > -1) : ch };
}
async function histAppend(area, entries) {
  const d = await subDir(dataPath(area) + "/_history", true);
  const fname = new Date().toISOString().slice(0, 7) + " " + String(LOCAL.user || "Hub user").replace(/[\\/:*?"<>|]/g, "_") + ".json";
  const j = (await readJson(d, fname)) || { v: 1, by: LOCAL.user, entries: [] };
  j.entries = (j.entries || []).concat(entries); j.saved = new Date().toISOString();
  await writeFileTo(d, fname, JSON.stringify(j));
  LOCAL.hist.push(...entries);
}
async function histLoad(dirs) {
  const seen = new Set(), all = [];
  const areas = [...new Set(AUDIT.map(c => COLL_AREA[c]))].filter(a => dirs[a]);
  for (const a of areas) {
    const h = await dirs[a].getDirectoryHandle("_history", { create: false }).catch(() => null); if (!h) continue;
    for await (const [fname, fh] of h.entries()) {
      if (fh.kind !== "file" || !/\.json$/i.test(fname)) continue;           // OneDrive conflict copies are read too
      const j = await readJson(h, fname).catch(() => null);
      ((j && j.entries) || []).forEach(e => { const id = e.ts + "|" + e.by + "|" + e.coll + "|" + e.key + "|" + e.act; if (!seen.has(id)) { seen.add(id); all.push(e); } });
    }
  }
  LOCAL.hist = all.sort((x, y) => x.ts < y.ts ? -1 : 1);
}
LOCAL.historyOf = (coll, key) => LOCAL.hist.filter(e => e.coll === coll && e.key === key);
LOCAL.lastSaved = (coll, key) => (LOCAL.meta[coll] || {})[key] || null;

/* ---------- backups: a dated copy of every _data folder in MIDU Hub Data/Backups ----------
   Runs by itself when the admin opens the Hub and the last backup is 7 days old or more (keeps the
   last 12 weekly copies). "Back up now" on the Hub users screen makes an extra copy at any time.    */
async function copyTree(src, dest, stat) {
  for await (const [n, h] of src.entries()) {
    if (h.kind === "file") { const f = await h.getFile(); await writeFileTo(dest, n, f); stat.files++; stat.bytes += f.size; }
    else await copyTree(h, await dest.getDirectoryHandle(n, { create: true }), stat);
  }
  return stat;
}
async function treeStat(dir, stat) {
  stat = stat || { files: 0, bytes: 0 };
  for await (const [, h] of dir.entries()) { if (h.kind === "file") { stat.files++; stat.bytes += (await h.getFile()).size; } else await treeStat(h, stat); }
  return stat;
}
async function verifyFile(h, destDir, name, label, bad) {
  try { const a = await h.getFile(), b = await (await destDir.getFileHandle(name)).getFile(); if (a.size !== b.size) bad.push(label); } catch (e) { bad.push(label); }
}
async function verifyCopy(src, dest, label, bad) {
  for await (const [n, h] of src.entries()) {
    if (h.kind === "file") await verifyFile(h, dest, n, label + "/" + n, bad);
    else { const d = await dest.getDirectoryHandle(n).catch(() => null); if (!d) bad.push(label + "/" + n + "/"); else await verifyCopy(h, d, label + "/" + n, bad); }
  }
}
const stampNow = () => { const d = new Date(), z = n => String(n).padStart(2, "0"); return d.getFullYear() + "-" + z(d.getMonth() + 1) + "-" + z(d.getDate()) + " " + z(d.getHours()) + z(d.getMinutes()); };
LOCAL.lastBackup = async function () {
  const b = await LOCAL.dir.getDirectoryHandle("Backups", { create: false }).catch(() => null); if (!b) return null;
  return await readJson(b, "_last backup.json").catch(() => null);
};
LOCAL.backup = async function (kind, label) {
  if (!LOCAL.isAdmin()) throw new Error(T("Only the Hub administrator can make backups.", "النسخ الاحتياطي للمسؤول فقط."));
  const root = await LOCAL.dir.getDirectoryHandle("Backups", { create: true });
  const name = stampNow() + (label ? " (" + label + ")" : kind === "auto" ? " (weekly)" : " (manual)");
  const dest = await root.getDirectoryHandle(name, { create: true }), stat = { files: 0, bytes: 0 };
  if (LOCAL.layout === 2) {
    for (const a of Object.keys(AREAS)) { const src = await subDir(dataPath(a)).catch(() => null); if (src) await copyTree(src, await dest.getDirectoryHandle(AREAS[a], { create: true }), stat); }
  } else await copyTree(await subDir("data"), await dest.getDirectoryHandle("data", { create: true }), stat);
  const check = await treeStat(dest);
  if (check.files !== stat.files || check.bytes !== stat.bytes) throw new Error(T("The backup copy does not match the original (", "النسخة لا تطابق الأصل (") + check.files + "/" + stat.files + ").");
  const info = { at: new Date().toISOString(), by: LOCAL.user, kind, folder: "Backups/" + name, files: stat.files, bytes: stat.bytes, layout: LOCAL.layout };
  await writeFileTo(dest, "Backup info.json", JSON.stringify(info, null, 1));
  if (kind !== "migration") await writeFileTo(root, "_last backup.json", JSON.stringify(info, null, 1));
  if (kind === "auto") {
    const weekly = []; for await (const [n, h] of root.entries()) if (h.kind === "directory" && / \(weekly\)$/.test(n)) weekly.push(n);
    weekly.sort().reverse().slice(12).forEach(n => root.removeEntry(n, { recursive: true }).catch(() => {}));
  }
  return info;
};
LOCAL.autoBackup = async function () {
  if (!LOCAL.connected || !LOCAL.isAdmin() || LOCAL.readOnly) return null;
  const last = await LOCAL.lastBackup();
  if (last && Date.now() - new Date(last.at).getTime() < 7 * 864e5) return null;
  const info = await LOCAL.backup("auto");
  toast("✔ " + tr("Weekly backup saved: ", "تم حفظ النسخة الاحتياطية الأسبوعية: ") + info.folder + " (" + info.files + " " + tr("files", "ملف") + ")");
  return info;
};

/* ---------- one-time move to area folders (admin, from the old "data" + "Documents" layout) ---------- */
LOCAL.offerMigration = async function () {
  if (LOCAL.layout !== 1 || !LOCAL.isAdmin() || LOCAL.readOnly) return;
  let skip = ""; try { skip = sessionStorage.getItem("midu-hub-nomigrate") || ""; } catch (e) {}
  if (skip) return;
  return new Promise(res => {
    const o = lockBox(`<h3 style="margin:0 0 6px">${T("Organise the Hub data into area folders", "تنظيم بيانات المنصة في مجلدات حسب القسم")}</h3>
      <p style="color:#444;margin:0 0 10px;line-height:1.5">${T("This update puts each area in its own folder inside MIDU Hub Data, so you can share Warehouse with the storekeeper without sharing HR, Procurement or Backups.", "يضع هذا التحديث كل قسم في مجلد خاص داخل MIDU Hub Data لتتمكن من مشاركة المستودع مع أمين المستودع دون مشاركة الموارد البشرية أو المشتريات.")}</p>
      <ul style="margin:0 0 10px;padding-inline-start:18px;color:#444;line-height:1.5"><li>${T("A full copy of the current data is saved in Backups first.", "تُحفظ نسخة كاملة من البيانات الحالية في Backups أولاً.")}</li><li>${T("Ask Turki to close the Hub until you finish (about a minute).", "اطلب من تركي إغلاق المنصة حتى تنتهي (دقيقة تقريباً).")}</li></ul>
      <div id="mgE" style="color:#B42318;min-height:18px;white-space:pre-wrap"></div><div id="mgP" style="color:#0E7490;min-height:18px"></div>
      <button id="mgOk" style="${lockBtn}">${T("Organise now", "نظّم الآن")}</button>
      <button id="mgNo" style="${lockBtn};background:#fff;color:#1B273D;border:1px solid #cfd6e3;margin-top:8px">${T("Later", "لاحقاً")}</button>`);
    o.querySelector("#mgNo").onclick = () => { try { sessionStorage.setItem("midu-hub-nomigrate", "1"); } catch (e) {} o.remove(); res(); };
    o.querySelector("#mgOk").onclick = async () => {
      const b = o.querySelector("#mgOk"), P = o.querySelector("#mgP"); b.disabled = o.querySelector("#mgNo").disabled = true;
      try { const r = await LOCAL.migrate(t => P.textContent = t); o.remove(); toast("✔ " + tr("Data organised into area folders", "تم تنظيم البيانات في مجلدات") + " (" + r.files + " " + tr("files", "ملف") + ")"); res(); }
      catch (e) { console.error(e); o.querySelector("#mgE").textContent = "✖ " + e.message + "\n" + T("Nothing was deleted. The Hub keeps working with the old folders.", "لم يُحذف شيء. تستمر المنصة بالمجلدات القديمة."); b.disabled = false; b.textContent = T("Try again", "حاول مجدداً"); o.querySelector("#mgNo").disabled = false; LOCAL.layout = 1; }
    };
  });
};
LOCAL.migrate = async function (say) {
  say = say || (() => {});
  const root = LOCAL.dir, oldData = await root.getDirectoryHandle("data"), oldDocs = await root.getDirectoryHandle("Documents", { create: false }).catch(() => null);
  say(T("1/4 Backup of the current data…", "1/4 نسخة احتياطية…"));
  await LOCAL.backup("migration", "before area folders");
  say(T("2/4 Copying data files…", "2/4 نسخ ملفات البيانات…"));
  const stat = { files: 0, bytes: 0 }, dataDirs = {};
  for (const a of Object.keys(AREAS)) dataDirs[a] = await (await root.getDirectoryHandle(AREAS[a], { create: true })).getDirectoryHandle("_data", { create: true });
  for await (const [n, h] of oldData.entries()) {
    if (h.kind === "directory") { await copyTree(h, await dataDirs.sys.getDirectoryHandle(n, { create: true }), stat); continue; }
    const coll = n.replace(/(-.+| \(\d+\))?\.json$/i, ""), area = COLL_AREA[coll] || "sys";
    const f = await h.getFile(); await writeFileTo(dataDirs[area], n, f); stat.files++; stat.bytes += f.size;
  }
  say(T("3/4 Moving documents (PDF, Excel)…", "3/4 نقل المستندات…"));
  if (oldDocs) await copyTree(oldDocs, root, stat);
  say(T("4/4 Checking the copies…", "4/4 التحقق من النسخ…"));
  const bad = [];
  for await (const [n, h] of oldData.entries()) {
    if (h.kind === "directory") { await verifyCopy(h, await dataDirs.sys.getDirectoryHandle(n), "data/" + n, bad); continue; }
    const coll = n.replace(/(-.+| \(\d+\))?\.json$/i, ""), area = COLL_AREA[coll] || "sys";
    await verifyFile(h, dataDirs[area], n, "data/" + n, bad);
  }
  if (oldDocs) await verifyCopy(oldDocs, root, "Documents", bad);
  if (bad.length) throw new Error(T("Copy check failed for ", "فشل التحقق من ") + bad.length + " " + T("file(s): ", "ملف: ") + bad.slice(0, 3).join(", "));
  await root.removeEntry("data", { recursive: true });
  if (oldDocs) await root.removeEntry("Documents", { recursive: true });
  LOCAL.layout = 2;
  return stat;
};

/* Screens of an area that is not shared with this person are hidden. */
(function () {
  const can0 = M365.can;
  M365.can = v => {
    if (!M365.signedIn && LOCAL.connected && LOCAL.layout === 2 && VIEW_AREA[v] && LOCAL.areaOK[VIEW_AREA[v]] === false) return false;
    return can0(v);
  };
})();
async function localSaveBlob(path, name, blob) {
  if (LOCAL.readOnly) throw new Error("View only — this user cannot save files.");
  const d = await subDir(docPath(path), true);
  await writeFileTo(d, name, blob);
  return { url: URL.createObjectURL(blob), path: "MIDU Hub Data/" + docPath(path) + "/" + name };
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
  const inbox = await subDir(docPath("05 HR/01 Attendance Inbox"), true), done = await subDir(docPath("05 HR/02 Attendance Processed"), true), errs = await subDir(docPath("05 HR/03 Attendance Errors"), true);
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
LOCAL.readDoc = async function (relPath) {
  const parts = String(relPath).split("/"); const name = parts.pop();
  const d = await subDir(docPath(parts.join("/")), false); const fh = await d.getFileHandle(name); return await fh.getFile();
};
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
