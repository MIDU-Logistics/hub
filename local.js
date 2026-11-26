/* ==========================================================================
   MIDU-Team Hub — shared-folder mode (local.js)  ·  Plan B until IT registers the app
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
  await LOCAL.signIn(); await LOCAL.offerMigration(); await LOCAL.weeklyInit().catch(() => {}); return true;
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
  await LOCAL.weeklyInit().catch(() => {});
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
    const o = lockBox(`<h3 style="margin:0 0 4px">MIDU-Team Hub</h3><p style="color:#555;margin:0 0 12px">${T("Choose your name and enter your PIN.", "اختر اسمك وأدخل الرمز.")}</p>
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
  if (!LOCAL.autoChecked) { LOCAL.autoChecked = true; setTimeout(() => LOCAL.autoBackup().catch(e => console.warn("backup", e)).then(() => LOCAL.xlDaily()).catch(e => console.warn("excel", e)), 2500); }
  out._missing = [];
  return out;
}
async function localUpsertMany(payload) {
  if (LOCAL.readOnly) throw new Error("View only — this user cannot save changes.");
  const now = new Date().toISOString(), result = {}, others = [], hist = {}, renamed = []; let merged = 0;
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
      /* (26 Nov 2026) two PCs saved a NEW document with the same number at the same time: keep both, give ours the next free number */
      if (name === "docs" && rec.no && !SNAP[name][rec.no] && file.records[rec.no] && strip(file.records[rec.no]) !== strip(rec) && /-\d+$/.test(rec.no)) {
        const old = rec.no, base = old.replace(/-\d+$/, "-"), w = old.match(/-(\d+)$/)[1].length;
        let n = +old.match(/-(\d+)$/)[1]; while (file.records[base + String(n).padStart(w, "0")]) n++;
        rec.no = base + String(n).padStart(w, "0"); renamed.push(old + " → " + rec.no);
      }
      /* stock is a quantity: if another PC changed the same stock line since we loaded it, add our change to theirs instead of overwriting */
      if (name === "stock") { const k0 = KEYS[name](rec), o0 = SNAP[name][k0], c0 = file.records[k0];
        if (o0 && c0 && strip(c0) !== o0) { try { const was = +JSON.parse(o0).qty || 0; rec.qty = (+c0.qty || 0) + ((+rec.qty || 0) - was); merged++; } catch (e) {} } }
      const k = KEYS[name](rec), old = SNAP[name][k], cur = file.records[k];
      if (cur && old && strip(cur) !== old && strip(cur) !== strip(rec) && name !== "stock") others.push((file.meta[k] || {}).by || "another user");
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
  try { xlSchedule([...new Set(Object.keys(result).map(n => COLL_AREA[n] || "sys"))], (payload.ut || []).map(u => String(u.date || "").slice(0, 7)).filter(Boolean)); } catch (e) { console.warn(e); }
  for (const a of Object.keys(hist)) { try { await histAppend(a, hist[a]); } catch (e) { console.warn("history", e); toast("⚠ " + tr("Saved, but the change history could not be written: ", "تم الحفظ لكن تعذّر تسجيل سجل التغييرات: ") + e.message); } }
  if (renamed.length) { toast("⚠ " + T("The document number was just used on another PC – saved as ", "رقم المستند استُخدم للتو على جهاز آخر – حُفظ باسم ") + renamed.join(", ")); try { if (typeof render === "function") render(); } catch (e) {} }
  if (merged) console.info("stock lines merged with another PC's change:", merged);
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

/* ======================= Excel linked with the Hub (both ways) =======================
   A. Utilities weekly workbook (Turki's layout), one per month, linked BOTH ways:
      - in the Hub folder:  02 Utilities/Excel (linked with Hub)/Weekly_Report_Utilities_Back_Charge_MIDU (<Month yyyy>).xlsx
      - in Turki's own folder (Utilities › Weekly Reports For Sub-contractors › <Month>/…), once someone links that
        folder on their PC (button on the Utilities screens). His own file is used, so his sheets, the Link sheet and
        the "Monthly Utilities Consumption" file that reads it keep working.
      Items that are not rows of his sheet (Laydown rows, extra equipment …) are on an added sheet "Hub extra items".
   B. Every area: <area>/Excel (linked with Hub)/<area> - Hub data.xlsx – one sheet per list. Columns marked ✎ can be
      edited in Excel and come back into the Hub; other columns are overwritten (they follow the Hub's own rules:
      stock and custody only through GRN/MDS/MRS, PR approvals in the Hub …).
   How a change is recognised: the Hub keeps what it last wrote in <area>/_data/_excel-sync.json. A cell that differs
   from that in Excel was typed in Excel → saved into the Hub; then the Hub writes the file again with everything.
   The Hub checks the files every 30 s while it is open, when its window gets focus, when it opens, and before every
   write – so an Excel change is never overwritten.                                                                 */
const XL_DIR = "Excel (linked with Hub)", XL_TPL = "Excel templates", XL_TPL_NAME = "Utilities weekly template";
const XL_MON = ["January","February","March","April","May","June","July","August","September","October","November","December"];
const XL_EXTRA = "Hub extra items", XL_EDIT_MARK = " ✎";
const XL_COLS = {
  emps: [["no", "Employee No"], ["name", "Name"], ["dept", "Department"], ["site", "Site"], ["iqama", "Iqama / ID"], ["mobile", "Mobile"], ["deviceId", "Device ID"]],
  items: [["code", "Item code"], ["desc", "Description"], ["descAr", "Description (AR)"], ["unit", "Unit"], ["kind", "Kind"], ["cat", "Category"], ["dept", "Department"], ["min", "Min qty"], ["price", "Unit price"], ["created", "Created"], ["createdBy", "Created by"], ["src", "Source"]],
  assets: [["id", "Asset ID"], ["type", "Type"], ["desc", "Description"], ["brand", "Brand"], ["model", "Model"], ["serial", "Serial No"], ["itemCode", "Item code"], ["status", "Status"], ["wh", "Warehouse"], ["loc", "Location"], ["own", "Ownership"], ["custName", "Holder"], ["cust", "Holder ID"]],
  stock: [["code", "Item code"], ["wh", "Warehouse"], ["loc", "Storage location"], ["sl", "Stock line"], ["qty", "Qty"]],
  docs: [["no", "Doc No"], ["type", "Type"], ["date", "Date"], ["dtype", "Dispatch type"], ["party", "Party"], ["iqama", "Employee / Iqama"], ["po", "PO No"], ["dn", "Delivery note"], ["wh", "Warehouse"], ["loc", "Location"], ["to", "To"], ["returnBy", "Return by"], ["by", "By"], ["remarks", "Remarks"]],
  ut: [["id", "Reading ID"], ["date", "Date"], ["item", "Item key"], ["title", "Item"], ["type", "Category"], ["c", "Location / Charge to"], ["company", "Company"], ["qty", "Qty"], ["unit", "Unit"], ["ref", "Reference"]],
  uitems: [["key", "Item key"], ["title", "Item"], ["cat", "Category"], ["loc", "Location"], ["unit", "Unit"], ["company", "Company"], ["chargeTo", "Charge to"], ["concretePct", "Concrete %"], ["mccPct", "MCC %"], ["emcoPct", "EMCO %"], ["active", "Active"]],
  rates: [["type", "Type"], ["unit", "Unit"], ["rate", "Rate (before VAT)"]],
  dns: [["no", "Debit note"], ["contractor", "Contractor"], ["period", "Period"], ["date", "Date"], ["sub", "Subtotal"], ["vat", "VAT"], ["status", "Status"], ["paid", "Paid"]],
  prs: [["no", "PR No"], ["date", "Date"], ["cc", "Cost center"], ["req", "Requested by"], ["need", "Need date"], ["just", "Purpose"], ["prio", "Priority"], ["status", "Status"], ["pd", "PD approved"], ["pdBy", "PD approved by"], ["pdDate", "PD approval date"], ["rejReason", "Reason not approved"]],
  cqs: [["no", "CQS No"], ["pr", "PR No"], ["date", "Date"], ["sel", "Selected supplier"], ["approved", "Approved"]],
  pos: [["no", "PO No"], ["kind", "Kind"], ["date", "Date"], ["pr", "PR No"], ["cqs", "CQS No"], ["sup", "Supplier"], ["cc", "Cost center"], ["fin", "Finance status"]],
  invs: [["po", "PO No"], ["no", "Invoice No"], ["date", "Invoice date"], ["amt", "Amount"], ["paid", "Paid date"]],
  budget: [["cc", "Cost center"], ["CAPEX", "CAPEX"], ["OPEX", "OPEX"]],
  punches: [["emp", "Employee No"], ["ts", "Punch time"], ["dev", "Device"], ["src", "Source file"]],
  areq: [["no", "Request No"], ["emp", "Employee No"], ["type", "Type"], ["date", "Date"], ["from", "From"], ["to", "To"], ["leaveType", "Leave type"], ["inTime", "In"], ["outTime", "Out"], ["reason", "Reason"], ["status", "Status"], ["requestedBy", "Requested by"], ["decidedBy", "Decided by"], ["decidedOn", "Decided on"]],
  mcl: [["no", "Checklist No"], ["month", "Month"], ["loc", "Location"], ["by", "Checked by"], ["date", "Date"]],
  mreq: [["no", "Request No"], ["date", "Date"], ["by", "Requested by"], ["asset", "Asset"], ["loc", "Location"], ["cat", "Category"], ["prio", "Priority"], ["desc", "Description"], ["status", "Status"], ["to", "Assigned to"]]
};
const XL_NAMES = { emps: "Employees", items: "Item list", assets: "Asset register", stock: "Stock", docs: "GRN MDS MRS", ut: "Utility readings", uitems: "Utility items", rates: "Utility rates", dns: "Debit notes",
  prs: "Purchase requests", cqs: "CQS", pos: "PO and petty cash", invs: "Invoices", budget: "Budget", punches: "Attendance punches", areq: "Attendance requests", mcl: "Monthly checklists", mreq: "Maintenance requests", assetTypes: "Asset types" };
const XL_LINES = { docs: "no", prs: "no", pos: "no", cqs: "no", dns: "no" };
/* Columns that may be changed in Excel (the same fields the Hub screens let people edit directly). add = new rows allowed. */
const XL_EDIT = {
  emps: { f: ["name", "dept", "site", "iqama", "mobile", "deviceId"], add: true },
  items: { f: ["desc", "descAr", "unit", "min", "price"] },
  assets: { f: ["desc", "brand", "model", "serial", "status", "wh", "loc", "own"] },
  ut: { f: ["qty", "ref"] },
  rates: { f: ["rate"] },
  dns: { f: ["status", "paid"] },
  budget: { f: ["CAPEX", "OPEX"], add: true },
  mreq: { f: ["status", "to", "prio"] }
};
const XL_NUM = new Set(["min", "price", "rate", "paid", "CAPEX", "OPEX", "qty", "sub", "vat", "amt"]);
const xlCell = v => v === undefined || v === null ? "" : typeof v === "object" ? JSON.stringify(v) : v;
const xlNorm = (f, v) => { if (v === undefined || v === null) return ""; if (XL_NUM.has(f)) { if (v === "") return ""; const n = +v; return isNaN(n) ? String(v).trim() : n; } return String(v).trim(); };
const xlKeyOf = (c, r) => String(KEYS[c](r));
const xlRows = (data, c) => data === (typeof D !== "undefined" ? D : null) && typeof rowsOf === "function" ? rowsOf(c) : (Array.isArray(data[c]) ? data[c] : Object.entries(data[c] || {}).map(([cc, v]) => Object.assign({ cc }, v)));

/* ---------------- small zip / sheet helpers ---------------- */
const xlColName = n => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const xlSerial = ymd => Math.round((Date.UTC(+ymd.slice(0, 4), +ymd.slice(5, 7) - 1, +ymd.slice(8, 10)) - Date.UTC(1899, 11, 30)) / 864e5);
const xlFromSerial = n => { const d = new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 864e5); return d.toISOString().slice(0, 10); };
const xlEsc = t => String(t).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const zOpen = u8 => XLSX.CFB.read(u8, { type: "array" });
const zFind = (z, p) => XLSX.CFB.find(z, "/" + p.replace(/^\//, ""));
const zText = (z, p) => { const e = zFind(z, p); return e ? new TextDecoder().decode(e.content) : null; };
const zPut = (z, p, s) => { const e = zFind(z, p), b = new TextEncoder().encode(s); if (e) { e.content = b; e.size = b.length; } else XLSX.CFB.utils.cfb_add(z, "/" + p, b); };
const zBytes = z => XLSX.CFB.write(z, { fileType: "zip", type: "array", compression: true });
/** sheet name → part path ("xl/worksheets/sheet2.xml") */
function zSheets(z) {
  const wb = zText(z, "xl/workbook.xml"), rels = zText(z, "xl/_rels/workbook.xml.rels"), out = {};
  const tgt = {}; (rels.match(/<Relationship [^>]*>/g) || []).forEach(r => { const id = /Id="([^"]+)"/.exec(r), t = /Target="([^"]+)"/.exec(r); if (id && t) tgt[id[1]] = t[1].replace(/^\/?xl\//, "").replace(/^\//, ""); });
  (wb.match(/<sheet [^>]*>/g) || []).forEach(s => { const n = /name="([^"]+)"/.exec(s), id = /r:id="([^"]+)"/.exec(s); if (n && id && tgt[id[1]]) out[n[1].replace(/&amp;/g, "&")] = "xl/" + tgt[id[1]]; });
  return out;
}
function xlSetCell(rowXml, ref, val) {
  const re = new RegExp('<c r="' + ref + '"((?: [\\w:]+="[^"]*")*)\\s*(?:/>|>[\\s\\S]*?</c>)');
  const m = re.exec(rowXml); if (!m) return null;
  const attrs = m[1].replace(/ t="[^"]*"/, "");
  const cell = val === null || val === undefined || val === "" ? '<c r="' + ref + '"' + attrs + '/>' : '<c r="' + ref + '"' + attrs + '><v>' + (+val) + '</v></c>';
  return rowXml.slice(0, m.index) + cell + rowXml.slice(m.index + m[0].length);
}
const xlRowRe = r => new RegExp('<row r="' + r + '"[^>]*>[\\s\\S]*?</row>');

/* ---------------- A. utilities weekly workbook ---------------- */
const xlUtilName = month => "Weekly_Report_Utilities_Back_Charge_MIDU (" + XL_MON[+month.slice(5, 7) - 1] + " " + month.slice(0, 4) + ").xlsx";
const xlMonthOf = serial => { const d = xlFromSerial(serial); return d.slice(0, 7); };
/** Find the row of every Hub utility item in "Daily Utilities Filling" (by name and location, in order). */
function xlUtilInfo(u8) {
  const wb = XLSX.read(u8, { type: "array" }), ws = wb.Sheets["Daily Utilities Filling"], su = wb.Sheets.Setup;
  if (!ws || !su) throw new Error("Not a utilities weekly workbook (sheets 'Setup' / 'Daily Utilities Filling' missing)");
  const b4 = su.B4, month = b4 && typeof b4.v === "number" ? xlMonthOf(b4.v) : "";
  const n = s => String(s === undefined || s === null ? "" : s).toLowerCase().replace(/[^a-z0-9]/g, "");
  const items = (D.uitems || []).slice().sort((a, b) => (+a.sort || 0) - (+b.sort || 0)), rows = {}, unmatched = [];
  const rng = XLSX.utils.decode_range(ws["!ref"]); let j = 0;
  for (let r = 8; r <= rng.e.r + 1; r++) {
    const A = (ws["A" + r] || {}).v, B = (ws["B" + r] || {}).v, C = (ws["C" + r] || {}).v;
    if (!A || /^total/i.test(String(A)) || (!B && !C)) continue;
    let k = -1;
    for (let i = j; i < items.length; i++) { const it = items[i]; if (n(it.title) === n(A) && n(it.loc) === n(B) && (!it.company || n(it.company) === n(C))) { k = i; break; } }
    if (k < 0) { unmatched.push(r + ": " + A + " / " + (B || "")); continue; }
    rows[items[k].key] = r; j = k + 1;
  }
  const extra = items.filter(i => !rows[i.key]);
  return { month, rows, extra, unmatched, sheets: null };
}
/** Values typed in the workbook: {month, cells:{"U001|2026-11-21":qty}, skip:Set(keys with text)} */
function xlReadUtilities(u8) {
  const info = xlUtilInfo(u8), wb = XLSX.read(u8, { type: "array" }), ws = wb.Sheets["Daily Utilities Filling"], ex = wb.Sheets[XL_EXTRA];
  const month = info.month, days = new Date(Date.UTC(+month.slice(0, 4), +month.slice(5, 7), 0)).getUTCDate(), cells = {}, skip = new Set();
  const take = (key, c) => { if (!c || c.v === "" || c.v === null || c.v === undefined) return; if (typeof c.v === "number") cells[key] = c.v; else skip.add(key); };
  Object.keys(info.rows).forEach(k => { for (let d = 1; d <= days; d++) take(k + "|" + month + "-" + String(d).padStart(2, "0"), ws[xlColName(3 + d) + info.rows[k]]); });
  if (ex && ex["!ref"]) { const rg = XLSX.utils.decode_range(ex["!ref"]);
    for (let r = 2; r <= rg.e.r + 1; r++) { const k = (ex["A" + r] || {}).v; if (!k || info.rows[k]) continue; for (let d = 1; d <= days; d++) take(k + "|" + month + "-" + String(d).padStart(2, "0"), ex[xlColName(5 + d) + r]); } }
  return { month, cells, skip, info };
}
/** Sheet XML for the "Hub extra items" sheet. */
function xlExtraSheet(items, month, val) {
  const days = new Date(Date.UTC(+month.slice(0, 4), +month.slice(5, 7), 0)).getUTCDate();
  const s = (ref, t) => '<c r="' + ref + '" t="inlineStr"><is><t>' + xlEsc(t) + '</t></is></c>', nC = (ref, v) => '<c r="' + ref + '"><v>' + v + '</v></c>';
  let x = '<row r="1">' + ["Key", "Item", "Category", "Location / Charge to", "Unit"].map((h, i) => s(xlColName(i + 1) + 1, h)).join("");
  for (let d = 1; d <= 31; d++) x += d <= days ? nC(xlColName(5 + d) + 1, d) : "";
  x += s(xlColName(37) + 1, "Month total") + "</row>";
  items.forEach((it, i) => { const r = i + 2; x += '<row r="' + r + '">' + [it.key, it.title, it.cat, [it.loc, it.company].filter(Boolean).join(" – "), it.unit].map((t, c) => s(xlColName(c + 1) + r, t || "")).join("");
    for (let d = 1; d <= days; d++) { const v = val(it.key + "|" + month + "-" + String(d).padStart(2, "0")); if (v) x += nC(xlColName(5 + d) + r, v); }
    x += '<c r="' + xlColName(37) + r + '"><f>SUM(F' + r + ':AJ' + r + ')</f></c></row>'; });
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<sheetViews><sheetView workbookViewId="0"><pane xSplit="5" ySplit="1" topLeftCell="F2" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>' +
    '<cols><col min="1" max="1" width="8" customWidth="1"/><col min="2" max="2" width="26" customWidth="1"/><col min="3" max="3" width="20" customWidth="1"/><col min="4" max="4" width="26" customWidth="1"/><col min="5" max="5" width="8" customWidth="1"/><col min="6" max="36" width="7" customWidth="1"/><col min="37" max="37" width="12" customWidth="1"/></cols>' +
    '<sheetData>' + x + '</sheetData></worksheet>';
}
function zEnsureSheet(z, name) {
  const have = zSheets(z); if (have[name]) return have[name];
  let n = 1; while (zFind(z, "xl/worksheets/sheet" + n + ".xml")) n++;
  const part = "xl/worksheets/sheet" + n + ".xml";
  let rels = zText(z, "xl/_rels/workbook.xml.rels"), id = 1; while (rels.indexOf('Id="rId' + id + '"') > -1) id++;
  rels = rels.replace("</Relationships>", '<Relationship Id="rId' + id + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + n + '.xml"/></Relationships>');
  zPut(z, "xl/_rels/workbook.xml.rels", rels);
  let wb = zText(z, "xl/workbook.xml"); const ids = (wb.match(/sheetId="(\d+)"/g) || []).map(x => +x.replace(/\D/g, ""));
  wb = wb.replace("</sheets>", '<sheet name="' + xlEsc(name) + '" sheetId="' + (Math.max(0, ...ids) + 1) + '" r:id="rId' + id + '"/></sheets>');
  zPut(z, "xl/workbook.xml", wb);
  let ct = zText(z, "[Content_Types].xml");
  ct = ct.replace("</Types>", '<Override PartName="/' + part + '" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
  zPut(z, "[Content_Types].xml", ct);
  return part;
}
/** Fill a utilities workbook (Turki's file or the template) with the Hub readings of one month. */
function xlFillUtilities(u8, month, readings) {
  const info = xlUtilInfo(u8), z = zOpen(u8), sh = zSheets(z), val = {};
  readings.filter(u => String(u.date).slice(0, 7) === month && +u.qty).forEach(u => { const k = u.item + "|" + String(u.date).slice(0, 10); val[k] = (val[k] || 0) + (+u.qty); });
  const days = new Date(Date.UTC(+month.slice(0, 4), +month.slice(5, 7), 0)).getUTCDate();
  /* daily sheet */
  let xml = zText(z, sh["Daily Utilities Filling"]);
  Object.keys(info.rows).forEach(k => { const r = info.rows[k], m = xlRowRe(r).exec(xml); if (!m) return; let row = m[0];
    for (let d = 1; d <= 31; d++) { const nr = xlSetCell(row, xlColName(3 + d) + r, d <= days ? val[k + "|" + month + "-" + String(d).padStart(2, "0")] : null); if (nr !== null) row = nr; }
    xml = xml.slice(0, m.index) + row + xml.slice(m.index + m[0].length); });
  zPut(z, sh["Daily Utilities Filling"], xml);
  /* Setup!B4 = first day of the month */
  let sx = zText(z, sh.Setup); const rm = xlRowRe(4).exec(sx);
  if (rm) { const nr = xlSetCell(rm[0], "B4", xlSerial(month + "-01")); if (nr) sx = sx.slice(0, rm.index) + nr + sx.slice(rm.index + rm[0].length); }
  zPut(z, sh.Setup, sx);
  /* Link sheet: refresh the saved results, so files that read it while it is closed (Monthly Utilities Consumption) see the new figures */
  if (sh.Link) {
    const dv = {}; (xml.match(/<c r="([A-Z]+)(\d+)"[^>]*><v>([^<]*)<\/v><\/c>/g) || []).forEach(c => { const m = /<c r="([A-Z]+)(\d+)"[^>]*><v>([^<]*)<\/v>/.exec(c); (dv[m[2]] = dv[m[2]] || {})[m[1]] = +m[3]; });
    let lx = zText(z, sh.Link);
    lx = lx.replace(/<row r="(\d+)"[^>]*>[\s\S]*?<\/row>/g, row => {
      let sum = 0, any = false;
      row = row.replace(/(<c r="[A-Z]+\d+"[^>]*><f>IF\((\d+)&gt;Setup!\$B\$5,0,N\(INDEX\('Daily Utilities Filling'!\$D\$(\d+):\$AH\$\d+,DAY\(Setup!\$B\$4\+(\d+)\)\)\)\)<\/f>)(?:<v>[^<]*<\/v>)?<\/c>/g, (all, head, nDay, dr, off) => {
        const v = +nDay > days ? 0 : ((dv[dr] || {})[xlColName(4 + (+off))] || 0); sum += v; any = true; return head + "<v>" + v + "</v></c>"; });
      if (any) row = row.replace(/(<c r="AH\d+"[^>]*><f[^>]*(?:\/>|>[^<]*<\/f>))(?:<v>[^<]*<\/v>)?<\/c>/, (all, head) => head + "<v>" + sum + "</v></c>");
      return row; });
    zPut(z, sh.Link, lx);
  }
  /* items that are not rows of the sheet */
  if (info.extra.length) zPut(z, zEnsureSheet(z, XL_EXTRA), xlExtraSheet(info.extra, month, k => val[k]));
  let wx = zText(z, "xl/workbook.xml"); if (!/fullCalcOnLoad/.test(wx)) { wx = wx.replace(/<calcPr([^>]*?)\/>/, '<calcPr$1 fullCalcOnLoad="1"/>'); zPut(z, "xl/workbook.xml", wx); }
  return { bytes: zBytes(z), cells: val, info };
}

/* ---------------- B. area workbooks ---------------- */
function xlAreaBook(area, data) {
  const wb = XLSX.utils.book_new();
  const info = [["MIDU-Team Hub – " + AREAS[area]], [],
    ["This file is linked with the Hub. Columns marked ✎ can be changed here: the change goes into the Hub within about 30 seconds while the Hub is open (or when it is next opened)."],
    ["Other columns are written by the Hub after every save – change them in the Hub (stock and custody only through GRN / MDS / MRS, approvals in the Hub)."],
    ["New rows can be added only in Employees and Budget. Rows deleted here are not deleted in the Hub."]];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(info), "About");
  Object.keys(COLL_AREA).filter(c => COLL_AREA[c] === area && (xlRows(data, c).length || (XL_EDIT[c] && XL_EDIT[c].add))).forEach(c => {
    const recs = xlRows(data, c), cols = (XL_COLS[c] || Object.keys(recs[0] || {}).filter(k => !Array.isArray(recs[0][k])).map(k => [k, k])).slice();
    const extra = [...new Set(recs.flatMap(r => Object.keys(r)))].filter(k => !cols.some(x => x[0] === k) && !Array.isArray((recs.find(r => r[k] !== undefined) || {})[k]) && typeof (recs.find(r => r[k] !== undefined) || {})[k] !== "object" && !/^(_|asset$|res$|quotes$)/.test(k));
    const all = cols.concat(extra.map(k => [k, k])), ed = (XL_EDIT[c] || { f: [] }).f;
    const rows = [all.map(([k, h]) => ed.indexOf(k) > -1 ? h + XL_EDIT_MARK : h)].concat(recs.map(r => all.map(([k]) => xlCell(r[k]))));
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = all.map(([, h], i) => ({ wch: Math.min(40, Math.max(8, h.length + 3, ...rows.slice(1, 200).map(r => String(r[i]).length + 1))) }));
    ws["!autofilter"] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(1, rows.length - 1), c: all.length - 1 } }) };
    XLSX.utils.book_append_sheet(wb, ws, (XL_NAMES[c] || c).slice(0, 31));
    if (XL_LINES[c] && recs.some(r => Array.isArray(r.lines) && r.lines.length)) {
      const lk = [...new Set(recs.flatMap(r => (r.lines || []).flatMap(l => Object.keys(l))))];
      const lr = [[XL_COLS[c] ? XL_COLS[c][0][1] : "No", "Line"].concat(lk)].concat(recs.flatMap(r => (r.lines || []).map((l, i) => [r[XL_LINES[c]], i + 1].concat(lk.map(k => xlCell(l[k]))))));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(lr), ((XL_NAMES[c] || c) + " lines").slice(0, 31));
    }
  });
  return XLSX.write(wb, { bookType: "xlsx", type: "array", compression: true });
}
/** What the Hub writes for the editable columns: {coll:{key:{field:value}}} */
function xlAreaSnap(area, data) {
  const out = {};
  Object.keys(XL_EDIT).filter(c => COLL_AREA[c] === area).forEach(c => { out[c] = {}; xlRows(data, c).forEach(r => { const o = {}; XL_EDIT[c].f.forEach(f => o[f] = xlNorm(f, r[f])); out[c][xlKeyOf(c, r)] = o; }); });
  return out;
}
/** Read the editable columns from an area workbook: {coll:{key:{field:value}}} */
function xlAreaRead(u8, area) {
  const wb = XLSX.read(u8, { type: "array" }), out = {};
  Object.keys(XL_EDIT).filter(c => COLL_AREA[c] === area).forEach(c => {
    const ws = wb.Sheets[(XL_NAMES[c] || c).slice(0, 31)]; if (!ws) return;
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" }); if (!aoa.length) return;
    const lab = {}; (XL_COLS[c] || []).forEach(([k, h]) => { lab[h] = k; lab[h + XL_EDIT_MARK] = k; lab[(h + XL_EDIT_MARK).trim()] = k; });
    const hdr = aoa[0].map(h => lab[String(h).trim()] || lab[String(h).replace(/\s*✎\s*$/, "").trim()] || String(h).replace(/\s*✎\s*$/, "").trim());
    const keyF = (XL_COLS[c] || [[]])[0][0], ki = hdr.indexOf(keyF); if (ki < 0) return;
    out[c] = {};
    aoa.slice(1).forEach(row => { const key = String(row[ki] === undefined ? "" : row[ki]).trim(); if (!key) return; const o = {};
      XL_EDIT[c].f.forEach(f => { const i = hdr.indexOf(f); if (i > -1) o[f] = xlNorm(f, row[i]); }); out[c][key] = o; });
  });
  return out;
}
/** Bring changes typed in an area workbook into D. Returns {n, rejected[]} */
function xlAreaApply(area, E, S) {
  let n = 0; const rej = [];
  Object.keys(E).forEach(c => {
    const ed = XL_EDIT[c]; (S[c] = S[c] || {});
    Object.keys(E[c]).forEach(key => {
      const e = E[c][key], s = S[c][key], isObj = c === "budget", recs = isObj ? null : (D[c] || []), rec = isObj ? (D.budget || {})[key] : recs.find(r => xlKeyOf(c, r) === key);
      if (!rec) {
        if (s || !ed.add) return;                                   // deleted in the Hub, or adding not allowed
        if (!Object.values(e).some(v => v !== "")) return;
        const nr = {}; nr[(XL_COLS[c] || [[]])[0][0]] = key; ed.f.forEach(f => { if (e[f] !== undefined && e[f] !== "") nr[f] = e[f]; });
        if (c === "emps" && !nr.name) { rej.push(XL_NAMES[c] + " " + key + ": name missing"); return; }
        if (isObj) { D.budget = D.budget || {}; D.budget[key] = { CAPEX: +nr.CAPEX || 0, OPEX: +nr.OPEX || 0 }; } else { recs.push(nr); D[c] = recs; } n++; return;
      }
      ed.f.forEach(f => {
        if (!(f in e)) return; const ev = e[f], sv = s ? s[f] : undefined, hv = xlNorm(f, rec[f]);
        if (sv !== undefined && ev === sv) return;                  // not changed in Excel
        if (sv === undefined && !ed.add) return;                     // row not written by the Hub yet
        if (ev === hv) return;
        if (XL_NUM.has(f) && ev !== "" && typeof ev !== "number") { rej.push(XL_NAMES[c] + " " + key + " – " + f + ": not a number"); return; }
        if (c === "assets" && f === "status" && ev && typeof arStatuses === "function" && arStatuses().indexOf(ev) < 0) { rej.push("Asset " + key + " – status \"" + ev + "\" is not a Hub status"); return; }
        if (c === "dns" && f === "status" && typeof DNST !== "undefined" && DNST.indexOf(ev) < 0) { rej.push("Debit note " + key + " – status \"" + ev + "\" is not a Hub status"); return; }
        if (c === "ut" && f === "qty" && (ev === "" ? 0 : ev) < 0) { rej.push("Reading " + key + ": negative"); return; }
        rec[f] = ev === "" && XL_NUM.has(f) ? 0 : ev; n++;
      });
    });
  });
  return { n, rej };
}

/* ---------------- files, sync state, the weekly folder link ---------------- */
LOCAL.xlKnown = {}; LOCAL.xlDirty = {}; LOCAL.xlMonths = new Set(); LOCAL.xlTimer = null; LOCAL.xlWarned = false; LOCAL.weekly = undefined;   // undefined = not checked yet on this open
async function xlCopies(d, name) {
  const base = name.replace(/\.xlsx$/i, ""), out = [];
  for await (const [n, h] of d.entries()) if (h.kind === "file" && n !== name && n.indexOf(base) === 0 && /^(-[^.]+| \(\d+\))\.xlsx$/i.test(n.slice(base.length))) out.push(n);
  return out;
}
async function xlReadFile(d, name) { const fh = await d.getFileHandle(name).catch(() => null); if (!fh) return null; const f = await fh.getFile(); return { f, u8: new Uint8Array(await f.arrayBuffer()) }; }
const xlSame = (a, b) => { if (!a || a.length !== b.length) return false; for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false; return true; };
async function xlWriteFile(d, name, bytes, id) {
  const cur = await xlReadFile(d, name), nb = new Uint8Array(bytes);
  if (!cur || !xlSame(cur.u8, nb)) await writeFileTo(d, name, new Blob([nb], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  for (const n of await xlCopies(d, name)) await d.removeEntry(n).catch(() => {});
  LOCAL.xlKnown[id] = (await (await d.getFileHandle(name)).getFile()).lastModified;
}
async function xlSyncRead(area) { return (await readJson(await subDir(dataPath(area)), "_excel-sync.json").catch(() => null)) || { v: 2, files: {} }; }
async function xlSyncWrite(area, j) { j.v = 2; j.saved = new Date().toISOString(); await writeFileTo(await subDir(dataPath(area)), "_excel-sync.json", JSON.stringify(j)); }
async function xlTemplateBytes() {
  const t = await subDir(AREAS.ut + "/" + XL_TPL).catch(() => null); if (!t) return null;
  const fh = await t.getFileHandle(XL_TPL_NAME + ".xlsx").catch(() => null); return fh ? new Uint8Array(await (await fh.getFile()).arrayBuffer()) : null;
}
/* Turki's weekly folder (outside MIDU Hub Data): linked once per PC, kept in this browser */
LOCAL.weeklyInit = async function () {
  /* Personal OneDrive folders are no longer linked: drop any link saved on this PC (Excel stays in MIDU Hub Data only). */
  try { await idb("readwrite", s => s.delete("utWeekly")); } catch (e) {}
  LOCAL.weekly = null; return;
  let h = null; try { h = await idb("readonly", s => s.get("utWeekly")); } catch (e) {}
  if (!h) { LOCAL.weekly = null; return; }
  try { const p = await h.queryPermission({ mode: "readwrite" }); LOCAL.weekly = { h, ok: p === "granted" }; } catch (e) { LOCAL.weekly = { h, ok: false }; }
};
LOCAL.weeklyLink = async function () {
  throw new Error("Linking a personal OneDrive folder is turned off. The weekly workbook is in MIDU Hub Data › 02 Utilities › Excel (linked with Hub).");
  let h = LOCAL.weekly && LOCAL.weekly.h;
  if (h && !LOCAL.weekly.ok) { if ((await h.requestPermission({ mode: "readwrite" })) === "granted") { LOCAL.weekly.ok = true; return h.name; } }
  h = await window.showDirectoryPicker({ id: "midu-ut-weekly", mode: "readwrite" });
  if ((await h.requestPermission({ mode: "readwrite" })) !== "granted") throw new Error("Editing files in \"" + h.name + "\" was not allowed.");
  let found = false; for await (const [n, x] of h.entries()) { if (x.kind === "directory" && XL_MON.indexOf(n) > -1) found = true; if (x.kind === "file" && /^Weekly_Report_Utilities_Back_Charge_MIDU/i.test(n)) found = true; }
  if (!found) throw new Error("\"" + h.name + "\" does not look like the weekly folder. Pick Utilities › Weekly Reports For Sub-contractors (it has the month folders November, December …).");
  try { await idb("readwrite", s => s.put(h, "utWeekly")); } catch (e) {}
  LOCAL.weekly = { h, ok: true }; return h.name;
};
LOCAL.weeklyUnlink = async function () { try { await idb("readwrite", s => s.delete("utWeekly")); } catch (e) {} LOCAL.weekly = null; };
/** The workbook of a month in the weekly folder: existing file whose Setup month matches, else a new one from the template. */
const XL_MCACHE = {};
async function xlWeeklyTarget(month) {
  if (!LOCAL.weekly || !LOCAL.weekly.ok) return null;
  const root = LOCAL.weekly.h, mName = XL_MON[+month.slice(5, 7) - 1];
  const d = await root.getDirectoryHandle(mName, { create: true });
  for await (const [n, h] of d.entries()) {
    if (h.kind !== "file" || !/^Weekly_Report_Utilities_Back_Charge_MIDU.*\.xlsx$/i.test(n) || /\)(-[^)]+| \(\d+\))\.xlsx$/i.test(n)) continue;
    const f = await h.getFile(), ck = mName + "/" + n + "|" + f.lastModified;
    if (!(ck in XL_MCACHE)) { try { XL_MCACHE[ck] = xlUtilInfo(new Uint8Array(await f.arrayBuffer())).month; } catch (e) { XL_MCACHE[ck] = ""; } }
    if (XL_MCACHE[ck] === month) return { d, name: n, id: "weekly:" + mName + "/" + n };
  }
  return { d, name: xlUtilName(month), id: "weekly:" + mName + "/" + xlUtilName(month), isNew: true };
}
/** Ask what to do when a workbook is linked for the first time and it differs from the Hub. */
function xlAskFirst(name, diffs) {
  return new Promise(res => {
    const list = diffs.slice(0, 15).map(([k, e, h]) => `<tr><td class="en">${xlEsc(k.split("|")[1])}</td><td>${xlEsc(((D.uitems || []).find(i => i.key === k.split("|")[0]) || {}).title || k.split("|")[0])}</td><td class="num">${h || "–"}</td><td class="num">${e || "–"}</td></tr>`).join("");
    const o = lockBox(`<h3 style="margin:0 0 6px">${T("Link the Excel workbook", "ربط ملف Excel")}</h3><p style="margin:0 0 8px;color:#444">${xlEsc(name)}<br>${diffs.length} ${T("readings are different in the Hub and in this workbook. Which should be kept?", "قراءة مختلفة بين المنصة وهذا الملف. أيهما يُعتمد؟")}</p>
      <div style="max-height:240px;overflow:auto"><table class="t" style="font-size:12px"><tr><th>${T("Day", "اليوم")}</th><th>${T("Item", "الصنف")}</th><th>Hub</th><th>Excel</th></tr>${list}</table></div>${diffs.length > 15 ? `<p class="sub">+ ${diffs.length - 15} …</p>` : ""}
      <button id="xfE" style="${lockBtn};margin-top:12px">${T("Use the Excel values", "اعتماد قيم Excel")}</button>
      <button id="xfH" style="${lockBtn};background:#fff;color:#1B273D;border:1px solid #cfd6e3;margin-top:8px">${T("Keep the Hub values", "الإبقاء على قيم المنصة")}</button>`);
    o.querySelector("#xfE").onclick = () => { o.remove(); res("excel"); };
    o.querySelector("#xfH").onclick = () => { o.remove(); res("hub"); };
  });
}
/** Readings: apply Excel values to D.ut. changes = [[key, value]] */
function xlApplyReadings(changes) {
  let n = 0;
  changes.forEach(([key, e]) => { const u = D.ut.find(r => r.id === key); const [k, date] = key.split("|"), it = (D.uitems || []).find(i => i.key === k); if (!it) return;
    if (u) { if ((+u.qty || 0) !== e) { u.qty = e; n++; } } else if (e) { D.ut.push({ id: key, item: k, date, type: it.cat, c: it.loc || "", qty: e, ref: "", title: it.title, unit: it.unit || "", company: it.company || "" }); n++; } });
  return n;
}
/** One utilities workbook: import its Excel changes (and its OneDrive conflict copies). Returns changes count. */
async function xlUtilImport(t, sync) {
  const snap = sync.files[t.id], files = [t.name].concat(await xlCopies(t.d, t.name)), changes = [];
  for (const fn of files) {
    const x0 = await xlReadFile(t.d, fn); if (!x0) continue;
    if (fn === t.name && snap && x0.f.lastModified === LOCAL.xlKnown[t.id]) continue;
    let x; try { x = xlReadUtilities(x0.u8); } catch (e) { console.warn("[hub] cannot read", fn, e); continue; }
    if (!x.month) continue;
    if (!snap) {                                                 // first link of an existing workbook
      const hub = {}; (D.ut || []).filter(u => String(u.date).slice(0, 7) === x.month && +u.qty).forEach(u => hub[u.id] = +u.qty);
      const diffs = [...new Set(Object.keys(hub).concat(Object.keys(x.cells)))].filter(k => !x.skip.has(k) && (x.info.rows[k.split("|")[0]] || x.info.extra.some(i => i.key === k.split("|")[0]) ) && (+x.cells[k] || 0) !== (hub[k] || 0)).sort().map(k => [k, +x.cells[k] || 0, hub[k] || 0]);
      if (diffs.length && (await xlAskFirst(fn, diffs)) === "excel") diffs.forEach(([k, e]) => changes.push([k, e]));
      continue;
    }
    const S = snap.month === x.month ? snap.cells : {};
    new Set(Object.keys(x.cells).concat(Object.keys(S))).forEach(key => {
      if (x.skip.has(key) || key.split("|")[1].slice(0, 7) !== x.month) return;
      const e = +(x.cells[key] || 0), s0 = +(S[key] || 0); if (e === s0) return;
      changes.push([key, e]);
    });
  }
  return changes;
}
/** Two-way sync of the utilities workbooks of one month (Hub folder + Turki's folder). */
async function xlUtilMonth(month) {
  const tpl = await xlTemplateBytes(); if (!tpl) return [];
  const sync = await xlSyncRead("ut"), targets = [];
  const hd = await subDir(AREAS.ut + "/" + XL_DIR, true);
  targets.push({ d: hd, name: xlUtilName(month), id: "hub:" + xlUtilName(month) });
  const wt = await xlWeeklyTarget(month).catch(e => { console.warn("[hub] weekly folder", e); return null; }); if (wt) targets.push(wt);
  const changes = [];
  for (const t of targets) { if (t.isNew) continue; changes.push(...await xlUtilImport(t, sync)); }
  const n = xlApplyReadings(changes);
  if (n) { persist(); xlRerender(); toast("✔ " + n + " " + tr("utility readings brought in from Excel", "قراءة خدمات أُدخلت من Excel")); }
  const done = [];
  for (const t of targets) {
    const cur = await xlReadFile(t.d, t.name), base = cur ? cur.u8 : tpl;
    let r; try { r = xlFillUtilities(base, month, D.ut || []); } catch (e) { console.warn("[hub] not a utilities workbook:", t.name, e); continue; }
    try { await xlWriteFile(t.d, t.name, r.bytes, t.id); sync.files[t.id] = { month, cells: r.cells, at: new Date().toISOString(), by: LOCAL.user }; done.push(t.id); }
    catch (e) { xlWarn(t.name, e); }
  }
  await xlSyncWrite("ut", sync);
  return done;
}
/** Two-way sync of an area workbook. */
async function xlAreaSync(area) {
  const d = await subDir(AREAS[area] + "/" + XL_DIR, true), name = AREAS[area] + " - Hub data.xlsx", id = "area:" + area, sync = await xlSyncRead(area);
  const snap = sync.files[id], files = [name].concat(await xlCopies(d, name)); let n = 0; const rej = [];
  if (snap && Object.keys(XL_EDIT).some(c => COLL_AREA[c] === area)) for (const fn of files) {
    const x0 = await xlReadFile(d, fn); if (!x0) continue;
    if (fn === name && x0.f.lastModified === LOCAL.xlKnown[id]) continue;
    let E; try { E = xlAreaRead(x0.u8, area); } catch (e) { console.warn("[hub] cannot read", fn, e); continue; }
    const r = xlAreaApply(area, E, snap.cells || {}); n += r.n; rej.push(...r.rej);
  }
  if (n) { persist(); xlRerender(); toast("✔ " + n + " " + tr("change(s) brought in from Excel", "تغيير أُدخل من Excel") + " (" + AREAS[area] + ")"); }
  if (rej.length) toast("⚠ " + rej.length + " " + tr("Excel change(s) not accepted: ", "تغيير من Excel غير مقبول: ") + rej.slice(0, 3).join(" · "));
  try { await xlWriteFile(d, name, xlAreaBook(area, D), id); sync.files[id] = { cells: xlAreaSnap(area, D), at: new Date().toISOString(), by: LOCAL.user }; await xlSyncWrite(area, sync); return name; }
  catch (e) { xlWarn(name, e); return null; }
}
const XL_RERENDER = ["home", "dashWH", "dashPF", "dashHR", "dashMT", "utRep", "utReg", "search", "assetReg", "pr", "po", "fin"];
function xlRerender() { try { if (typeof view !== "undefined" && XL_RERENDER.indexOf(view) > -1 && !document.querySelector("#ov.on")) render(true); } catch (e) {} }
function xlWarn(name, e) {
  console.warn("[hub] Excel", name, e);
  if (!LOCAL.xlWarned) { LOCAL.xlWarned = true; toast("⚠ " + tr("Excel file not updated (is it open in Excel?): ", "لم يُحدَّث ملف Excel (هل هو مفتوح؟): ") + name + tr(". Nothing is lost – it is tried again every 30 seconds.", ". لا يضيع شيء – تُعاد المحاولة كل 30 ثانية.")); }
}
const xlMonthsNow = () => { const now = new Date(), m = [now.getFullYear() + "-" + String(now.getMonth() + 1).padStart(2, "0")]; if (now.getDate() <= 10) { const p = new Date(now.getFullYear(), now.getMonth() - 1, 1); m.push(p.getFullYear() + "-" + String(p.getMonth() + 1).padStart(2, "0")); } return m; };
/* Update 58 (25 Nov 2026): only the Hub administrator (Azzam) edits Excel. The Excel copies stay linked BOTH ways, but only on the administrator's PC; for everyone else the Hub never reads or writes Excel (their access to the Excel folders is view only). */
const xlCan = () => LOCAL.connected && !LOCAL.readOnly && LOCAL.isAdmin() && LOCAL.layout === 2 && typeof XLSX !== "undefined" && typeof D !== "undefined" && D && D.ut;
/** Write (after first bringing in Excel changes) the workbooks of the given areas. */
async function xlRefreshNow(areas, months) {
  if (!xlCan()) return [];
  areas = areas || Object.keys(AREAS).filter(a => LOCAL.areaOK[a]);
  const done = [];
  for (const a of areas) {
    if (!LOCAL.areaOK[a]) continue;
    try { const r = await xlAreaSync(a); if (r) done.push(r);
      if (a === "ut") { const ms = new Set((months && months.size ? [...months] : []).concat(xlMonthsNow())); for (const m of ms) done.push(...await xlUtilMonth(m)); }
    } catch (e) { xlWarn(AREAS[a], e); }
  }
  return done;
}
LOCAL.xlBusy = Promise.resolve();
const xlQueue = fn => { const p = LOCAL.xlBusy.then(fn); LOCAL.xlBusy = p.catch(e => console.warn(e)); return p; };
LOCAL.xlRefresh = (areas, months) => xlQueue(() => xlRefreshNow(areas, months));
function xlSchedule(areas, months) {
  areas.forEach(a => LOCAL.xlDirty[a] = 1); (months || []).forEach(m => LOCAL.xlMonths.add(m));
  clearTimeout(LOCAL.xlTimer);
  LOCAL.xlTimer = setTimeout(() => { const a = Object.keys(LOCAL.xlDirty), m = LOCAL.xlMonths; LOCAL.xlDirty = {}; LOCAL.xlMonths = new Set(); LOCAL.xlRefresh(a, m); }, 4000);
}
/** Every 30 s: has any linked workbook changed (typed in Excel, or synced from another PC)? */
LOCAL.xlCheck = () => xlQueue(async () => {
  if (!xlCan()) return;
  const stale = new Set(), months = new Set();
  for (const a of Object.keys(AREAS).filter(a => LOCAL.areaOK[a])) {
    const d = await subDir(AREAS[a] + "/" + XL_DIR).catch(() => null); if (!d) continue;
    const name = AREAS[a] + " - Hub data.xlsx", x = await d.getFileHandle(name).catch(() => null);
    if (x && ((await x.getFile()).lastModified !== LOCAL.xlKnown["area:" + a] || (await xlCopies(d, name)).length)) stale.add(a);
    if (a === "ut") for (const m of xlMonthsNow()) {
      const ts = [{ d, name: xlUtilName(m), id: "hub:" + xlUtilName(m) }]; const w = await xlWeeklyTarget(m).catch(() => null); if (w && !w.isNew) ts.push(w);
      for (const t of ts) { const fh = await t.d.getFileHandle(t.name).catch(() => null); if (fh && ((await fh.getFile()).lastModified !== LOCAL.xlKnown[t.id] || (await xlCopies(t.d, t.name)).length)) months.add(m); }
    }
  }
  for (const a of stale) await xlAreaSync(a).catch(e => xlWarn(AREAS[a], e));
  for (const m of months) await xlUtilMonth(m).catch(e => xlWarn(xlUtilName(m), e));
});
setInterval(() => LOCAL.xlCheck(), 30000);
if (typeof window !== "undefined") window.addEventListener("focus", () => LOCAL.xlCheck());
/* when the Hub opens: link state of the weekly folder, then make sure every workbook exists and is current */
LOCAL.xlDaily = async function () { await LOCAL.weeklyInit(); return LOCAL.xlRefresh(); };

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
