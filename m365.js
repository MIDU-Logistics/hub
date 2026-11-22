/* ==========================================================================
   MIDU Logistics Hub — Microsoft 365 data layer (m365.js)
   --------------------------------------------------------------------------
   Replaces the Google Apps Script layer (Code.gs). The hub screens still call
   GAS.withSuccessHandler(..).withFailureHandler(..).<fn>(..) — this file gives
   them the same object, backed by Microsoft Graph:
     getAll()                     → read every SharePoint list
     upsertMany(payload)          → create / update only the changed records
     savePdf(html, name, path)    → PDF into the document library, by month
     exportSheets(path, title, s) → Excel file (.xlsx) into the document library
     importAttendance()           → ZKTeco files from "05 HR/01 Attendance Inbox"
   Sign-in: MSAL.js, delegated Graph permissions User.Read + Sites.ReadWrite.All.
   The page holds no data and no secrets — only the IDs in CONFIG below.
   ========================================================================== */

/* ============================ CONFIG ============================ */
const M365_CONFIG = {
  clientId: "",                                   // Application (client) ID — from IT (app registration)
  tenantId: "464e5e59-26fe-4715-ac16-9efb6ccf733f", // Directory (tenant) ID
  site: "mblaec-my.sharepoint.com:/personal/m_elmasri_midu_com_sa", // SharePoint site that holds the lists
  docRoot: "MIDU Logistics (Claude)/Hub Documents", // folder in that site's document library for PDFs and Excel files
  scopes: ["User.Read", "Sites.ReadWrite.All"]
};

/* Asset Register location → hub warehouse (approved 21 Nov 2026). First match wins. */
const M365_WH_MAP = [
  ["MAIN", "Main Warehouse", l => /^Laydown/i.test(l)],
  ["ASP",  "ASP Site Store", l => /^ASP\b/i.test(l) || /^PL 1 Workshop/i.test(l)],
  ["KHU",  "Khulais",        l => /Khulais/i.test(l)],
  ["SITE", "Site / Office",  l => true]
];

/* Default back-charge rates per Utility Items category, used until the "Utility Rates" list exists. */
const M365_DEFAULT_RATES = [
  ["Generators diesel", "L", 1.85 / 1.15], ["Tower lights diesel", "L", 1.85 / 1.15], ["Equipment diesel", "L", 1.85 / 1.15],   // 1.85 incl. VAT
  ["Potable water", "m³", 15], ["Sewage", "Skips", 350], ["Drinking water (labour)", "L", 0.08],
  ["Ice cubes", "Bags", 5], ["Drinking water cartons", "Cartons", 12], ["Garbage skip", "Skips", 650],
  ["Concrete trucks", "Trips", 1540], ["Concrete vacuum", "Trips", 1000], ["Concrete waste skip", "Skips", 1540]
].map(([type, unit, rate]) => ({ type, unit, rate }));

/* Screens each role can open (Hub Users list: Title = e-mail, Role = one or more roles separated by commas). */
const M365_ROLES = {
  admin: "*", management: "*",
  logistics:   ["home","grn","mds","mrs","asset","count","labels","utDaily","utRep","utBC","utReg","pr","mcl","mreq","search"],
  storekeeper: ["home","grn","mds","mrs","search"],
  procurement: ["home","pr","cqs","po","search"],
  finance:     ["home","po","fin","capex","utReg","search"],
  hr:          ["home","att","search"],
  maintenance: ["home","mcl","mreq","search"]
};

/* ============================ LIST MAPPING ============================
   Each hub collection → SharePoint list. cols: [hubField, column, kind]
   kinds: t = text, n = number, b = yes/no, d = date (written yyyy-MM-ddT12:00:00Z), j = JSON text.
   The key column is the column used to find a record (Title unless stated).            */
const iso12 = d => d ? String(d).slice(0, 10) + "T12:00:00Z" : null;
const day10 = v => v ? String(v).slice(0, 10) : "";
const satOf = ds => { const d = new Date(ds + "T12:00:00Z"); d.setUTCDate(d.getUTCDate() - ((d.getUTCDay() + 1) % 7)); return d.toISOString().slice(0, 10); };

const M365_LISTS = {
  /* ---------- existing lists (reused, never duplicated) ---------- */
  assets: { list: "MIDU Asset Register", custom: "assets" },
  emps: { list: "MIDU Employees", keyCol: "EmployeeID", cols: [
    ["no", "EmployeeID", "t"], ["iqama", "IqamaNo", "t"], ["name", "Title", "t"], ["dept", "Department", "t"], ["deviceId", "DeviceID", "t"]] },
  assetTypes: { list: "MIDU Asset Types", readOnly: true, keyCol: "field_1", cols: [
    ["prefix", "field_1", "t"], ["name", "Title", "t"], ["code", "DefaultItemCode0", "t"],
    ["cls", "field_2", "t"], ["gl", "field_3", "t"], ["life", "field_4", "n"]] },
  uitems: { list: "Utility Items", readOnly: true, keyCol: "ItemKey", cols: [
    ["key", "ItemKey", "t"], ["title", "Title", "t"], ["cat", "Category", "t"], ["loc", "Location", "t"], ["unit", "Unit", "t"],
    ["company", "Company", "t"], ["sort", "SortOrder", "n"], ["chargeTo", "ChargeTo", "t"], ["active", "Active", "b"],
    ["concretePct", "ConcretePct", "p"], ["mccPct", "MCCPct", "p"], ["emcoPct", "EMCOPct", "p"]] },
  ut: { list: "Utility Readings", custom: "ut", keyCol: "DayKey" },

  /* ---------- new lists (created only after approval) ---------- */
  items: { list: "Hub Items", cols: [["code","Title","t"],["desc","Description","t"],["unit","Unit","t"],["kind","Kind","t"],["min","MinQty","n"],["price","UnitPrice","n"]] },
  stock: { list: "Hub Stock", cols: [["code","ItemCode","t"],["wh","Warehouse","t"],["loc","StorageLocation","t"],["qty","Qty","n"]] },
  docs: { list: "Hub Documents", cols: [["no","Title","t"],["type","DocType","t"],["date","DocDate","d"],["party","Party","t"],["iqama","Iqama","t"],["dtype","DispatchType","t"],
    ["po","PONo","t"],["dn","DeliveryNote","t"],["wh","Warehouse","t"],["loc","StorageLocation","t"],["to","ToLocation","t"],["returnBy","ReturnBy","d"],["by","ReceivedBy","t"],
    ["remarks","Remarks","t"],["lines","Lines","j"],["asset","AssetSnapshot","j"]] },
  prs: { list: "Purchase Requests", cols: [["no","Title","t"],["date","PRDate","d"],["cc","CostCenter","t"],["req","RequestedBy","t"],["need","NeedDate","d"],["just","Justification","t"],["pd","PDApproved","b"],["lines","Lines","j"]] },
  cqs: { list: "CQS", cols: [["no","Title","t"],["pr","PRNo","t"],["date","CQSDate","d"],["sel","SelectedSupplier","t"],["approved","Approved","b"],["quotes","Quotes","j"]] },
  pos: { list: "PO & Petty Cash", cols: [["no","Title","t"],["kind","Kind","t"],["date","PODate","d"],["pr","PRNo","t"],["cqs","CQSNo","t"],["sup","Supplier","t"],["cc","CostCenter","t"],["fin","FinanceStatus","t"],["lines","Lines","j"]] },
  invs: { list: "Invoices", cols: [["po","PONo","t"],["no","InvoiceNo","t"],["date","InvoiceDate","d"],["amt","Amount","n"],["paid","PaidDate","d"]] },
  budget: { list: "Budget CAPEX-OPEX", cols: [["cc","Title","t"],["CAPEX","CAPEX","n"],["OPEX","OPEX","n"]] },
  rates: { list: "Utility Rates", cols: [["type","Title","t"],["unit","Unit","t"],["rate","Rate","n"]] },
  dns: { list: "Debit Notes", cols: [["no","Title","t"],["contractor","Contractor","t"],["period","Period","t"],["date","InvoiceDate","d"],["sub","Subtotal","n"],
    ["status","Status","t"],["paid","PaidAmount","n"],["auto","AutoCreated","b"],["lines","Lines","j"],["vat","VAT","p"]] },
  punches: { list: "Attendance Punches", cols: [["emp","EmployeeNo","t"],["ts","PunchTime","t"],["dev","Device","t"],["src","SourceFile","t"]] },
  mcl: { list: "Monthly Checklists", cols: [["no","Title","t"],["month","Month","t"],["loc","Location","t"],["by","CheckedBy","t"],["date","CheckDate","d"],["res","Results","j"]] },
  areq: { list: "Attendance Requests", cols: [["no","Title","t"],["emp","EmployeeNo","t"],["type","ReqType","t"],["date","ReqDate","d"],["from","FromDate","d"],["to","ToDate","d"],
    ["leaveType","LeaveType","t"],["inTime","InTime","t"],["outTime","OutTime","t"],["reason","Reason","t"],["status","Status","t"],["requestedBy","RequestedBy","t"],
    ["decidedBy","DecidedBy","t"],["decidedOn","DecidedOn","d"],["file","Attachment","t"],["fileName","FileName","t"]] },
  mreq: { list: "Maintenance Requests", cols: [["no","Title","t"],["date","RequestDate","d"],["by","RequestedBy","t"],["asset","AssetID","t"],["loc","Location","t"],["cat","Category","t"],
    ["prio","Priority","t"],["desc","Description","t"],["status","Status","t"],["to","AssignedTo","t"]] }
};
const M365_USERS_LIST = "Hub Users";

/* ============================ STATE ============================ */
const M365 = {
  msal: null, account: null, signedIn: false, user: null, roles: null,
  siteId: "", driveId: "", lists: {}, cols: {}, choices: {}, ids: {}, missing: [],
  run: null
};

/* ============================ SIGN-IN ============================ */
M365.configured = () => !!M365_CONFIG.clientId;
M365.init = async function () {
  if (!M365.configured() || typeof msal === "undefined") return false;
  M365.msal = new msal.PublicClientApplication({
    auth: { clientId: M365_CONFIG.clientId, authority: "https://login.microsoftonline.com/" + M365_CONFIG.tenantId, redirectUri: location.origin + location.pathname },
    cache: { cacheLocation: "localStorage" }
  });
  const res = await M365.msal.handleRedirectPromise().catch(e => { console.error(e); return null; });
  M365.account = (res && res.account) || M365.msal.getAllAccounts()[0] || null;
  if (!M365.account) return false;
  M365.msal.setActiveAccount(M365.account);
  try {
    const me = await M365.graph("GET", "/me?$select=displayName,mail,userPrincipalName");
    M365.user = { name: me.displayName, email: String(me.mail || me.userPrincipalName || "").toLowerCase() };
    await M365.connect();
    M365.signedIn = true;
  } catch (e) {
    console.error(e); M365.error = e.message; return false;
  }
  return true;
};
M365.signIn = () => M365.msal && M365.msal.loginRedirect({ scopes: M365_CONFIG.scopes });
M365.signOut = () => M365.msal && M365.msal.logoutRedirect({ account: M365.account });

M365.token = async function () {
  const req = { scopes: M365_CONFIG.scopes, account: M365.account };
  try { return (await M365.msal.acquireTokenSilent(req)).accessToken; }
  catch (e) {
    if (e instanceof msal.InteractionRequiredAuthError) { await M365.msal.acquireTokenRedirect(req); }
    throw e;
  }
};

/* ============================ GRAPH ============================ */
M365.graph = async function (method, url, body, opt = {}) {
  const full = /^https:/.test(url) ? url : "https://graph.microsoft.com/v1.0" + url;
  for (let attempt = 0; attempt < 5; attempt++) {
    const headers = { Authorization: "Bearer " + await M365.token(), Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" };
    let payload = body;
    if (body !== undefined && !(body instanceof Blob) && !(body instanceof ArrayBuffer)) { headers["Content-Type"] = "application/json"; payload = JSON.stringify(body); }
    if (opt.type) headers["Content-Type"] = opt.type;
    const r = await fetch(full, { method, headers, body: payload });
    if (r.status === 429 || r.status === 503 || r.status === 504) {
      const wait = (+r.headers.get("Retry-After") || 2 ** attempt) * 1000;
      await new Promise(ok => setTimeout(ok, wait)); continue;
    }
    if (r.status === 204) return null;
    if (opt.raw) { if (!r.ok) throw new Error("Graph " + r.status); return r; }
    const j = await r.json().catch(() => ({}));
    if (!r.ok) { const err = new Error((j.error && j.error.message) || ("Graph error " + r.status)); err.status = r.status; throw err; }
    return j;
  }
  throw new Error("Microsoft Graph is busy — try again in a minute.");
};
M365.all = async function (url) {
  let out = [], next = url;
  while (next) { const j = await M365.graph("GET", next); out = out.concat(j.value || []); next = j["@odata.nextLink"]; }
  return out;
};
const pool = async (arr, n, fn) => { const res = []; let i = 0; await Promise.all(Array.from({ length: Math.min(n, arr.length) }, async () => { while (i < arr.length) { const k = i++; res[k] = await fn(arr[k], k); } })); return res; };
const odq = s => String(s).replace(/'/g, "''");

/* Site, lists, columns and choice values (loaded once after sign-in). */
M365.connect = async function () {
  const site = await M365.graph("GET", "/sites/" + M365_CONFIG.site + "?$select=id");
  M365.siteId = site.id;
  const lists = await M365.all("/sites/" + M365.siteId + "/lists?$select=id,displayName&$top=200");
  lists.forEach(l => M365.lists[l.displayName] = l.id);
  M365.missing = Object.values(M365_LISTS).map(c => c.list).filter((n, i, a) => a.indexOf(n) === i && !M365.lists[n]);
  await pool(Object.keys(M365_LISTS).filter(k => M365.lists[M365_LISTS[k].list]), 4, async k => {
    const name = M365_LISTS[k].list, cols = await M365.all("/sites/" + M365.siteId + "/lists/" + M365.lists[name] + "/columns?$top=500");
    M365.cols[name] = {};
    cols.forEach(c => { M365.cols[name][c.name] = c; if (c.choice) M365.choices[name + "|" + c.name] = c.choice.choices || []; });
  });
  /* roles from "Hub Users" (no list or empty list = set-up mode, every signed-in user has all screens) */
  M365.roles = "*";
  if (M365.lists[M365_USERS_LIST]) {
    const users = await M365.all("/sites/" + M365.siteId + "/lists/" + M365.lists[M365_USERS_LIST] + "/items?$expand=fields($select=Title,Role,Sites)&$top=999");
    if (users.length) {
      const me = users.find(u => String(u.fields.Title || "").trim().toLowerCase() === M365.user.email);
      M365.roles = me ? String(me.fields.Role || "").toLowerCase().split(/[,;]/).map(s => s.trim()).filter(Boolean) : [];
    }
  }
};
M365.can = view => {
  if (!M365.signedIn && typeof LOCAL !== "undefined" && LOCAL.connected && LOCAL.roles) return LOCAL.roles.some(r => M365_ROLES[r] === "*" || (M365_ROLES[r] || []).indexOf(view) > -1);
  if (!M365.signedIn || M365.roles === "*") return true;
  return M365.roles.some(r => M365_ROLES[r] === "*" || (M365_ROLES[r] || []).indexOf(view) > -1);
};
M365.choice = (list, col) => M365.choices[list + "|" + col] || [];
M365.hasCol = (list, col) => !!(M365.cols[list] && M365.cols[list][col]);

/* ============================ RECORD MAPPING ============================ */
function fromCol(kind, v) {
  if (kind === "n") return v === null || v === undefined || v === "" ? 0 : +v;
  if (kind === "p") return v === null || v === undefined || v === "" ? null : +v;   // share: empty = not set
  if (kind === "b") return v === true || String(v).toLowerCase() === "true";
  if (kind === "d") return day10(v);
  if (kind === "j") { if (!v) return []; try { return JSON.parse(v); } catch (e) { return []; } }
  return v === null || v === undefined ? "" : String(v);
}
function toCol(kind, v) {
  if (kind === "n" || kind === "p") return v === "" || v === null || v === undefined ? null : +v;
  if (kind === "b") return v === true || v === "true";
  if (kind === "d") return v ? iso12(v) : null;
  if (kind === "j") return JSON.stringify(v === undefined ? null : v);
  return v === null || v === undefined ? "" : String(v);
}
const same = (a, b) => JSON.stringify(a === undefined ? null : a) === JSON.stringify(b === undefined ? null : b);

/* generic lists */
function genRead(name, f) { const o = {}; M365_LISTS[name].cols.forEach(([h, c, k]) => o[h] = fromCol(k, f[c])); return o; }
function genWrite(name, rec, old) {
  const cfg = M365_LISTS[name], out = {};
  cfg.cols.forEach(([h, c, k]) => {
    if (!M365.hasCol(cfg.list, c)) return;                 // column not in the list (e.g. DeviceID before it is added)
    if (old && same(rec[h], old[h])) return;               // PATCH only what changed
    out[c] = toCol(k, rec[h]);
  });
  if (!old && !("Title" in out) && M365.hasCol(cfg.list, "Title") && !cfg.cols.some(x => x[1] === "Title")) out.Title = KEYS[name](rec);
  if (name === "dns" && (!old || !same(rec.sub, old.sub))) {    // keep VAT / Total columns in step when the list has them
    const vat = rec.vat !== undefined && rec.vat !== null && rec.vat !== "" ? +rec.vat : Math.round(rec.sub * 15) / 100;
    if (M365.hasCol(cfg.list, "VAT")) out.VAT = vat;
    if (M365.hasCol(cfg.list, "Total")) out.Total = Math.round((rec.sub + vat) * 100) / 100;
  }
  if (name === "emps" && out.Department !== undefined && M365.choice(cfg.list, "Department").length && M365.choice(cfg.list, "Department").indexOf(out.Department) < 0) delete out.Department;
  return out;
}
const keyColOf = name => M365_LISTS[name].keyCol || "Title";

/* Asset Register */
const whOfLoc = loc => (M365_WH_MAP.find(m => m[2](loc || "")) || M365_WH_MAP[M365_WH_MAP.length - 1])[0];
function assetRead(f) {
  const empNo = String(f.EmployeeID || "").trim();
  const e = empNo ? (D.emps || []).find(x => x.no === empNo || x.iqama === empNo) : null;
  return {
    id: f.Title || "", type: f.field_1 || "", desc: f.field_4 || "", brand: f.field_5 || "", model: f.field_6 || "", serial: f.field_7 || "",
    itemCode: String(f.WarehouseItemCode || "").split(" | ")[0].trim(), own: f.Ownership || "",
    cust: e ? e.iqama || e.no : empNo, custName: f.field_10 || "", loc: f.field_9 || "", wh: whOfLoc(f.field_9), status: f.field_11 || ""
  };
}
function assetWrite(rec, old) {
  const L = "MIDU Asset Register", out = {}, warn = [];
  const ch = (h) => !old || !same(rec[h], old[h]);
  const pick = (col, v) => { const c = M365.choice(L, col); if (!v) return null; if (!c.length || c.indexOf(v) > -1) return v; warn.push(col + ": " + v); return undefined; };
  if (!old) {
    const t = (D.assetTypes || []).find(x => x.name === rec.type) || {};
    out.Title = rec.id; out.field_1 = rec.type; out.field_15 = 1; out.field_16 = "Pcs";
    if (M365.choice(L, "Company").indexOf("MIDU") > -1) out.Company = "MIDU";
    if (rec.own !== "Rented") { if (t.cls) out.field_2 = t.cls; if (t.gl) out.field_3 = t.gl; if (t.life) out.field_12 = t.life; }
  }
  if (ch("desc")) out.field_4 = rec.desc;
  if (ch("brand")) out.field_5 = rec.brand;
  if (ch("model")) out.field_6 = rec.model;
  if (ch("serial")) out.field_7 = String(rec.serial || "").toUpperCase();
  if (ch("own")) { const v = pick("Ownership", rec.own); if (v !== undefined) out.Ownership = v; }
  if (ch("itemCode")) {
    const full = M365.choice(L, "WarehouseItemCode").find(c => c.split(" | ")[0].trim() === rec.itemCode);
    if (full) out.WarehouseItemCode = full; else if (rec.itemCode) warn.push("Warehouse Item Code: " + rec.itemCode);
  }
  if (ch("loc")) { const v = pick("field_9", rec.loc); if (v !== undefined) out.field_9 = v; }
  if (ch("cust")) {
    const e = rec.cust ? emp(rec.cust) : null;
    out.EmployeeID = e ? e.no : (rec.cust || "");            // Employee ID = employee number (decision 21 Nov)
    out.field_10 = e ? e.name : "";
    if (e && e.dept && M365.choice(L, "Department").indexOf(e.dept) > -1) out.Department = e.dept;
    if (rec.cust) out.HandoverDate = iso12(today());
  }
  if (ch("status") || ("EmployeeID" in out)) { const v = pick("field_11", rec.status); if (v) out.field_11 = v; }
  return { out, warn };
}

/* Utility Readings — one row per item per day (DayKey = ItemKey|yyyy-MM-dd) */
function utRead(f) {
  return { id: f.DayKey || ((f.ItemKey || "") + "|" + day10(f.ReadingDate)), item: f.ItemKey || "", date: day10(f.ReadingDate), type: f.Category || "", c: f.Location || "",
    qty: +f.Quantity || 0, ref: f.Notes || "", title: f.Title || "", unit: f.Unit || "", company: f.Company || "" };
}
function utWrite(rec, old) {
  const out = {};
  if (!old) {
    const it = (D.uitems || []).find(x => x.key === rec.item) || {};
    Object.assign(out, { Title: it.title || rec.title || rec.item, ItemKey: rec.item, Category: it.cat || rec.type, Location: it.loc || rec.c, Unit: it.unit || rec.unit,
      Company: it.company || rec.company || "", ReadingDate: iso12(rec.date), MonthKey: rec.date.slice(0, 7), WeekStart: iso12(satOf(rec.date)), DayKey: rec.id });
  }
  if (!old || !same(rec.qty, old.qty)) out.Quantity = +rec.qty;
  if (!old || !same(rec.ref || "", old.ref || "")) out.Notes = rec.ref || "";
  return out;
}

/* ============================ API USED BY THE HUB ============================ */
async function listItems(name) {
  const cfg = M365_LISTS[name], lid = M365.lists[cfg.list];
  if (!lid) return [];
  const rows = await M365.all("/sites/" + M365.siteId + "/lists/" + lid + "/items?$expand=fields&$top=999");
  M365.ids[name] = {};
  const out = [];
  rows.forEach(r => {
    const f = r.fields || {};
    const rec = cfg.custom === "assets" ? assetRead(f) : cfg.custom === "ut" ? utRead(f) : genRead(name, f);
    if (name === "uitems" && rec.active === false && f.Active !== undefined) return;
    const key = KEYS[name] ? KEYS[name](rec) : rec.key;
    if (!key || key === "|" || key === "||") return;
    M365.ids[name][key] = r.id;
    out.push(rec);
  });
  return out;
}

async function getAll() {
  const out = {};
  out.emps = await listItems("emps");               // first: assets and attendance resolve employees
  if (typeof D !== "undefined") D.emps = out.emps;
  const names = Object.keys(M365_LISTS).filter(k => k !== "emps");
  const got = await pool(names, 4, n => listItems(n));
  names.forEach((n, i) => out[n] = got[i]);
  out._missing = M365.missing.slice();
  return out;
}

async function findId(name, key) {
  if (M365.ids[name] && M365.ids[name][key]) return M365.ids[name][key];
  const cfg = M365_LISTS[name], col = keyColOf(name);
  const r = await M365.graph("GET", "/sites/" + M365.siteId + "/lists/" + M365.lists[cfg.list] + "/items?$select=id&$filter=fields/" + col + " eq '" + odq(key) + "'&$top=1");
  return r.value && r.value[0] ? r.value[0].id : null;
}

async function upsertMany(payload) {
  const result = {}, problems = [], failed = [];
  for (const name of Object.keys(payload || {})) {
    const cfg = M365_LISTS[name];
    if (!cfg) { failed.push("Unknown collection: " + name); continue; }
    if (cfg.readOnly) continue;
    const lid = M365.lists[cfg.list];
    if (!lid) { failed.push("The SharePoint list \"" + cfg.list + "\" does not exist yet — it is created in step 4."); continue; }
    const base = "/sites/" + M365.siteId + "/lists/" + lid + "/items";
    M365.ids[name] = M365.ids[name] || {};
    SNAP[name] = SNAP[name] || {};
    let ok = 0;
    await pool(payload[name], 4, async rec => {
      const key = KEYS[name](rec), json = JSON.stringify(rec);
      const old = SNAP[name][key] ? JSON.parse(SNAP[name][key]) : null;
      let fields, warn = [];
      if (cfg.custom === "assets") { const r = assetWrite(rec, old); fields = r.out; warn = r.warn; }
      else if (cfg.custom === "ut") fields = utWrite(rec, old);
      else fields = genWrite(name, rec, old);
      if (warn.length) problems.push(key + " — not a valid choice: " + warn.join(", "));
      try {
        if (Object.keys(fields).length) {
          let id = M365.ids[name][key] || await findId(name, key);
          if (id) await M365.graph("PATCH", base + "/" + id + "/fields", fields);
          else { const created = await M365.graph("POST", base, { fields: fields }); M365.ids[name][key] = created.id; }
        }
        SNAP[name][key] = json; ok++;                       // written: never sent again unless it changes
      } catch (e) { failed.push(cfg.list + " / " + key + ": " + e.message); }
    });
    result[name] = ok;
  }
  if (problems.length) toast("⚠ " + problems.slice(0, 2).join(" · "));
  if (failed.length) throw new Error(failed.slice(0, 2).join(" · ") + (failed.length > 2 ? " (+" + (failed.length - 2) + ")" : ""));
  return result;
}

/* ---------- files in the document library ---------- */
const encPath = p => String(p).split("/").filter(Boolean).map(encodeURIComponent).join("/");
async function uploadFile(path, name, blob) {
  const full = encPath(M365_CONFIG.docRoot + "/" + path + "/" + name);
  const root = "/sites/" + M365.siteId + "/drive/root:/" + full;
  if (blob.size < 3.5 * 1024 * 1024) return M365.graph("PUT", root + ":/content", blob, { type: blob.type || "application/octet-stream" });
  const s = await M365.graph("POST", root + ":/createUploadSession", { item: { "@microsoft.graph.conflictBehavior": "replace" } });
  const chunk = 5 * 320 * 1024; let last = null;
  for (let start = 0; start < blob.size; start += chunk) {
    const end = Math.min(start + chunk, blob.size);
    const r = await fetch(s.uploadUrl, { method: "PUT", headers: { "Content-Range": `bytes ${start}-${end - 1}/${blob.size}` }, body: blob.slice(start, end) });
    if (!r.ok && r.status !== 202) throw new Error("Upload failed (" + r.status + ")");
    last = await r.json().catch(() => null);
  }
  return last;
}

async function savePdf(html, name, path) {
  /* the document page (with its own .doc styles) is rendered to an A4 PDF in the browser, then uploaded */
  const body = /<body>([\s\S]*)<\/body>/.exec(html), css = /<style>([\s\S]*?)<\/style>/.exec(html);
  const src = `<div style="width:190mm;background:#fff"><style>${css ? css[1].replace(/@page\{[^}]*\}/g, "") : ""}</style>${body ? body[1] : html}</div>`;
  const blob = await html2pdf().set({ margin: 10, filename: name + ".pdf", image: { type: "jpeg", quality: 0.95 },
    html2canvas: { scale: 2, backgroundColor: "#ffffff", scrollY: 0 }, jsPDF: { unit: "mm", format: "a4", orientation: "portrait" },
    pagebreak: { mode: ["css", "legacy"] } }).from(src, "string").outputPdf("blob");
  const f = await uploadFile(path, name + ".pdf", new Blob([blob], { type: "application/pdf" }));
  return { url: f.webUrl, path: path + "/" + name + ".pdf" };
}

async function exportSheets(path, title, sheets) {
  const wb = XLSX.utils.book_new();
  sheets.forEach(s => {
    const ws = XLSX.utils.aoa_to_sheet(s.rows || []);
    const width = Math.max(1, ...(s.rows || []).map(r => r.length));
    ws["!cols"] = Array.from({ length: width }, (_, c) => ({ wch: Math.min(48, Math.max(10, ...(s.rows || []).map(r => String(r[c] === undefined ? "" : r[c]).length + 2))) }));
    for (const ref in ws) if (ref[0] !== "!" && typeof ws[ref].v === "number") ws[ref].z = "#,##0.00";
    XLSX.utils.book_append_sheet(wb, ws, String(s.name).replace(/[\\\/\?\*\[\]:]/g, " ").substring(0, 31));
  });
  const buf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const f = await uploadFile(path, title + ".xlsx", new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  return { url: f.webUrl, name: title + ".xlsx" };
}

/* ---------- HR — ZKTeco attendance (same rules as Code.gs importAttendance) ---------- */
function toTs(a, b) {
  const p = n => ("0" + n).slice(-2);
  const fmtD = d => d.getFullYear() + "-" + p(d.getMonth() + 1) + "-" + p(d.getDate()), fmtT = d => p(d.getHours()) + ":" + p(d.getMinutes());
  if (a instanceof Date && b === undefined) return fmtD(a) + "T" + fmtT(a);
  let d = a instanceof Date ? fmtD(a) : String(a === undefined ? "" : a).trim();
  let t = b === undefined ? "" : (b instanceof Date ? fmtT(b) : String(b).trim());
  if (!t) { const m = d.match(/^(\S+)[ T](\d{1,2}:\d{2})/); if (!m) return ""; d = m[1]; t = m[2]; }
  let m = d.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})$/);
  if (!m) { const x = d.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})$/); if (x) m = [null, x[3], x[2], x[1]]; }
  const tm = t.match(/^(\d{1,2}):(\d{2})/);
  if (!m || !tm) return "";
  return m[1] + "-" + p(m[2]) + "-" + p(m[3]) + "T" + p(tm[1]) + ":" + tm[2];
}
/** Parses one terminal export (ArrayBuffer + file name) into [{id, ts, dev}] */
function readPunchFile(buf, fileName) {
  const name = String(fileName).toLowerCase();
  let table;
  if (/\.(xlsx|xls)$/.test(name)) {
    const wb = XLSX.read(buf, { type: "array", cellDates: true });
    table = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: "" });
  } else {
    const text = new TextDecoder("utf-8").decode(buf).replace(/^﻿/, "");
    const delim = text.indexOf("\t") > -1 ? "\t" : (text.indexOf(";") > -1 && text.indexOf(",") < 0 ? ";" : ",");
    table = text.split(/\r?\n/).map(l => l.split(delim).map(s => s.trim().replace(/^"|"$/g, "")));
  }
  table = table.filter(r => r.some(c => String(c).trim() !== ""));
  if (!table.length) return [];
  const head = table[0].map(h => String(h).toLowerCase().trim());
  const hasHeader = head.some(h => /(id|no\.?|emp|person|user|name|time|date)/.test(h)) && !/^\d+$/.test(String(table[0][0]).trim());
  let iId = 0, iDT = 1, iD = -1, iT = -1, iDev = -1, start = 0;
  if (hasHeader) {
    start = 1;
    const find = re => head.findIndex(h => re.test(h));
    iId = find(/^(emp|employee|person|user|ac-?no|no\.?$|id$|.*\bid\b)/);
    iDT = find(/(date\s*time|datetime|punch\s*time|check\s*time|att.*time)/);
    iD = find(/^date$/); iT = find(/^time$/);
    iDev = find(/(device|terminal|alias|sn|machine)/);
    if (iId < 0) iId = 0;
  }
  const out = [];
  for (let k = start; k < table.length; k++) {
    const r = table[k];
    const id = String(r[iId] === undefined ? "" : r[iId]).trim();
    if (!id) continue;
    let ts = "";
    if (iDT >= 0) ts = toTs(r[iDT]); else if (iD >= 0 && iT >= 0) ts = toTs(r[iD], r[iT]);
    if (!ts) continue;
    out.push({ id, ts, dev: iDev >= 0 ? String(r[iDev]).trim() : (hasHeader ? "" : "Terminal") });
  }
  return out;
}
/** Matching (first hit wins): deviceId → employee No. → Iqama. Adds new punches to D.punches. */
function matchPunches(rows, src) {
  const byDevice = {}, byNo = {}, byIqama = {};
  D.emps.forEach(e => {
    if (e.deviceId) byDevice[String(e.deviceId).replace(/^0+/, "")] = e.no;
    byNo[String(e.no).toUpperCase()] = e.no;
    byNo[String(e.no).replace(/^E-/i, "").replace(/^0+/, "")] = e.no;
    if (e.iqama) byIqama[String(e.iqama)] = e.no;
  });
  const match = id => { const s = String(id).trim(), n = s.replace(/^0+/, ""); return byDevice[n] || byNo[s.toUpperCase()] || byNo[n] || byIqama[s] || ""; };
  const existing = {}; D.punches.forEach(p => existing[p.emp + "|" + p.ts] = true);
  const res = { imported: 0, duplicates: 0, unmatched: 0, unmatchedRows: [] };
  rows.forEach(r => {
    const e = match(r.id);
    if (!e) { res.unmatched++; res.unmatchedRows.push([r.id, r.ts, r.dev, src]); return; }
    const k = e + "|" + r.ts;
    if (existing[k]) { res.duplicates++; return; }
    existing[k] = true; D.punches.push({ emp: e, ts: r.ts, dev: r.dev || "", src: src }); res.imported++;
  });
  return res;
}
async function folderItem(path) {
  const full = encPath(M365_CONFIG.docRoot + "/" + path);
  try { return await M365.graph("GET", "/sites/" + M365.siteId + "/drive/root:/" + full); }
  catch (e) {
    if (e.status !== 404) throw e;
    const parts = path.split("/"), leaf = parts.pop(), parent = await folderItem(parts.join("/"));
    return M365.graph("POST", "/sites/" + M365.siteId + "/drive/items/" + parent.id + "/children", { name: leaf, folder: {}, "@microsoft.graph.conflictBehavior": "fail" });
  }
}
const stamp = () => new Date().toISOString().slice(0, 16).replace(/[-:T]/g, "");
async function saveUnmatched(rows) {
  if (!rows.length) return;
  const csv = "Device user ID,Punch time,Device,File\r\n" + rows.map(r => r.map(c => '"' + String(c).replace(/"/g, '""') + '"').join(",")).join("\r\n");
  await uploadFile("05 HR/03 Attendance Errors", "Attendance Unmatched " + stamp() + ".csv", new Blob([csv], { type: "text/csv" }));
}
async function importAttendance() {
  const inbox = await folderItem("05 HR/01 Attendance Inbox"), done = await folderItem("05 HR/02 Attendance Processed"), errs = await folderItem("05 HR/03 Attendance Errors");
  const files = (await M365.all("/sites/" + M365.siteId + "/drive/items/" + inbox.id + "/children?$top=200")).filter(f => f.file);
  const res = { files: 0, imported: 0, duplicates: 0, unmatched: 0, errors: 0 }, unmatched = [], before = D.punches.length;
  for (const f of files) {
    res.files++;
    try {
      const r = await M365.graph("GET", "/sites/" + M365.siteId + "/drive/items/" + f.id + "/content", undefined, { raw: true });
      const m = matchPunches(readPunchFile(await r.arrayBuffer(), f.name), f.name);
      res.imported += m.imported; res.duplicates += m.duplicates; res.unmatched += m.unmatched; unmatched.push(...m.unmatchedRows);
      await M365.graph("PATCH", "/sites/" + M365.siteId + "/drive/items/" + f.id, { parentReference: { id: done.id }, name: stamp() + " " + f.name });
    } catch (e) {
      res.errors++; console.error(e);
      await M365.graph("PATCH", "/sites/" + M365.siteId + "/drive/items/" + f.id, { parentReference: { id: errs.id }, name: stamp() + " " + f.name }).catch(() => {});
    }
  }
  await saveUnmatched(unmatched);
  const added = D.punches.slice(before);
  if (added.length) {
    await upsertMany({ punches: added });
    SNAP.punches = SNAP.punches || {};
    added.forEach(r => SNAP.punches[KEYS.punches(r)] = JSON.stringify(r));
  }
  return res;
}
/** Manual upload from the hub (works offline too; signed in it also keeps a copy of the file). */
async function importPunchFile(file) {
  const buf = await file.arrayBuffer();
  const m = matchPunches(readPunchFile(buf, file.name), file.name);
  if (M365.signedIn) {
    await uploadFile("05 HR/02 Attendance Processed", stamp() + " " + file.name, new Blob([buf]));
    await saveUnmatched(m.unmatchedRows);
  }
  else if (typeof LOCAL !== "undefined" && LOCAL.connected) await localSaveBlob("05 HR/02 Attendance Processed", stamp() + " " + file.name, new Blob([buf]));
  return m;
}

/* ============================ google.script.run-style wrapper ============================ */
function makeRunner() {
  const fns = { getAll, upsertMany, savePdf, exportSheets, importAttendance };
  const chain = (ok, ko) => {
    const o = { withSuccessHandler: f => chain(f, ko), withFailureHandler: f => chain(ok, f) };
    Object.keys(fns).forEach(k => o[k] = (...args) => { fns[k](...args).then(r => ok && ok(r)).catch(e => { console.error(e); ko && ko(e); }); });
    return o;
  };
  return chain(null, null);
}
M365.run = makeRunner();
M365.readPunchFile = readPunchFile;
M365.importPunchFile = importPunchFile;
M365.whOfLoc = whOfLoc;
M365.uploadDoc = async (path, name, blob) => (await uploadFile(path, name, blob)).webUrl;
