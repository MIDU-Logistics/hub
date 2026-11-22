/* MIDU Logistics Hub — service worker: caches the app files only (never data or sign-in calls). */
const CACHE = "midu-hub-v5";
const FILES = ["./", "index.html", "m365.js", "local.js", "manifest.webmanifest", "icon-192.png", "icon-512.png", "apple-touch-icon.png", "favicon-32.png",
  "vendor/msal-browser.min.js", "vendor/xlsx.full.min.js", "vendor/html2pdf.bundle.min.js", "vendor/jsbarcode.min.js", "vendor/jspdf.umd.min.js", "vendor/qrcode.js", "vendor/midu-label-logo.js"];
self.addEventListener("install", e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)).then(() => self.skipWaiting())); });
self.addEventListener("activate", e => { e.waitUntil(caches.keys().then(k => Promise.all(k.filter(n => n !== CACHE).map(n => caches.delete(n)))).then(() => self.clients.claim())); });
self.addEventListener("fetch", e => {
  const url = new URL(e.request.url);
  if (e.request.method !== "GET" || url.origin !== location.origin) return;   // Graph, sign-in and SharePoint go straight to the network
  /* network first (always the latest version when online), cache when offline */
  e.respondWith(fetch(e.request, { cache: "no-cache" }).then(r => { const copy = r.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); return r; })
    .catch(() => caches.match(e.request, { ignoreSearch: true }).then(r => r || caches.match("index.html"))));
});
