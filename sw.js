// Offline cache. Bump VERSION on every release so users get the new files on their next visit.
const VERSION = "v3";
const FILES = [
  "./", "index.html", "style.css", "app.js", "worker.js", "backend.py", "manifest.webmanifest",
  "icons/icon.svg", "icons/favicon-32.png", "icons/icon-192.png", "icons/icon-512.png",
  "vendor/fonts/Geist-Variable.woff2", "vendor/fonts/GeistMono-Variable.woff2",
  "vendor/pyodide/pyodide.js", "vendor/pyodide/pyodide.asm.js", "vendor/pyodide/pyodide.asm.wasm",
  "vendor/pyodide/python_stdlib.zip", "vendor/pyodide/pyodide-lock.json",
  "vendor/wheels/charset_normalizer-3.5.1-py3-none-any.whl",
  "vendor/wheels/et_xmlfile-2.0.0-py3-none-any.whl",
  "vendor/wheels/openpyxl-3.1.5-py2.py3-none-any.whl",
  "vendor/wheels/pdfminer_six-20260107-py3-none-any.whl",
  "vendor/wheels/pdfplumber-0.11.10-py3-none-any.whl",
];

self.addEventListener("install", (e) => e.waitUntil(caches.open(VERSION).then((c) => c.addAll(FILES)).then(() => self.skipWaiting())));
self.addEventListener("activate", (e) => e.waitUntil(
  caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== VERSION).map((k) => caches.delete(k)))).then(() => self.clients.claim())
));
// network first, so an update shows up as soon as the user is online; cache when offline
self.addEventListener("fetch", (e) => {
  if (e.request.method !== "GET" || new URL(e.request.url).origin !== location.origin) return;
  e.respondWith(
    fetch(e.request).then((r) => {
      const copy = r.clone();
      if (r.ok) caches.open(VERSION).then((c) => c.put(e.request, copy));
      return r;
    }).catch(() => caches.match(e.request, { ignoreSearch: true }))
  );
});
