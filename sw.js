// Tiene in memoria i file dell'app per aprirla anche offline. I dati arrivano da OneDrive e non passano di qui.
const VERSIONE = "archivio-app-1";
const FILE = ["./", "index.html", "app.css", "app.js", "config.js", "manifest.webmanifest", "icon-180.png", "icon-192.png", "icon-512.png",
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js",
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js"];
const URLS = new Set(FILE.map((f) => new URL(f, self.location).href));

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(VERSIONE).then((c) => c.addAll(FILE)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys()
    .then((ks) => Promise.all(ks.filter((k) => k.startsWith("archivio-app-") && k !== VERSIONE).map((k) => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener("fetch", (e) => {
  const u = new URL(e.request.url);
  if (e.request.method !== "GET") return;
  const pagina = e.request.mode === "navigate" && u.origin === self.location.origin;
  const chiave = pagina ? new URL("./", self.location).href : u.origin + u.pathname;
  if (!pagina && !URLS.has(chiave)) return;          // OneDrive, accesso Microsoft ecc.: sempre dalla rete
  // prima la rete (per ricevere gli aggiornamenti dell'app), poi la copia salvata se si è offline
  e.respondWith(fetch(e.request).then((r) => {
    if (r.ok) { const copia = r.clone(); caches.open(VERSIONE).then((c) => c.put(chiave, copia)); }
    return r;
  }).catch(() => caches.match(chiave).then((r) => r || Response.error())));
});
