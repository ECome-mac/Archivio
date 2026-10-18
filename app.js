"use strict";
// Archivio medico - consultazione in sola lettura da OneDrive (Microsoft Graph, permesso Files.Read)
const CFG = window.ARCHIVIO_CONFIG || {};
const NR = "Non rilevato";
const CATALOGO = "_catalogo-app.json";
const AUTH = "https://login.microsoftonline.com/consumers/oauth2/v2.0";
const GRAPH = "https://graph.microsoft.com/v1.0/me/drive/root:/";
const SCOPE = "https://graph.microsoft.com/Files.Read offline_access";
const REDIRECT = new URL("./", location.href).href;
const CACHE_DOC = "archivio-documenti";
// prova in locale con dati inventati: http://localhost:<porta>/?demo
const DEMO = /^(localhost|127\.0\.0\.1)$/.test(location.hostname) && new URLSearchParams(location.search).has("demo");
const PDF_WORKER = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

const $ = (s) => document.querySelector(s);
const leggi = (k, d) => { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (_) { return d; } };
const scrivi = (k, v) => { try { v == null ? localStorage.removeItem(k) : localStorage.setItem(k, JSON.stringify(v)); } catch (_) {} };

let catalogo = leggi("catalogo", null);
let salvati = new Set(leggi("salvati", []));
let categoria = "";

// ---------- accesso Microsoft (OAuth 2.0 con PKCE, nessun segreto nell'app) ----------
const b64url = (buf) => btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
const casuale = (n) => b64url(crypto.getRandomValues(new Uint8Array(n)));

async function accedi() {
  const verifier = casuale(48), state = casuale(16);
  const challenge = b64url(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier)));
  scrivi("pkce", { verifier, state });
  location.assign(AUTH + "/authorize?" + new URLSearchParams({
    client_id: CFG.clientId, response_type: "code", redirect_uri: REDIRECT, response_mode: "query",
    scope: SCOPE, state, code_challenge: challenge, code_challenge_method: "S256", prompt: "select_account",
  }));
}

async function chiediToken(params) {
  const r = await fetch(AUTH + "/token", {
    method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id: CFG.clientId, scope: SCOPE, ...params }),
  });
  const j = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error(j.error_description || "Accesso non riuscito");
  const t = { access: j.access_token, refresh: j.refresh_token, scade: Date.now() + (j.expires_in - 120) * 1000 };
  scrivi("token", t);
  return t;
}

async function ritornoDaAccesso() {
  const q = new URLSearchParams(location.search);
  if (!q.has("code") && !q.has("error")) return;
  history.replaceState(null, "", REDIRECT);
  const p = leggi("pkce", null); scrivi("pkce", null);
  if (q.has("error")) throw new Error(q.get("error_description") || "Accesso annullato");
  if (!p || p.state !== q.get("state")) throw new Error("Accesso non valido: riprova");
  await chiediToken({ grant_type: "authorization_code", code: q.get("code"), redirect_uri: REDIRECT, code_verifier: p.verifier });
}

class ServeAccesso extends Error {}
async function token(forza) {
  const t = leggi("token", null);
  if (!t) throw new ServeAccesso();
  if (!forza && t.scade > Date.now()) return t.access;
  try { return (await chiediToken({ grant_type: "refresh_token", refresh_token: t.refresh })).access; }
  catch (e) {
    if (e instanceof TypeError) throw new Error("Connessione assente: riprova quando sei online.");   // offline: l'accesso resta valido
    scrivi("token", null); throw new ServeAccesso();
  }
}

async function graph(url, opz = {}) {
  for (let tentativo = 0; tentativo < 2; tentativo++) {
    const r = await fetch(url, { ...opz, headers: { Authorization: "Bearer " + await token(tentativo > 0) } });
    if (r.status !== 401) return r;
  }
  scrivi("token", null);
  throw new ServeAccesso();
}

const percorsoGraph = (rel) => GRAPH + [CFG.cartella || "Archivio medico", ...rel.split("/")].map(encodeURIComponent).join("/");

// ---------- dati ----------
async function scaricaCatalogo() {
  if (DEMO) return (await fetch("demo/catalogo.json", { cache: "no-store" })).json();
  const r = await graph(percorsoGraph(CATALOGO) + ":/content");
  if (r.status === 404) throw new Error("Catalogo non trovato in OneDrive: premi Sincronizza sul PC.");
  if (!r.ok) throw new Error("OneDrive non risponde (" + r.status + "): riprova più tardi.");
  return r.json();
}

async function scaricaDocumento(d) {
  if (DEMO) { const r = await fetch("demo/" + d.percorso.split("/").map(encodeURIComponent).join("/")); if (!r.ok) throw new Error("File non trovato"); return r.blob(); }
  const r = await graph(percorsoGraph(d.percorso) + "?select=id,@microsoft.graph.downloadUrl");
  if (r.status === 404) throw new Error("Documento non trovato in OneDrive: forse è stato spostato. Aggiorna l'elenco.");
  if (!r.ok) throw new Error("OneDrive non risponde (" + r.status + ")");
  const url = (await r.json())["@microsoft.graph.downloadUrl"];
  const f = await fetch(url);
  if (!f.ok) throw new Error("Download non riuscito (" + f.status + ")");
  return f.blob();
}

const chiaveCache = (d) => new URL("_doc/" + d.id, REDIRECT).href;
async function documento(d) {
  let cache = null;
  try { cache = await caches.open(CACHE_DOC); } catch (_) {}
  const c = cache && await cache.match(chiaveCache(d));
  if (c) return c.blob();
  const b = await scaricaDocumento(d);
  if (cache) {
    try {
      await cache.put(chiaveCache(d), new Response(b, { headers: { "Content-Type": b.type || "application/octet-stream" } }));
      salvati.add(d.id); scrivi("salvati", [...salvati]);
    } catch (_) {}
  }
  return b;
}

async function aggiorna(silenzioso) {
  const btn = $("#btn-aggiorna");
  btn.disabled = true;
  if (!silenzioso) messaggio("Aggiornamento…");
  try {
    const c = await scaricaCatalogo();
    if (!c || !Array.isArray(c.documenti)) throw new Error("Catalogo non leggibile");
    catalogo = c; scrivi("catalogo", c);
    // i documenti non più in archivio escono anche dalla memoria del telefono
    const ids = new Set(c.documenti.map((d) => d.id));
    const vecchi = [...salvati].filter((id) => !ids.has(id));
    if (vecchi.length) {
      try { const cache = await caches.open(CACHE_DOC); await Promise.all(vecchi.map((id) => cache.delete(chiaveCache({ id })))); } catch (_) {}
      vecchi.forEach((id) => salvati.delete(id)); scrivi("salvati", [...salvati]);
    }
    messaggio("");
    preparaFiltri(); disegna();
  } catch (e) {
    if (e instanceof ServeAccesso) return mostra("accesso");
    messaggio(navigator.onLine ? e.message : "Sei offline: elenco aggiornato al " + quando(catalogo && catalogo.aggiornato) + ".", navigator.onLine);
  } finally { btn.disabled = false; }
}

// ---------- elenco ----------
function messaggio(t, errore) { const m = $("#msg"); m.textContent = t || ""; m.classList.toggle("bad", !!errore); }
function quando(iso) {
  if (!iso) return "?";
  const d = new Date(iso);
  return isNaN(d) ? iso : d.toLocaleDateString("it-IT") + " " + d.toLocaleTimeString("it-IT", { hour: "2-digit", minute: "2-digit" });
}
const fmtData = (d) => (/^\d{4}-\d{2}-\d{2}$/.test(d) ? d.slice(8) + "/" + d.slice(5, 7) + "/" + d.slice(0, 4) : d || NR);
const anno = (d) => (/^\d{4}/.test(d || "") ? d.slice(0, 4) : NR);
const docs = () => (catalogo ? catalogo.documenti : []);
const inFondo = (v) => (v === NR ? 2 : v === "Altro" ? 1 : 0);
const ordina = (a, b) => inFondo(a) - inFondo(b) || a.localeCompare(b, "it");
const valori = (k) => [...new Set(docs().map((d) => d[k] || NR))].sort(ordina);

function preparaFiltri() {
  const conta = {};
  for (const d of docs()) conta[d.categoria] = (conta[d.categoria] || 0) + 1;
  if (categoria && !conta[categoria]) categoria = "";
  const chip = (val, testo) => {
    const b = document.createElement("button");
    b.type = "button"; b.textContent = testo; b.setAttribute("aria-pressed", String(categoria === val));
    b.addEventListener("click", () => { categoria = val; preparaFiltri(); disegna(); });
    return b;
  };
  $("#chips").replaceChildren(chip("", "Tutte (" + docs().length + ")"), ...valori("categoria").map((c) => chip(c, c + " (" + conta[c] + ")")));
  for (const [sel, k, label] of [["#f-str", "struttura", "Struttura: tutte"], ["#f-med", "medico", "Medico: tutti"]]) {
    const s = $(sel), cur = s.value;
    s.replaceChildren(new Option(label, ""), ...valori(k).map((v) => new Option(v, v)));
    s.value = [...s.options].some((o) => o.value === cur) ? cur : "";
  }
}

function filtrati() {
  const q = $("#f-q").value.trim().toLowerCase();
  const da = $("#f-da").value, a = $("#f-a").value, s = $("#f-str").value, m = $("#f-med").value;
  return docs().filter((d) => {
    if (categoria && d.categoria !== categoria) return false;
    if (s && d.struttura !== s) return false;
    if (m && d.medico !== m) return false;
    if ((da || a) && !/^\d{4}-/.test(d.data_esame)) return false;
    if (da && d.data_esame < da) return false;
    if (a && d.data_esame > a) return false;
    if (q && ![d.titolo, d.esame, d.categoria, d.struttura, d.medico, fmtData(d.data_esame), d.data_esame]
      .some((v) => (v || "").toLowerCase().includes(q))) return false;
    return true;
  });
}

function confronto() {
  const o = $("#f-sort").value;
  const perData = (dir) => (x, y) => {
    const a = /^\d{4}/.test(x.data_esame) ? x.data_esame : null, b = /^\d{4}/.test(y.data_esame) ? y.data_esame : null;
    if (!a || !b) return (a ? -1 : 0) + (b ? 1 : 0);    // date non rilevate sempre in fondo
    return a.localeCompare(b) * dir;
  };
  if (o === "titolo") return (x, y) => x.titolo.localeCompare(y.titolo, "it", { numeric: true });
  return (x, y) => perData(o === "data-asc" ? 1 : -1)(x, y) || x.titolo.localeCompare(y.titolo, "it");
}

function voce(d) {
  const li = document.createElement("li");
  const b = document.createElement("button");
  b.type = "button"; b.className = "doc";
  const t = document.createElement("span"); t.className = "titolo"; t.textContent = d.titolo;
  const r1 = document.createElement("span"); r1.className = "riga";
  const cat = document.createElement("span"); cat.className = "cat"; cat.textContent = d.categoria;
  const dt = document.createElement("span"); dt.textContent = fmtData(d.data_esame); if (d.data_esame === NR) dt.className = "nr";
  r1.append(cat, dt);
  if (salvati.has(d.id)) { const sv = document.createElement("span"); sv.className = "salvato"; sv.textContent = "  ● sul telefono"; r1.append(sv); }
  const r2 = document.createElement("span"); r2.className = "riga";
  r2.textContent = [d.struttura, d.medico].filter((v) => v && v !== NR).join(" · ");
  b.append(t, r1);
  if (r2.textContent) b.append(r2);
  b.addEventListener("click", () => apri(d));
  li.append(b);
  return li;
}

function disegna() {
  const tutti = docs();
  const stat = [
    [tutti.length, "Documenti"],
    [new Set(tutti.map((d) => d.categoria).filter((v) => v !== NR && v !== "Altro")).size, "Categorie"],
    [catalogo && !isNaN(new Date(catalogo.aggiornato)) ? new Date(catalogo.aggiornato).toLocaleDateString("it-IT", { day: "2-digit", month: "2-digit", year: "2-digit" }) : "–", "Aggiornato al"],
  ];
  $("#stats").replaceChildren(...stat.map(([n, l]) => {
    const div = document.createElement("div"); div.className = "stat";
    const b = document.createElement("b"); b.textContent = n;
    const sp = document.createElement("span"); sp.textContent = l;
    div.append(b, sp); return div;
  }));
  $("#menu-info").textContent = "Elenco aggiornato al " + quando(catalogo && catalogo.aggiornato) + ". " +
    salvati.size + " document" + (salvati.size === 1 ? "o salvato" : "i salvati") + " sul telefono per la consultazione offline.";
  const lista = filtrati().sort(confronto());
  const g = $("#f-group").value;
  const righe = [];
  if (g) {
    const chiave = (d) => (g === "anno" ? anno(d.data_esame) : d[g] || NR);
    const gruppi = new Map();
    for (const d of lista) { const k = chiave(d); if (!gruppi.has(k)) gruppi.set(k, []); gruppi.get(k).push(d); }
    const nomi = [...gruppi.keys()].sort((a, b) => (g === "anno" ? inFondo(a) - inFondo(b) || b.localeCompare(a) : ordina(a, b)));
    for (const n of nomi) {
      const li = document.createElement("li"); li.className = "gruppo"; li.textContent = n + " (" + gruppi.get(n).length + ")";
      righe.push(li, ...gruppi.get(n).map(voce));
    }
  } else righe.push(...lista.map(voce));
  $("#lista").replaceChildren(...righe);
  $("#count").textContent = lista.length + " document" + (lista.length === 1 ? "o" : "i");
}

// ---------- visualizzazione ----------
let aperto = null;
let urlImmagine = null;
// il worker di PDF.js viene caricato come blob: un worker non può partire da un altro dominio
let workerPdf = null;
async function preparaWorker() {
  if (workerPdf || !window.pdfjsLib) return;
  const testo = await (await fetch(PDF_WORKER)).text();
  workerPdf = new Worker(URL.createObjectURL(new Blob([testo], { type: "text/javascript" })));
  pdfjsLib.GlobalWorkerOptions.workerPort = workerPdf;
}

function chiudiViewer(daStoria) {
  if (!aperto) return;
  aperto = null;
  $("#viewer").hidden = true;
  $("#v-pagine").replaceChildren();
  if (urlImmagine) { URL.revokeObjectURL(urlImmagine); urlImmagine = null; }
  document.body.style.overflow = "";
  if (!daStoria && history.state && history.state.viewer) history.back();
  disegna();
}

async function apri(d) {
  aperto = d;
  const v = $("#viewer");
  $("#v-titolo").textContent = d.titolo;
  const meta = [["Data", fmtData(d.data_esame)], ["Categoria", d.categoria], ["Tecnica", d.esame], ["Struttura", d.struttura], ["Medico", d.medico]];
  $("#v-meta").replaceChildren(...meta.filter(([, x]) => x && x !== NR).flatMap(([k, x]) => {
    const dt = document.createElement("dt"); dt.textContent = k;
    const dd = document.createElement("dd"); dd.textContent = x;
    return [dt, dd];
  }));
  $("#v-pagine").replaceChildren();
  $("#v-condividi").hidden = true;
  stato("Apertura del documento…");
  v.hidden = false; v.scrollTop = 0;
  document.body.style.overflow = "hidden";
  history.pushState({ viewer: true }, "");
  try {
    const blob = await documento(d);
    if (aperto !== d) return;
    const nome = d.percorso.split("/").pop();
    const pdf = /\.pdf$/i.test(nome) || blob.type === "application/pdf";
    const file = new File([blob], nome, { type: blob.type || (pdf ? "application/pdf" : "image/jpeg") });
    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      const b = $("#v-condividi"); b.hidden = false;
      b.onclick = () => navigator.share({ files: [file], title: d.titolo }).catch(() => {});
    }
    if (pdf) await mostraPdf(blob, d); else mostraImmagine(blob);
  } catch (e) {
    if (aperto !== d) return;
    if (e instanceof ServeAccesso) { chiudiViewer(); return mostra("accesso"); }
    stato(navigator.onLine ? e.message : "Sei offline e questo documento non è ancora salvato sul telefono.", true);
  }
}

function stato(t, errore) { const s = $("#v-stato"); s.textContent = t || ""; s.classList.toggle("bad", !!errore); }

function mostraImmagine(blob) {
  urlImmagine = URL.createObjectURL(blob);
  const img = new Image();
  img.alt = "Documento";
  img.onload = () => stato("");
  img.onerror = () => stato("Formato immagine non supportato da questo telefono: usa Condividi per aprirla.", true);
  img.src = urlImmagine;
  $("#v-pagine").append(img);
}

async function mostraPdf(blob, d) {
  if (!window.pdfjsLib) throw new Error("Lettore PDF non disponibile: controlla la connessione e riprova.");
  await preparaWorker();
  const doc = await pdfjsLib.getDocument({ data: new Uint8Array(await blob.arrayBuffer()), isEvalSupported: false }).promise;
  if (aperto !== d) return doc.destroy();
  stato(doc.numPages > 1 ? doc.numPages + " pagine" : "");
  const box = $("#v-pagine");
  const larghezza = Math.min(box.clientWidth || 360, 900);
  const scalaSchermo = Math.min(window.devicePixelRatio || 1, 2.5);
  const prima = await doc.getPage(1);
  const rapporto = prima.getViewport({ scale: 1 }).height / prima.getViewport({ scale: 1 }).width;
  const disegnaPagina = async (n, segnaposto) => {
    const p = await doc.getPage(n);
    const base = p.getViewport({ scale: 1 });
    const vp = p.getViewport({ scale: (larghezza / base.width) * scalaSchermo * 1.5 });
    const cv = document.createElement("canvas");
    cv.width = Math.floor(vp.width); cv.height = Math.floor(vp.height);
    cv.setAttribute("aria-label", "Pagina " + n + " di " + doc.numPages);
    await p.render({ canvasContext: cv.getContext("2d"), viewport: vp }).promise;
    if (aperto === d) segnaposto.replaceWith(cv);
    p.cleanup();
  };
  // le pagine si disegnano quando stanno per entrare nello schermo, per non consumare memoria
  const osservatore = new IntersectionObserver((voci) => {
    for (const v of voci) if (v.isIntersecting) { osservatore.unobserve(v.target); disegnaPagina(+v.target.dataset.n, v.target).catch(() => {}); }
  }, { root: $("#viewer"), rootMargin: "800px 0px" });
  for (let n = 1; n <= doc.numPages; n++) {
    const s = document.createElement("div"); s.className = "segnaposto"; s.dataset.n = n;
    s.style.height = Math.round(larghezza * rapporto) + "px";
    box.append(s); osservatore.observe(s);
  }
}

// ---------- avvio ----------
function mostra(cosa) {
  for (const id of ["accesso", "config", "app"]) $("#" + id).hidden = id !== cosa;
}

$("#btn-accedi").addEventListener("click", () => accedi().catch((e) => ($("#accesso-err").textContent = e.message)));
$("#btn-aggiorna").addEventListener("click", () => aggiorna(false));
$("#btn-menu").addEventListener("click", () => {
  const m = $("#menu"); m.hidden = !m.hidden; $("#btn-menu").setAttribute("aria-expanded", String(!m.hidden));
});
$("#btn-filtri").addEventListener("click", () => {
  const a = $("#altri"); a.hidden = !a.hidden;
  $("#btn-filtri").setAttribute("aria-expanded", String(!a.hidden));
  $("#btn-filtri").textContent = a.hidden ? "Altri filtri" : "Nascondi filtri";
});
$("#btn-svuota").addEventListener("click", async () => {
  if (!confirm("Eliminare dal telefono i documenti salvati per la consultazione offline?\n\nI documenti restano in OneDrive.")) return;
  try { await caches.delete(CACHE_DOC); } catch (_) {}
  salvati = new Set(); scrivi("salvati", []); disegna();
});
$("#btn-esci").addEventListener("click", async () => {
  if (!confirm("Uscire? L'elenco e i documenti salvati sul telefono verranno cancellati.")) return;
  try { await caches.delete(CACHE_DOC); } catch (_) {}
  for (const k of ["token", "catalogo", "salvati", "pkce"]) scrivi(k, null);
  catalogo = null; salvati = new Set();
  location.replace(REDIRECT);
});
["#f-q", "#f-da", "#f-a", "#f-str", "#f-med", "#f-group", "#f-sort"].forEach((s) =>
  $(s).addEventListener(s === "#f-q" ? "input" : "change", disegna));
$("#btn-reset").addEventListener("click", () => {
  ["#f-q", "#f-da", "#f-a", "#f-str", "#f-med", "#f-group"].forEach((s) => ($(s).value = ""));
  $("#f-sort").value = "data-desc"; categoria = ""; preparaFiltri(); disegna();
});
$("#v-chiudi").addEventListener("click", () => chiudiViewer(false));
window.addEventListener("popstate", () => chiudiViewer(true));
window.addEventListener("online", () => aggiorna(true));

if ("serviceWorker" in navigator && !DEMO) navigator.serviceWorker.register("sw.js").catch(() => {});

(async () => {
  if (!DEMO && !CFG.clientId) return mostra("config");
  try { await ritornoDaAccesso(); } catch (e) { mostra("accesso"); $("#accesso-err").textContent = e.message; return; }
  if (!DEMO && !leggi("token", null)) return mostra("accesso");
  mostra("app");
  if (catalogo) { preparaFiltri(); disegna(); }
  aggiorna(!!catalogo);
})();
