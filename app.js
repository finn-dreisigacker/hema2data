"use strict";

const T = {
  de: {
    install: "Als App installieren", restart: "Neu starten",
    dropTitle: "Blutbild-PDFs oder einen ganzen Ordner hierher ziehen.",
    pickFiles: "Dateien auswählen", pickFolder: "Ordner auswählen", dropNow: "Loslassen zum Einlesen",
    local: "Läuft nur in diesem Browser",
    booting: "Auswertung wird vorbereitet …", reading: "PDF {i} von {n} wird gelesen …",
    okHead: "Keine Fehler: alle {n} PDFs vollständig gelesen",
    errHead: "{e} von {n} PDFs mit Fehlern",
    warnHead: "Keine Fehler, aber {w} von {n} PDFs mit Warnungen",
    warnHint: "Das Gerät hat Werte markiert. Bitte im PDF prüfen; die Werte stehen in der Excel.",
    errHint: "Die Meldungen stehen in der Tabelle. Betroffene Werte bitte im PDF prüfen; fehlende Zellen bleiben in der Excel leer.",
    noPdf: "Keine PDF-Dateien gefunden.",
    files: "Dateien", colFile: "Datei", colParticipant: "Proband", colDate: "Entnahme", colStatus: "Status", colAction: "Aktion",
    ok: "OK", warn: "Warnung", error: "Fehler", openPdf: "PDF öffnen",
    preview: "Vorschau der Excel", fileName: "blutbilder.xlsx",
    sheets: "Tabellenblätter", exportHint: "Fehlende Werte bleiben leer.",
    advanced: "Erweitert", addCheck: "Blatt „Check“ mit allen Meldungen hinzufügen",
    download: "Excel herunterladen",
    more: "Weitere PDFs oder einen Ordner hierher ziehen. Die Dateien werden nur gelesen, nicht hochgeladen.",
    madeBy: "Erstellt von",
    failed: "Fehler bei der Auswertung: {m}",
  },
  en: {
    install: "Install as app", restart: "Start over",
    dropTitle: "Drop blood count PDFs or a whole folder here.",
    pickFiles: "Choose files", pickFolder: "Choose folder", dropNow: "Release to read",
    local: "Runs only in this browser",
    booting: "Preparing …", reading: "Reading PDF {i} of {n} …",
    okHead: "No errors: all {n} PDFs read completely",
    errHead: "{e} of {n} PDFs have errors",
    warnHead: "No errors, but {w} of {n} PDFs have warnings",
    warnHint: "The device flagged some values. Please check the PDF; the values are in the Excel file.",
    errHint: "The messages are in the table. Please check the affected values in the PDF; missing cells stay empty in the Excel file.",
    noPdf: "No PDF files found.",
    files: "Files", colFile: "File", colParticipant: "Participant", colDate: "Collected", colStatus: "Status", colAction: "Action",
    ok: "OK", warn: "Warning", error: "Error", openPdf: "Open PDF",
    preview: "Excel preview", fileName: "bloodcounts.xlsx",
    sheets: "Sheets", exportHint: "Missing values stay empty.",
    advanced: "Advanced", addCheck: "Add a “Check” sheet with every message",
    download: "Download Excel",
    more: "Drop more PDFs or a folder here. Files are only read, never uploaded.",
    madeBy: "Made by",
    failed: "Processing failed: {m}",
  },
};

const $ = (id) => document.getElementById(id);
const state = { lang: "de", results: [], files: [], sheets: null, sheet: null, xlsx: null, step: "start", progress: null };
const worker = new Worker("worker.js");
worker.postMessage({ type: "warmup" }); // load Pyodide while the user picks files

function t(key, vars = {}) {
  return T[state.lang][key].replace(/\{(\w+)\}/g, (_, k) => vars[k]);
}

try { state.lang = localStorage.getItem("lang") || (navigator.language.startsWith("de") ? "de" : "en"); } catch { /* storage blocked */ }

function setLang(lang) {
  state.lang = lang;
  try { localStorage.setItem("lang", lang); } catch { /* storage blocked */ }
  document.documentElement.lang = lang;
  document.querySelectorAll("[data-i18n]").forEach((el) => { el.textContent = t(el.dataset.i18n); });
  document.querySelectorAll("[data-lang]").forEach((b) => b.setAttribute("aria-pressed", String(b.dataset.lang === lang)));
  // follow the language only while the user has not typed their own name
  if (!state.nameEdited) $("filename").value = t("fileName");
  renderLoading();
  if (state.step === "done") renderDone();
}

function show(step) {
  state.step = step;
  for (const id of ["start", "loading", "done"]) $(id).hidden = id !== step;
  $("brand").hidden = $("restart").hidden = step === "start";
}

// ---- reading files and folders ----

const isPdf = (f) => f.name.toLowerCase().endsWith(".pdf");

async function entryFiles(entry) {
  if (entry.isFile) return [await new Promise((res, rej) => entry.file(res, rej))];
  const reader = entry.createReader();
  const out = [];
  for (;;) {
    const batch = await new Promise((res, rej) => reader.readEntries(res, rej));
    if (!batch.length) break;
    for (const e of batch) out.push(...(await entryFiles(e)));
  }
  return out;
}

async function droppedFiles(dt) {
  const entries = [...dt.items].map((i) => i.webkitGetAsEntry?.()).filter(Boolean);
  if (!entries.length) return [...dt.files];
  return (await Promise.all(entries.map(entryFiles))).flat();
}

async function process(list) {
  const pdfs = list.filter(isPdf).sort((a, b) => (a.webkitRelativePath || a.name).localeCompare(b.webkitRelativePath || b.name));
  if (!pdfs.length) { alert(t("noPdf")); return; }
  show("loading");
  state.progress = null;
  renderLoading();
  const files = await Promise.all(pdfs.map(async (f) => ({ name: f.name, bytes: await f.arrayBuffer() })));
  state.files.push(...pdfs);
  worker.postMessage({ type: "run", files }, files.map((f) => f.bytes));
}

worker.onmessage = ({ data }) => {
  if (data.type === "progress") { state.progress = data; renderLoading(); }
  else if (data.type === "result") state.results.push(data.result);
  else if (data.type === "done") buildExcel();
  else if (data.type === "excel") {
    state.xlsx = data.bytes;
    state.sheets = data.sheets;
    if (!state.sheets[state.sheet]) state.sheet = Object.keys(state.sheets)[0];
    show("done");
    renderDone();
  } else if (data.type === "error") {
    alert(t("failed", { m: data.message }));
    show(state.results.length ? "done" : "start");
  }
};

function chosenSheets() {
  const out = { Values: false, Overview: true, Check: false };
  document.querySelectorAll("[data-sheet]").forEach((c) => { out[c.dataset.sheet] = c.checked; });
  return out;
}

function buildExcel() {
  $("download").disabled = !state.results.length;
  if (state.results.length) worker.postMessage({ type: "excel", sheets: chosenSheets() });
}

// ---- rendering ----

function renderLoading() {
  const p = state.progress;
  $("loading-text").textContent = p ? t("reading", p) : t("booting");
  $("track-fill").style.width = p ? `${Math.round((p.i / p.n) * 100)}%` : "0%";
}

const ICON_OK = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="9"/><path d="m8.5 12 2.5 2.5 4.5-5"/></svg>';
const ICON_BAD = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="9"/><path d="M12 7.5v5.5M12 16.5h.01"/></svg>';

function el(tag, cls, text) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text != null) e.textContent = text;
  return e;
}

function renderDone() {
  const rs = state.results;
  const n = rs.length;
  const e = rs.filter((r) => r.errors.length).length;
  const w = rs.filter((r) => r.warnings.length).length;
  const kind = e ? "bad" : w ? "warn" : "ok";

  const notice = $("notice");
  notice.className = "notice " + kind;
  notice.innerHTML = kind === "ok" ? ICON_OK : ICON_BAD;
  const box = el("div");
  box.append(el("strong", null, { bad: t("errHead", { e, n }), warn: t("warnHead", { w, n }), ok: t("okHead", { n }) }[kind]));
  if (kind !== "ok") box.append(el("span", null, t(kind === "bad" ? "errHint" : "warnHint")));
  notice.append(box);

  const body = $("file-rows");
  body.replaceChildren();
  rs.forEach((r, idx) => {
    const tr = el("tr");
    tr.append(el("td", "file", r.file), el("td", null, r.participant_id || "–"));
    const d = el("td", "num");
    if (r.collected_at) { d.append(r.collected_at.slice(0, 10) + " "); d.append(el("span", "muted", r.collected_at.slice(11, 16))); } else d.textContent = "–";
    tr.append(d);

    const st = el("td");
    st.append(r.errors.length ? el("span", "pill bad", t("error")) : r.warnings.length ? el("span", "pill warn", t("warn")) : el("span", "pill ok", t("ok")));
    const msgs = [...r.errors, ...r.warnings];
    if (msgs.length) {
      const ul = el("ul", "msgs");
      msgs.forEach((m) => ul.append(el("li", null, m)));
      st.append(ul);
    }
    tr.append(st);

    const act = el("td");
    if (r.errors.length || r.warnings.length) {
      const b = el("button", "btn btn-small", t("openPdf"));
      b.onclick = () => window.open(URL.createObjectURL(state.files[idx]), "_blank", "noopener");
      act.append(b);
    }
    tr.append(act);
    body.append(tr);
  });

  renderSheet();
}

const colName = (i) => (i >= 26 ? colName(Math.floor(i / 26) - 1) : "") + String.fromCharCode(65 + (i % 26));

function renderSheet() {
  const names = Object.keys(state.sheets || {});
  const tabs = $("tabs");
  tabs.replaceChildren();
  names.forEach((name) => {
    const b = el("button", null, name);
    b.setAttribute("role", "tab");
    b.setAttribute("aria-selected", String(name === state.sheet));
    b.onclick = () => { state.sheet = name; renderSheet(); };
    tabs.append(b);
  });

  const rows = state.sheets?.[state.sheet] || [];
  const width = Math.max(0, ...rows.map((r) => r.length));
  const statusCol = rows[0]?.indexOf("Status") ?? -1;
  const table = el("table");
  const head = el("tr");
  head.append(el("th", "corner"));
  for (let c = 0; c < width; c++) head.append(el("th", null, colName(c)));
  const thead = el("thead");
  thead.append(head);
  table.append(thead);
  const tb = el("tbody");
  rows.forEach((r, ri) => {
    const tr = el("tr", ri === 0 ? "head" : null);
    const th = el("th", null, String(ri + 1));
    th.scope = "row";
    tr.append(th);
    for (let c = 0; c < width; c++) {
      const v = r[c];
      const td = el("td", typeof v === "number" ? "n" : null, v == null ? "" : String(v));
      // an empty parameter cell in an errored row is the value that could not be read
      if (ri > 0 && statusCol >= 0 && c > statusCol && v == null && r[statusCol] === "error") td.classList.add("missing");
      tr.append(td);
    }
    tb.append(tr);
  });
  table.append(tb);
  $("sheet").replaceChildren(table);
}

// ---- wiring ----

document.querySelectorAll("[data-lang]").forEach((b) => { b.onclick = () => setLang(b.dataset.lang); });
for (const id of ["pick-files", "pick-folder", "pick-more"]) {
  $(id).addEventListener("change", (ev) => { process([...ev.target.files]); ev.target.value = ""; });
}
document.querySelectorAll("[data-sheet]").forEach((c) => c.addEventListener("change", () => buildExcel()));

$("filename").addEventListener("input", () => { state.nameEdited = $("filename").value.trim() !== ""; });

function fileName() {
  const name = $("filename").value.trim().replace(/[\\/:*?"<>|]+/g, "_") || t("fileName");
  return /\.xlsx$/i.test(name) ? name : name + ".xlsx";
}

$("download").onclick = () => {
  const url = URL.createObjectURL(new Blob([state.xlsx], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  const a = Object.assign(document.createElement("a"), { href: url, download: fileName() });
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
};

$("restart").onclick = () => {
  worker.postMessage({ type: "reset" });
  Object.assign(state, { results: [], files: [], sheets: null, xlsx: null });
  show("start");
};

let dragDepth = 0;
window.addEventListener("dragenter", (ev) => { if (state.step === "loading") return; ev.preventDefault(); dragDepth++; $("dropping").hidden = false; });
window.addEventListener("dragleave", () => { if (--dragDepth <= 0) { dragDepth = 0; $("dropping").hidden = true; } });
window.addEventListener("dragover", (ev) => ev.preventDefault());
window.addEventListener("drop", async (ev) => {
  ev.preventDefault();
  dragDepth = 0;
  $("dropping").hidden = true;
  if (state.step !== "loading") process(await droppedFiles(ev.dataTransfer));
});

// address is only assembled on click, so it never appears in the HTML that spam bots scrape
$("mail").onclick = (ev) => {
  ev.preventDefault();
  location.href = "mailto:" + ["dreisigacker.finn", "web.de"].join("@");
};

let installPrompt = null;
window.addEventListener("beforeinstallprompt", (ev) => { ev.preventDefault(); installPrompt = ev; $("install").hidden = false; });
$("install").onclick = async () => { await installPrompt?.prompt(); installPrompt = null; $("install").hidden = true; };

if ("serviceWorker" in navigator && location.protocol !== "file:") navigator.serviceWorker.register("sw.js");

setLang(state.lang);
show("start");
