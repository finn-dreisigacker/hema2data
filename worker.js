// Runs backend.py unchanged in Pyodide. Everything happens in this worker's memory; nothing is sent anywhere.
importScripts("vendor/pyodide/pyodide.js");

const WHEELS = [
  "charset_normalizer-3.5.1-py3-none-any.whl",
  "et_xmlfile-2.0.0-py3-none-any.whl",
  "openpyxl-3.1.5-py2.py3-none-any.whl",
  "pdfminer_six-20260107-py3-none-any.whl",
  "pdfplumber-0.11.10-py3-none-any.whl",
];

const GLUE = `
import sys, types, importlib.abc, importlib.machinery, json
from pathlib import Path

# pypdfium2 (images) and cryptography (encrypted PDFs) have no browser build; backend.py needs neither.
class _Stub(types.ModuleType):
    __path__ = []
    def __getattr__(self, name):
        if name.startswith("__"): raise AttributeError(name)
        return _Stub(f"{self.__name__}.{name}")
    def __call__(self, *a, **k):
        raise RuntimeError(f"{self.__name__} is not available in the browser version (encrypted PDF?)")
class _Finder(importlib.abc.MetaPathFinder, importlib.abc.Loader):
    def find_spec(self, name, path, target=None):
        if name.split(".")[0] in ("pypdfium2", "cryptography"):
            return importlib.machinery.ModuleSpec(name, self)
    def create_module(self, spec): return _Stub(spec.name)
    def exec_module(self, m): pass
sys.meta_path.insert(0, _Finder())

import backend

ergebnisse, gesehen = [], {}

# Same per-file handling as backend.main(): unreadable PDFs and duplicates are reported, not fatal.
def lies(pfad):
    pfad = Path(pfad)
    try:
        e = backend.lies_pdf(pfad)
    except Exception as fehler:  # noqa: BLE001
        e = {"file": pfad.name, "participant_id": None, "collected_at": None, "values": [],
             "errors": [f"PDF unreadable: {fehler}"], "error_parameters": [], "warnings": []}
    else:
        if e["sha256"] in gesehen:
            e["warnings"].append(f"Same file as {gesehen[e['sha256']]}.")
        else:
            gesehen[e["sha256"]] = e["file"]
    ergebnisse.append(e)
    return json.dumps(e, ensure_ascii=False)

def excel(blaetter):
    backend.EXCEL_BLAETTER = json.loads(blaetter)
    backend.excel_schreiben(ergebnisse, "/tmp/blutbilder.xlsx")
    # the preview is read back from the written file, so it shows exactly what the Excel holds
    from openpyxl import load_workbook
    wb = load_workbook("/tmp/blutbilder.xlsx", read_only=True)
    return json.dumps({ws.title: [list(r) for r in ws.iter_rows(values_only=True)] for ws in wb},
                      ensure_ascii=False, default=str)

def reset():
    ergebnisse.clear(); gesehen.clear()
`;

let ready = null;
let py = null;
let counter = 0;

async function boot() {
  py = await loadPyodide({ indexURL: new URL("vendor/pyodide/", self.location).href });
  await py.loadPackage(WHEELS.map((w) => new URL("vendor/wheels/" + w, self.location).href));
  const src = await (await fetch("backend.py", { cache: "no-cache" })).text();
  // shortcut: drops the two debug lines in backend.py that read a fixed path on import; remove when they go.
  const patched = src.replace(/^ergebnis = lies_pdf\(EINGABEN\[0\]\)\s*\npprint\(ergebnis.*\)\s*$/m, "");
  py.FS.mkdirTree("/home/pyodide");
  py.FS.writeFile("/home/pyodide/backend.py", patched);
  py.runPython(GLUE);
  // The page's CSP does not reach this worker (GitHub Pages cannot send headers), so once everything is loaded
  // the worker gives up its network APIs: from here on nothing in it can send data anywhere.
  for (const api of ["fetch", "XMLHttpRequest", "WebSocket", "EventSource", "importScripts", "Worker"]) {
    Object.defineProperty(self, api, { value: undefined, writable: false, configurable: false });
  }
}

// File names come from the user's disk; keep them a single path segment inside MEMFS.
const safeName = (name) => name.replace(/[\\/\0]/g, "_").replace(/^\.+$/, "_") || "unnamed.pdf";

self.onmessage = async ({ data }) => {
  try {
    ready ??= boot();
    await ready;
    if (data.type === "run") {
      const n = data.files.length;
      for (let i = 0; i < n; i++) {
        self.postMessage({ type: "progress", i: i + 1, n });
        const name = safeName(data.files[i].name);
        let result;
        try {
          const dir = `/tmp/in/${counter++}`;
          py.FS.mkdirTree(dir);
          const path = `${dir}/${name}`;
          py.FS.writeFile(path, new Uint8Array(data.files[i].bytes));
          result = JSON.parse(py.globals.get("lies")(path));
          py.FS.unlink(path);
        } catch (err) {
          // one bad file must not stop the rest of the batch
          result = { file: name, participant_id: null, collected_at: null, values: [],
                     errors: [`PDF unreadable: ${err.message || err}`], error_parameters: [], warnings: [] };
        }
        self.postMessage({ type: "result", result });
      }
      self.postMessage({ type: "done" });
    } else if (data.type === "excel") {
      const sheets = JSON.parse(py.globals.get("excel")(JSON.stringify(data.sheets)));
      const bytes = py.FS.readFile("/tmp/blutbilder.xlsx");
      self.postMessage({ type: "excel", bytes, sheets }, [bytes.buffer]);
    } else if (data.type === "reset") {
      py.globals.get("reset")();
    } else if (data.type === "warmup") {
      self.postMessage({ type: "ready" });
    }
  } catch (err) {
    self.postMessage({ type: "error", message: String(err.message || err) });
  }
};
