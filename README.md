# Hema2Data

Turns blood count PDFs into a checked Excel file, entirely in the browser. `backend.py` runs unchanged in
[Pyodide](https://pyodide.org); PDFs never leave the computer.

## Run locally

```bash
python3 -m http.server 8765
```

Open http://localhost:8765 (opening `index.html` directly from disk does not work: workers need a server).

## Publish (GitHub Pages)

1. Push this folder to a GitHub repository.
2. Settings → Pages → Source: "Deploy from a branch", branch `main`, folder `/`.
3. Share `https://<user>.github.io/<repo>/`.

After changing any file, raise `VERSION` in `sw.js` so installed apps pick up the update.

## Why nothing is uploaded

- All libraries are in `vendor/` (Pyodide 0.27.7, pdfplumber, pdfminer.six, openpyxl, Geist font); no CDN, no tracking.
- The Content-Security-Policy in `index.html` allows connections only to the page's own origin, which only serves
  static files.
- DevTools → Network shows no request while PDFs are read. After the first visit the app works offline.

## Notes

- `pypdfium2` and `cryptography` have no browser build and are stubbed in `worker.js`; `backend.py` needs neither.
  Encrypted PDFs therefore fail with a clear message.
- `worker.js` skips the two debug lines `ergebnis = lies_pdf(EINGABEN[0])` / `pprint(...)` in `backend.py`, which read
  a fixed local path on import.
