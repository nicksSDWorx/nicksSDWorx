# PDF Merger — Windows executable

A local, no-install PDF merger for Windows. Double-click `pdf_merger.exe` and
your default browser opens with the merger UI. All processing happens on your
machine — nothing is uploaded anywhere.

## Usage

1. Double-click `pdf_merger.exe`.
2. A small console window appears and your default browser opens automatically.
   If it doesn't, copy the `http://127.0.0.1:PORT/` URL from the console window
   into your browser.
3. Drag PDF files into the drop zone (or click to browse), reorder them, and
   click **Merge PDFs**. The merged file downloads as `merged.pdf`.
4. To stop the app, close the console window.

The first time you run it, Windows SmartScreen may warn because the executable
is unsigned. Click **More info → Run anyway**.

## How it works

`pdf_merger.exe` is a tiny Go program that:

- Embeds `pdf_merger.html` directly into the binary (via `//go:embed`).
- Starts an HTTP server bound to `127.0.0.1` on a random free port.
- Opens your default browser to that local URL.

The HTML page uses [`pdf-lib`](https://pdf-lib.js.org/) for merging and
[`pdf.js`](https://mozilla.github.io/pdf.js/) for page-count display, loaded
from a CDN. An internet connection is required on first load so the browser
can fetch those two libraries.

## Building from source

Requires Go 1.24+.

```bash
# Windows x64 executable
GOOS=windows GOARCH=amd64 CGO_ENABLED=0 go build -ldflags="-s -w" -o pdf_merger.exe .

# Or for the current platform
go build -o pdf_merger .
```
