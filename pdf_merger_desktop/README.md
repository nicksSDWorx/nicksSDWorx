# PDF Merger — Offline Windows desktop app

A native Windows application that merges multiple PDFs. 100% offline — no
internet required, no browser launches, no files leave your machine.

## Usage

1. Double-click `pdf_merger.exe`.
2. A native desktop window opens.
3. Click **Add PDFs…** (or drag files in, if tkinterdnd2 is available).
4. Reorder with **Move Up / Move Down**, remove files, or **Clear All**.
5. Click **Merge PDFs** → choose where to save → done.

On first run the launcher extracts a bundled Python 3.11 runtime + `pypdf` to
`%LOCALAPPDATA%\PdfMerger\<hash>\`. Subsequent runs start instantly.
When a new version of the `.exe` is distributed, it extracts into a new
folder and cleans up older ones automatically.

The first launch may trigger a Windows SmartScreen warning because the
executable is unsigned — click **More info → Run anyway**.

## How it works

`pdf_merger.exe` is a small Go wrapper that embeds a `bundle.zip` containing
everything needed to run the app offline:

- `python.exe` / `pythonw.exe` + `python311.dll` (conda-forge Windows build)
- MSVC runtime DLLs (`vcruntime140.dll`, `vcruntime140_1.dll`) so the
  target machine doesn't need a VC++ Redistributable installed
- Runtime DLLs for Python's stdlib (`zlib.dll`, `libcrypto-3-x64.dll`,
  `libssl-3-x64.dll`, `ffi-8.dll`, `libbz2.dll`, `libexpat.dll`,
  `liblzma.dll`, `sqlite3.dll`)
- Python standard library (trimmed: idlelib / test / distutils removed)
- Tcl/Tk 8.6 runtime for the Tkinter GUI
- `pypdf` 6.10 (pure Python; uses its built-in crypto fallback so no
  `cryptography` / `pycryptodome` DLLs are required)
- `pdf_merger.py` — the Tkinter app

On launch the wrapper:

1. Computes a SHA-256 hash of the embedded zip.
2. Ensures `%LOCALAPPDATA%\PdfMerger\<hash>\` exists and contains a
   `.ready` marker; if not, extracts the zip there.
3. Spawns `pythonw.exe pdf_merger.py` with `PYTHONHOME`, `PYTHONPATH`,
   `TCL_LIBRARY`, `TK_LIBRARY`, and a scrubbed `PATH` that points only at
   the install directory — so the embedded Python can't pick up any
   system-installed Python or mismatched DLLs.
4. Redirects stdout/stderr to `launch.log` inside the install dir and
   waits 1.5 s. If Python exits within that window (missing DLL, wrong
   version, etc.), the launcher pops up a Windows MessageBox showing the
   tail of the log. Otherwise it detaches — the app runs as a standalone
   Tkinter window with no console.

## Troubleshooting

If `pdf_merger.exe` refuses to start, check
`%LOCALAPPDATA%\PdfMerger\<hash>\launch.log` for the Python stderr output.
Deleting the entire `%LOCALAPPDATA%\PdfMerger\` folder forces a clean
re-extraction on the next launch.

## Building from source

Requires Go 1.24+ and (for rebuilding the bundle) access to conda-forge /
pypi.

```bash
# 1. Build the bundle (Windows Python + tkinter + pypdf + app)
./build-bundle.sh   # (see project history for the exact script)

# 2. Cross-compile the Windows launcher from any platform
cd launcher
GOOS=windows GOARCH=amd64 CGO_ENABLED=0 \
    go build -ldflags="-s -w -H=windowsgui" -o pdf_merger.exe .
```

## File layout

```
pdf_merger_desktop/
├── pdf_merger.exe          ← the standalone Windows app (12 MB)
├── pdf_merger.py           ← Tkinter/pypdf source (runs on any Python 3.9+)
├── launcher/
│   ├── main.go             ← extraction + launch logic
│   ├── platform_windows.go ← Windows-only helpers (hide console, MessageBox)
│   ├── platform_other.go   ← stubs for other platforms (build-only)
│   ├── bundle.zip          ← embedded by main.go via //go:embed
│   └── go.mod
└── README.md
```
