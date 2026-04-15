# Dashboard AI Worx

A dark-mode desktop dashboard that discovers, configures and launches local
tools. Built with Python + PyWebView and a single-file HTML/CSS/JS frontend.

The app syncs all tools from
[`nicksSDWorx/Dashboard-AI-Worx`](https://github.com/nicksSDWorx/Dashboard-AI-Worx)
into a local `repo/` folder, groups them by first-level subdirectory, and
launches them through a configurable per-extension handler table.

## File layout

```
dashboard/
├── dashboard.py      # Python backend + PyWebView launcher
├── ui.html           # All HTML, CSS and JS in one file
├── settings.json     # Auto-created on first run (extension handlers)
└── repo/             # Created after the first GitHub sync
```

## Requirements

- **Python 3.10+** on Windows (the app also runs on macOS/Linux for development)
- **PyWebView** — the only external dependency:

  ```powershell
  pip install pywebview
  ```

  On Windows, PyWebView uses the system WebView2 runtime. It ships with
  Windows 11 and recent Windows 10 builds; otherwise install it once from
  <https://developer.microsoft.com/microsoft-edge/webview2/>.

All other functionality (settings, GitHub sync, process launching) uses only
the Python standard library (`urllib`, `subprocess`, `threading`, `json`,
`os`, `shutil`).

## Running from source

Double-click **`run.bat`**, or from a terminal:

```powershell
cd dashboard
python dashboard.py
```

> Do **not** open `ui.html` in a browser directly — it has no access to
> the Python backend and therefore no tools, sync or launching. The app
> must be started through `dashboard.py` (or the built exe) so that
> PyWebView injects its `window.pywebview.api` bridge.

On first launch:

1. `settings.json` is created with sensible defaults.
2. The sidebar shows **Dashboard**, **GitHub** and **Settings**.
3. The Dashboard tab shows an empty state — click **Sync repo** to download
   all tools from GitHub.
4. After the sync completes, tools appear as cards grouped by category.
5. Click a card to launch the underlying tool.

## Configuring handlers

Open the **Settings** tab to see all configured extension handlers. You can
add, edit or delete them. Three special launcher values are recognised:

| launcher  | behaviour                                                           |
| --------- | ------------------------------------------------------------------- |
| `python`  | Uses the same interpreter that runs the dashboard (`sys.executable`) |
| `direct`  | Runs the file itself (e.g. `.exe`, `.bat`)                          |
| *other*   | Treated as an executable name/path (e.g. `node`, `powershell`)      |

Each handler may also define a list of extra `args` prepended before the
script path. Changes are persisted immediately to `settings.json`.

## Changing the branch

Edit the `BRANCH` constant near the top of `dashboard.py`:

```python
BRANCH = "main"
```

## Building a standalone `.exe`

The quick way — double-click **`build.bat`**. It installs `pywebview` and
`pyinstaller` if needed, then produces `dist\DashboardAIWorx.exe`.

Manual equivalent:

```powershell
pip install pywebview pyinstaller
pyinstaller ^
  --name DashboardAIWorx ^
  --noconsole ^
  --onefile ^
  --add-data "ui.html;." ^
  dashboard.py
```

The resulting binary lives in `dist\DashboardAIWorx.exe`. It is fully
self-contained — drop it anywhere. `settings.json` and `repo/` are created
**beside the exe** (not inside the temp bundle), so your configuration and
synced tools survive between runs.

> Note: on macOS/Linux replace `;` with `:` in the `--add-data` argument.

## Troubleshooting

- **Nothing happens when clicking a card** — open **Settings** and make sure
  the file extension has a handler configured.
- **Sync fails with HTTP 403** — GitHub rate limits unauthenticated API
  requests to 60/hour per IP. Wait an hour or set up a personal access token.
- **WebView2 missing** — install the Evergreen runtime from Microsoft.
