# Dashboard AI Worx

A dark-mode desktop dashboard that discovers, configures and launches local
tools. The frontend is a single HTML/CSS/JS file, the backend is a pure
**stdlib-only** Python program — no `pywebview`, no `pythonnet`, no
native extensions. Works on every Python version from 3.10 onwards
(including **3.14**).

The app syncs all tools from
[`nicksSDWorx/Dashboard-AI-Worx`](https://github.com/nicksSDWorx/Dashboard-AI-Worx)
into a local `repo/` folder, groups them by first-level subdirectory, and
launches them through a configurable per-extension handler table.

Each first-level subdirectory is shown as **one** tile: the dashboard
walks the folder recursively (skipping `IGNORE_DIRS` plus `_internal/`
so PyInstaller bundles don't flood the UI with helper EXEs) and picks
the single entry file:

1. a file whose stem matches the folder name (e.g.
   `BrokenURLFinder/dist/.../Broken_URL_Finder.exe`);
2. a file named `main.*`, `app.*`, `run.*`, `start.*`, `launch.*`, or
   `index.*`;
3. otherwise, the lone executable in the whole subtree, preferring a
   single `.exe` if multiple file types coexist.

Supporting files (READMEs, data, helpers, assets) still sync to disk
so the tool can use them but stay invisible in the UI. Folders with
no detectable entry file don't appear at all.

## How it works

1. `dashboard.py` starts a tiny HTTP server on `127.0.0.1:<random port>`
   that serves `ui.html` and exposes the Python API as JSON-RPC endpoints
   (`POST /api/<method>` with a JSON array of arguments).
2. It then spawns **Microsoft Edge** (or Chrome, if Edge isn't installed)
   in `--app` mode with a dedicated `--user-data-dir`. This gives a
   chromeless, native-feeling window with no tabs / address bar, pointed
   at the local server.
3. Closing the window kills the Edge child process, which causes the
   Python entrypoint to shut down the server and exit.

A random per-run auth token is baked into the URL and required on every
API call, so LAN port-scanners can't talk to the backend.

## File layout

```
dashboard/
├── dashboard.py      # Python backend + HTTP server + browser launcher
├── ui.html           # All HTML, CSS and JS in one file
├── build.bat         # Build DashboardAIWorx.exe via PyInstaller
├── run.bat           # Run from source
├── settings.json     # Auto-created on first run (extension handlers)
└── repo/             # Created after the first GitHub sync
```

## Requirements

- **Python 3.10 or newer** (tested on 3.10–3.14). No pip packages needed
  at runtime — everything is stdlib.
- **Microsoft Edge** (bundled with Windows 10/11) or Google Chrome.
  Without either, the app falls back to opening the default browser in a
  regular tab.

## Running from source

Double-click **`run.bat`**, or from a terminal:

```powershell
cd dashboard
python dashboard.py
```

## Building a standalone `.exe`

Double-click **`build.bat`**. It installs PyInstaller (the only build-time
dependency) and produces `dist\DashboardAIWorx.exe`. Manual equivalent:

```powershell
pip install pyinstaller
pyinstaller ^
  --name DashboardAIWorx ^
  --noconsole ^
  --onefile ^
  --add-data "ui.html;." ^
  dashboard.py
```

Drop the exe anywhere. `settings.json` and `repo/` are created **beside
the exe** (not inside the PyInstaller temp bundle), so configuration and
synced tools survive between runs.

> On macOS/Linux replace `;` with `:` in the `--add-data` argument.

## Running tools in the dashboard

Clicking a tool card no longer pops up a separate console window.
Instead the dashboard captures stdout/stderr via a pipe and streams the
output into an embedded console panel that slides up from the bottom:

- Each launch gets its own tab. Launching another tool adds a tab — the
  existing tool keeps running.
- A pulsing orange dot indicates a running job; green means exit 0;
  red means a non-zero exit code.
- **Stop** sends SIGTERM (Windows: TerminateProcess) and, if the child
  ignores it, SIGKILL after a 2s grace.
- **×** on a tab stops the job (if still running) and closes the tab.
- **Sluit** hides the panel; running jobs keep running and polling
  resumes next time the panel is opened.
- **Open in venster** re-runs the current tab's command in a fresh
  console window (CREATE_NEW_CONSOLE on Windows). Useful for tools
  that need a real TTY (interactive prompts, colour output, Windows
  console APIs).

### Automatic fallback

If a tool exits with a non-zero exit code within 3 seconds and the
user didn't press Stop, the dashboard automatically relaunches it
in a new console window and shows a toast. This catches tools that
simply can't run headless — without forcing you to click anything.

> GUI applications (`.exe` that open their own window) still show their
> own window — that's controlled by the app itself, not by the launcher.
> What changed is that no *extra* empty console appears behind them.

## Configuring handlers

The **Settings** tab lists all configured extension handlers. Three
special launcher values are recognised:

| launcher  | behaviour                                                           |
| --------- | ------------------------------------------------------------------- |
| `python`  | Uses the same interpreter that runs the dashboard (`sys.executable`) |
| `direct`  | Runs the file itself (e.g. `.exe`, `.bat`)                          |
| *other*   | Treated as an executable name/path (e.g. `node`, `powershell`)      |

Each handler can include extra `args` prepended before the script path.
Changes persist immediately to `settings.json`.

## Changing the branch

Edit the `BRANCH` constant near the top of `dashboard.py`:

```python
BRANCH = "main"
```

## Uploading a folder to the repo

The **GitHub** tab has an *Upload naar repo* panel. It takes a local
folder and writes everything under it to `<target>/` on the configured
branch as a single commit (Git Data API: blobs → tree → commit → update
ref). Flow:

1. Provide a Personal Access Token with the `repo` scope. Either:
   - paste it into **Settings → GitHub-authenticatie** (stored in
     `settings.json`, which is in `.gitignore`), *or*
   - export it as the `GITHUB_TOKEN` environment variable before
     launching (Windows: `setx GITHUB_TOKEN <pat>`, open a new shell).
     Settings wins over the environment variable.
2. From the **Dashboard** tab click **Upload** (or open **GitHub** and
   click **Map kiezen**) and pick a local folder. The status line shows
   the file count and total size.
3. Adjust **Doelmap in repo** and **Commit-bericht** if needed. The
   target folder defaults to the local folder's name.
4. Click **Push nu**. Progress streams into the log. On success the
   dashboard automatically runs **Sync nu** so the new folder appears as
   a tile.

## Troubleshooting

- **Window opens as a normal Chrome/Edge tab instead of chromeless** —
  Edge/Chrome couldn't be located. Install Edge from
  <https://www.microsoft.com/edge> or Chrome from <https://google.com/chrome>.
- **Nothing happens when clicking a card** — open **Settings** and make
  sure the file extension has a handler configured.
- **Sync fails with HTTP 403** — GitHub rate-limits unauthenticated API
  requests to 60/hour per IP. Wait an hour or add a personal access
  token to `dashboard.py`.
- **Sync fails with `CERTIFICATE_VERIFY_FAILED` / `unable to get local
  issuer certificate`** — you're behind a corporate SSL-inspecting
  proxy. The app already auto-imports every certificate from the
  Windows `ROOT` / `CA` stores on each sync, which fixes most cases.
  If it still fails, open **Settings** and untick *"SSL-certificaten
  verifiëren"*. This sets `"ssl_verify": false` in `settings.json` and
  skips HTTPS verification for the GitHub sync.
