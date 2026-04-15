"""Dashboard-AI-Worx — desktop dashboard hub for launching local tools.

Pure-stdlib implementation: runs a tiny localhost HTTP server that
serves the UI and exposes a JSON-RPC-style API. A chromeless Edge /
Chrome window is spawned in --app mode to host the UI, giving a
native-desktop feel without requiring pywebview/pythonnet (and
therefore compatible with every Python version, including 3.14).
"""

import json
import os
import shutil
import socket
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.request
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

# ---------------------------------------------------------------------------
# Paths (PyInstaller-aware)
# ---------------------------------------------------------------------------

if getattr(sys, "frozen", False):
    # --onefile: exe folder holds persistent data; _MEIPASS holds bundled ui.html
    APP_DIR = Path(sys.executable).resolve().parent
    BUNDLE_DIR = Path(getattr(sys, "_MEIPASS", APP_DIR))
else:
    APP_DIR = Path(__file__).resolve().parent
    BUNDLE_DIR = APP_DIR

REPO_DIR = APP_DIR / "repo"
SETTINGS_PATH = APP_DIR / "settings.json"
UI_PATH = BUNDLE_DIR / "ui.html"

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

GITHUB_OWNER = "nicksSDWorx"
GITHUB_REPO = "Dashboard-AI-Worx"
BRANCH = "main"

TOOL_EXTENSIONS = {".py", ".exe", ".bat", ".ps1", ".js", ".sh"}

IGNORE_FILES = {"__init__.py", "setup.py", "requirements.txt", "utils.py", "helpers.py"}
IGNORE_DIRS = {".git", "__pycache__", "venv", ".venv", "node_modules", ".github", "tests", "docs"}

ICON_CYCLE = [
    "\u25C6", "\u25B2", "\u25CF", "\u2726", "\u2756",
    "\u2734", "\u25A0", "\u25D0", "\u2605", "\u2B22",
]

DEFAULT_SETTINGS = {
    "handlers": {
        ".py":  {"launcher": "python",     "args": [],                                        "label": "Python script"},
        ".exe": {"launcher": "direct",     "args": [],                                        "label": "Executable"},
        ".bat": {"launcher": "direct",     "args": [],                                        "label": "Batch file"},
        ".ps1": {"launcher": "powershell", "args": ["-ExecutionPolicy", "Bypass", "-File"],  "label": "PowerShell"},
        ".js":  {"launcher": "node",       "args": [],                                        "label": "Node.js script"},
    }
}

CATEGORY_FALLBACK = "Algemeen"


# ---------------------------------------------------------------------------
# Settings persistence
# ---------------------------------------------------------------------------

def load_settings() -> dict:
    if not SETTINGS_PATH.exists():
        save_settings(DEFAULT_SETTINGS)
        return json.loads(json.dumps(DEFAULT_SETTINGS))
    try:
        with open(SETTINGS_PATH, "r", encoding="utf-8") as fh:
            data = json.load(fh)
    except (OSError, json.JSONDecodeError):
        save_settings(DEFAULT_SETTINGS)
        return json.loads(json.dumps(DEFAULT_SETTINGS))
    if "handlers" not in data or not isinstance(data["handlers"], dict):
        data["handlers"] = dict(DEFAULT_SETTINGS["handlers"])
        save_settings(data)
    return data


def save_settings(data: dict) -> None:
    with open(SETTINGS_PATH, "w", encoding="utf-8") as fh:
        json.dump(data, fh, indent=2, ensure_ascii=False)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def snake_to_title(name: str) -> str:
    cleaned = name.replace("-", "_").replace(" ", "_")
    parts = [p for p in cleaned.split("_") if p]
    return " ".join(p.capitalize() for p in parts) if parts else name


def read_py_description(path: Path) -> str:
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as fh:
            text = fh.read(4096)
    except OSError:
        return ""
    stripped = text.lstrip()
    if stripped.startswith("#!"):
        newline = stripped.find("\n")
        if newline != -1:
            stripped = stripped[newline + 1 :].lstrip()
    for quote in ('"""', "'''"):
        if stripped.startswith(quote):
            end = stripped.find(quote, len(quote))
            if end != -1:
                block = stripped[len(quote):end].strip()
                first = next((ln.strip() for ln in block.splitlines() if ln.strip()), "")
                return first
    for line in stripped.splitlines():
        line = line.strip()
        if line.startswith("#"):
            return line.lstrip("#").strip()
        if line:
            break
    return ""


def read_readme_description(folder: Path) -> str:
    for candidate in ("README.md", "readme.md", "Readme.md", "README.MD"):
        path = folder / candidate
        if path.exists():
            try:
                with open(path, "r", encoding="utf-8", errors="replace") as fh:
                    for raw in fh:
                        line = raw.strip().lstrip("#").strip()
                        if line:
                            return line
            except OSError:
                pass
    return ""


# ---------------------------------------------------------------------------
# Tool discovery
# ---------------------------------------------------------------------------

def discover_tools() -> dict:
    result = {"categories": [], "empty": True, "repo_path": str(REPO_DIR)}
    if not REPO_DIR.exists():
        return result

    result["empty"] = False
    icon_idx = 0
    groups: dict[str, dict] = {}

    def add_tool(category_key: str, category_label: str, category_desc: str, file_path: Path):
        nonlocal icon_idx
        if category_key not in groups:
            groups[category_key] = {
                "key": category_key,
                "name": category_label,
                "description": category_desc,
                "tools": [],
            }
        ext = file_path.suffix.lower()
        tool_name = snake_to_title(file_path.stem)
        description = read_py_description(file_path) if ext == ".py" else ""
        if not description:
            description = f"{ext.lstrip('.').upper()} tool"
        icon = ICON_CYCLE[icon_idx % len(ICON_CYCLE)]
        icon_idx += 1
        rel_path = file_path.relative_to(REPO_DIR).as_posix()
        groups[category_key]["tools"].append({
            "name": tool_name,
            "description": description,
            "extension": ext,
            "icon": icon,
            "path": str(file_path),
            "rel_path": rel_path,
            "category": category_label,
        })

    for entry in sorted(REPO_DIR.iterdir(), key=lambda p: p.name.lower()):
        if entry.is_file() and entry.suffix.lower() in TOOL_EXTENSIONS and entry.name not in IGNORE_FILES:
            add_tool(CATEGORY_FALLBACK, CATEGORY_FALLBACK, "Losse tools uit de root van de repo.", entry)

    for sub in sorted(REPO_DIR.iterdir(), key=lambda p: p.name.lower()):
        if not sub.is_dir() or sub.name in IGNORE_DIRS:
            continue
        category_label = snake_to_title(sub.name)
        category_desc = read_readme_description(sub) or f"Tools in {category_label.lower()}."
        for path in sorted(sub.rglob("*"), key=lambda p: p.as_posix().lower()):
            if not path.is_file():
                continue
            if any(part in IGNORE_DIRS for part in path.relative_to(sub).parts[:-1]):
                continue
            if path.name in IGNORE_FILES:
                continue
            if path.suffix.lower() not in TOOL_EXTENSIONS:
                continue
            add_tool(sub.name, category_label, category_desc, path)

    ordered = []
    if CATEGORY_FALLBACK in groups:
        ordered.append(groups.pop(CATEGORY_FALLBACK))
    for key in sorted(groups.keys(), key=str.lower):
        ordered.append(groups[key])

    result["categories"] = [g for g in ordered if g["tools"]]
    return result


# ---------------------------------------------------------------------------
# GitHub sync
# ---------------------------------------------------------------------------

class GitHubSync:
    def __init__(self):
        self._thread: threading.Thread | None = None
        self._lock = threading.Lock()
        self._log: list[str] = []
        self._running = False
        self._done = False
        self._success = False
        self._started_at = 0.0

    def _append(self, line: str) -> None:
        with self._lock:
            self._log.append(line)

    def status(self) -> dict:
        with self._lock:
            return {
                "running": self._running,
                "done": self._done,
                "success": self._success,
                "log": list(self._log),
            }

    def start(self) -> dict:
        with self._lock:
            if self._running:
                return {"ok": False, "error": "Sync is al bezig."}
            self._log = []
            self._running = True
            self._done = False
            self._success = False
            self._started_at = time.time()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()
        return {"ok": True}

    def _http_get(self, url: str, accept: str | None = None) -> bytes:
        req = urllib.request.Request(url, headers={
            "User-Agent": "Dashboard-AI-Worx/1.0",
            "Accept": accept or "application/vnd.github+json",
        })
        with urllib.request.urlopen(req, timeout=30) as resp:
            return resp.read()

    def _list_tree(self) -> list[dict]:
        url = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/git/trees/{BRANCH}?recursive=1"
        self._append(f"GET {url}")
        raw = self._http_get(url)
        data = json.loads(raw.decode("utf-8"))
        if data.get("truncated"):
            self._append("! Waarschuwing: tree is afgekapt door GitHub (>100k entries).")
        return data.get("tree", []) or []

    def _download_file(self, rel_path: str, dest: Path) -> None:
        url = f"https://raw.githubusercontent.com/{GITHUB_OWNER}/{GITHUB_REPO}/{BRANCH}/{rel_path}"
        raw = self._http_get(url, accept="*/*")
        dest.parent.mkdir(parents=True, exist_ok=True)
        with open(dest, "wb") as fh:
            fh.write(raw)

    def _run(self) -> None:
        ok = False
        try:
            self._append(f"Start sync {GITHUB_OWNER}/{GITHUB_REPO}@{BRANCH}")
            if REPO_DIR.exists():
                self._append("Oude repo/ map opruimen...")
                shutil.rmtree(REPO_DIR, ignore_errors=True)
            REPO_DIR.mkdir(parents=True, exist_ok=True)

            entries = self._list_tree()
            files = [e for e in entries if e.get("type") == "blob"]
            self._append(f"{len(files)} bestanden gevonden in repo.")

            downloaded = 0
            for idx, entry in enumerate(files, 1):
                rel = entry.get("path", "")
                if not rel:
                    continue
                parts = rel.split("/")
                if any(p in IGNORE_DIRS for p in parts[:-1]):
                    continue
                dest = REPO_DIR / rel
                try:
                    self._download_file(rel, dest)
                    downloaded += 1
                    if downloaded % 5 == 0 or idx == len(files):
                        self._append(f"[{idx}/{len(files)}] {rel}")
                except urllib.error.HTTPError as exc:
                    self._append(f"! HTTP {exc.code} bij {rel}")
                except urllib.error.URLError as exc:
                    self._append(f"! Netwerkfout bij {rel}: {exc.reason}")

            elapsed = time.time() - self._started_at
            self._append(f"Klaar: {downloaded} bestanden gedownload in {elapsed:.1f}s.")
            ok = True
        except urllib.error.HTTPError as exc:
            self._append(f"FOUT: HTTP {exc.code} — {exc.reason}")
        except urllib.error.URLError as exc:
            self._append(f"FOUT: kan GitHub niet bereiken ({exc.reason}).")
        except Exception as exc:  # noqa: BLE001
            self._append(f"FOUT: {exc}")
        finally:
            with self._lock:
                self._running = False
                self._done = True
                self._success = ok


# ---------------------------------------------------------------------------
# API (same surface as the old PyWebView js_api)
# ---------------------------------------------------------------------------

class Api:
    def __init__(self):
        self.settings = load_settings()
        self.sync = GitHubSync()

    def get_tools(self) -> dict:
        return discover_tools()

    def get_repo_info(self) -> dict:
        return {
            "owner": GITHUB_OWNER,
            "repo": GITHUB_REPO,
            "branch": BRANCH,
            "url": f"https://github.com/{GITHUB_OWNER}/{GITHUB_REPO}",
            "repo_exists": REPO_DIR.exists(),
            "repo_path": str(REPO_DIR),
        }

    def get_settings(self) -> dict:
        return self.settings

    def add_handler(self, extension: str, launcher: str, args, label: str) -> dict:
        ext = (extension or "").strip().lower()
        if not ext:
            return {"ok": False, "error": "Extensie ontbreekt."}
        if not ext.startswith("."):
            ext = "." + ext
        if not (launcher or "").strip():
            return {"ok": False, "error": "Launcher ontbreekt."}
        if isinstance(args, str):
            args_list = [a for a in args.split() if a]
        elif isinstance(args, list):
            args_list = [str(a) for a in args]
        else:
            args_list = []
        self.settings.setdefault("handlers", {})[ext] = {
            "launcher": launcher.strip(),
            "args": args_list,
            "label": (label or "").strip() or ext.lstrip(".").upper(),
        }
        save_settings(self.settings)
        return {"ok": True, "handlers": self.settings["handlers"]}

    def delete_handler(self, extension: str) -> dict:
        ext = (extension or "").strip().lower()
        if ext in self.settings.get("handlers", {}):
            del self.settings["handlers"][ext]
            save_settings(self.settings)
            return {"ok": True, "handlers": self.settings["handlers"]}
        return {"ok": False, "error": f"Geen handler voor {ext}."}

    def sync_repo(self) -> dict:
        return self.sync.start()

    def get_sync_status(self) -> dict:
        return self.sync.status()

    def launch_tool(self, script_path: str) -> dict:
        if not script_path:
            return {"ok": False, "error": "Geen pad opgegeven."}
        path = Path(script_path)
        if not path.exists() or not path.is_file():
            return {"ok": False, "error": f"Bestand niet gevonden: {path}"}

        ext = path.suffix.lower()
        handler = self.settings.get("handlers", {}).get(ext)
        if not handler:
            return {
                "ok": False,
                "error": (
                    f"Geen handler geconfigureerd voor '{ext}'. "
                    "Voeg er een toe in de Settings tab."
                ),
            }

        launcher = (handler.get("launcher") or "").strip()
        args = list(handler.get("args") or [])

        if launcher == "python":
            cmd = [sys.executable, *args, str(path)]
        elif launcher == "direct":
            cmd = [str(path), *args]
        else:
            cmd = [launcher, *args, str(path)]

        try:
            popen_kwargs = {"cwd": str(path.parent)}
            if os.name == "nt":
                popen_kwargs["creationflags"] = getattr(subprocess, "CREATE_NEW_CONSOLE", 0)
            subprocess.Popen(cmd, **popen_kwargs)
        except FileNotFoundError:
            return {"ok": False, "error": f"Launcher '{launcher}' niet gevonden op PATH."}
        except OSError as exc:
            return {"ok": False, "error": f"Kan tool niet starten: {exc}"}

        return {"ok": True, "cmd": " ".join(cmd)}

    # Called by the HTTP server when the UI navigates away / closes.
    def shutdown(self) -> dict:
        threading.Thread(target=lambda: (time.sleep(0.2), os._exit(0)), daemon=True).start()
        return {"ok": True}


# ---------------------------------------------------------------------------
# HTTP server (serves ui.html + /api/<method> JSON-RPC)
# ---------------------------------------------------------------------------

def build_handler(api: Api, ui_html: bytes, auth_token: str):
    """Build a BaseHTTPRequestHandler subclass bound to the given api."""

    class Handler(BaseHTTPRequestHandler):
        # Silence the default stderr access log.
        def log_message(self, format, *args):  # noqa: A002
            return

        def _write_json(self, status: int, payload) -> None:
            body = json.dumps(payload).encode("utf-8")
            self.send_response(status)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(body)

        def _check_token(self) -> bool:
            # Require the token in either the query string or the X-Auth header.
            # This stops drive-by LAN access if the port gets scanned.
            if "token=" + auth_token in (self.path or ""):
                return True
            if self.headers.get("X-Auth") == auth_token:
                return True
            return False

        def do_GET(self):
            path = self.path.split("?", 1)[0]
            if path in ("/", "/index.html"):
                if not self._check_token():
                    self.send_error(403, "Forbidden")
                    return
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(ui_html)))
                self.send_header("Cache-Control", "no-store")
                self.end_headers()
                self.wfile.write(ui_html)
                return
            self.send_error(404, "Not Found")

        def do_POST(self):
            if not self.path.startswith("/api/"):
                self.send_error(404, "Not Found")
                return
            if not self._check_token():
                self.send_error(403, "Forbidden")
                return

            method_name = self.path[len("/api/"):].split("?", 1)[0]
            if not method_name or method_name.startswith("_"):
                self.send_error(404, "Not Found")
                return
            method = getattr(api, method_name, None)
            if not callable(method):
                self.send_error(404, f"Unknown method: {method_name}")
                return

            length = int(self.headers.get("Content-Length", "0") or "0")
            raw = self.rfile.read(length) if length else b"[]"
            try:
                args = json.loads(raw.decode("utf-8") or "[]")
            except json.JSONDecodeError:
                self._write_json(400, {"error": "Invalid JSON body."})
                return
            if not isinstance(args, list):
                self._write_json(400, {"error": "Body must be a JSON array of arguments."})
                return

            try:
                result = method(*args)
            except TypeError as exc:
                self._write_json(400, {"error": f"Bad arguments: {exc}"})
                return
            except Exception as exc:  # noqa: BLE001
                self._write_json(500, {"error": str(exc)})
                return

            self._write_json(200, result)

    return Handler


def pick_free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


# ---------------------------------------------------------------------------
# Browser launcher (chromeless --app mode)
# ---------------------------------------------------------------------------

def find_chromium_browser() -> str | None:
    """Locate Microsoft Edge or Chrome on the system, Windows-first."""
    if os.name == "nt":
        candidates = [
            r"%ProgramFiles%\Microsoft\Edge\Application\msedge.exe",
            r"%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe",
            r"%LOCALAPPDATA%\Microsoft\Edge\Application\msedge.exe",
            r"%ProgramFiles%\Google\Chrome\Application\chrome.exe",
            r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe",
            r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe",
        ]
        for c in candidates:
            expanded = os.path.expandvars(c)
            if os.path.isfile(expanded):
                return expanded
    # Fallbacks on PATH (also works on macOS/Linux for dev).
    for name in ("msedge", "microsoft-edge", "google-chrome", "chrome", "chromium", "chromium-browser"):
        resolved = shutil.which(name)
        if resolved:
            return resolved
    return None


def launch_browser(url: str) -> subprocess.Popen | None:
    browser = find_chromium_browser()
    profile_dir = APP_DIR / ".browser_profile"
    profile_dir.mkdir(exist_ok=True)
    if browser:
        args = [
            browser,
            f"--app={url}",
            f"--user-data-dir={profile_dir}",
            "--window-size=1280,820",
            "--no-first-run",
            "--no-default-browser-check",
            "--disable-features=TranslateUI",
        ]
        try:
            return subprocess.Popen(args)
        except OSError as exc:
            print(f"[dashboard] Failed to launch {browser}: {exc}", file=sys.stderr)
            return None
    # Last resort: open in the default browser. UX is worse (tab, URL bar)
    # but at least functional.
    import webbrowser
    print(f"[dashboard] No Edge/Chrome found; opening {url} in default browser.")
    webbrowser.open(url)
    return None


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    if not UI_PATH.exists():
        raise SystemExit(f"ui.html ontbreekt (verwacht: {UI_PATH})")

    load_settings()  # ensure settings.json exists
    with open(UI_PATH, "rb") as fh:
        ui_html = fh.read()

    api = Api()
    auth_token = os.urandom(16).hex()
    port = pick_free_port()
    server = ThreadingHTTPServer(("127.0.0.1", port), build_handler(api, ui_html, auth_token))

    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()

    url = f"http://127.0.0.1:{port}/?token={auth_token}"
    proc = launch_browser(url)

    try:
        if proc is not None:
            # Block until the app window is closed.
            proc.wait()
        else:
            # Default browser fallback — keep the server alive until Ctrl+C.
            print("Druk op Ctrl+C om af te sluiten.")
            while True:
                time.sleep(1.0)
    except KeyboardInterrupt:
        pass
    finally:
        server.shutdown()
        server.server_close()


if __name__ == "__main__":
    main()
