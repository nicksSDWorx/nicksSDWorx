"""Dashboard-AI-Worx — desktop dashboard hub for launching local tools.

Pure-stdlib implementation: runs a tiny localhost HTTP server that
serves the UI and exposes a JSON-RPC-style API. A chromeless Edge /
Chrome window is spawned in --app mode to host the UI, giving a
native-desktop feel without requiring pywebview/pythonnet (and
therefore compatible with every Python version, including 3.14).
"""

import base64
import json
import os
import re
import shutil
import socket
import ssl
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.request
from datetime import datetime, timedelta
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
TOOL_WINDOW_PATH = BUNDLE_DIR / "tool_window.html"

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

GITHUB_OWNER = "nicksSDWorx"
GITHUB_REPO = "Dashboard-AI-Worx"
BRANCH = "main"

TOOL_EXTENSIONS = {".py", ".exe", ".bat", ".ps1", ".js", ".sh"}

IGNORE_FILES = {"__init__.py", "setup.py", "requirements.txt", "utils.py", "helpers.py"}
IGNORE_DIRS = {".git", "__pycache__", "venv", ".venv", "node_modules", ".github", "tests", "docs"}

# Priority order for picking the single "entry" file inside a tool-folder.
# Files that don't match are treated as supporting material (docs, data,
# helpers) and stay invisible in the dashboard while still syncing to disk.
ENTRY_STEMS = ("main", "app", "run", "start", "launch", "index")

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
    },
    # Set to false if you sit behind a corporate SSL-inspecting proxy
    # whose root CA can't be found even after auto-importing the
    # Windows certificate store. Disables HTTPS verification entirely.
    "ssl_verify": True,
    # User-controlled tile grouping. ``order`` lists category names in
    # the order they should appear; ``assignments`` maps a tool's
    # rel_path → category name. The virtual "Ongecategoriseerd" bucket
    # is always rendered last and never stored here.
    "categories": {"order": [], "assignments": {}},
    # Scheduled tool runs. See Scheduler / compute_next_run for the
    # accepted shape per entry.
    "schedules": [],
    # Where scheduled-run output is written. ``path`` empty = use the
    # built-in ``schedule_runs/`` folder next to the dashboard. Set it
    # to an absolute path (e.g. a OneDrive sync folder or network
    # share) when you want a future automation to pick the files up
    # from elsewhere.
    "runs_storage": {
        "path": "",
        "max_per_schedule": 50,
    },
}

VALID_SCHEDULE_TYPES = {"once", "daily", "weekly", "interval"}
MAX_SCHEDULE_NAME = 80

UNCATEGORIZED = "Ongecategoriseerd"
MAX_CATEGORY_NAME = 64

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
    changed = False
    if "handlers" not in data or not isinstance(data["handlers"], dict):
        data["handlers"] = dict(DEFAULT_SETTINGS["handlers"])
        changed = True
    cats = data.get("categories")
    if not isinstance(cats, dict):
        cats = {}
        data["categories"] = cats
        changed = True
    if not isinstance(cats.get("order"), list):
        cats["order"] = []
        changed = True
    if not isinstance(cats.get("assignments"), dict):
        cats["assignments"] = {}
        changed = True
    if not isinstance(data.get("schedules"), list):
        data["schedules"] = []
        changed = True
    # Drop abandoned mail blocks from earlier iterations.
    for legacy in ("graph_mail", "smtp_mail"):
        if legacy in data:
            data.pop(legacy, None)
            changed = True
    # Strip dead mail fields off existing schedules so the file stays clean.
    for s in data.get("schedules") or []:
        if not isinstance(s, dict):
            continue
        for k in ("mail_trigger", "mail_to", "mail_status"):
            if k in s:
                s.pop(k, None)
                changed = True
    rs = data.get("runs_storage")
    if not isinstance(rs, dict):
        data["runs_storage"] = dict(DEFAULT_SETTINGS["runs_storage"])
        changed = True
    else:
        for k, v in DEFAULT_SETTINGS["runs_storage"].items():
            if k not in rs:
                rs[k] = v
                changed = True
    if changed:
        save_settings(data)
    return data


def save_settings(data: dict) -> None:
    with open(SETTINGS_PATH, "w", encoding="utf-8") as fh:
        json.dump(data, fh, indent=2, ensure_ascii=False)


# ---------------------------------------------------------------------------
# SSL context (corporate-proxy friendly)
# ---------------------------------------------------------------------------

def build_ssl_context(verify: bool = True) -> ssl.SSLContext:
    """Create an SSL context that trusts both the Python CA bundle AND
    every certificate in the Windows ROOT/CA stores.

    On corporate networks (SD Worx, banks, government, etc.) an HTTPS
    inspecting proxy signs traffic with a private root CA that IT
    pushed to every employee's Windows certificate store. Python's
    default ``load_default_certs()`` only imports certs whose trust
    flags include ``SERVER_AUTH``; some corporate roots get filtered
    out by that check, producing
    ``CERTIFICATE_VERIFY_FAILED: unable to get local issuer certificate``.

    We sidestep that by enumerating the stores manually and adding
    everything OpenSSL will accept — it still won't use a cert for the
    wrong purpose, because the X.509 EKU extension on each certificate
    is checked at handshake time.
    """
    if not verify:
        ctx = ssl._create_unverified_context()  # noqa: SLF001 — intentional
        return ctx

    ctx = ssl.create_default_context()
    if os.name == "nt":
        imported = 0
        for store in ("ROOT", "CA", "MY"):
            try:
                certs = ssl.enum_certificates(store)
            except (OSError, AttributeError):
                continue
            for cert_bytes, encoding, _trust in certs:
                if encoding != "x509_asn":
                    continue
                try:
                    ctx.load_verify_locations(cadata=cert_bytes)
                    imported += 1
                except ssl.SSLError:
                    # Not a CA cert, or duplicate — fine.
                    pass
        # Leave a breadcrumb; handy when diagnosing SSL issues.
        if imported:
            os.environ.setdefault("DASHBOARD_WINCERTS_IMPORTED", str(imported))
    return ctx


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


def find_entry_file(folder: Path) -> Path | None:
    """Pick the single executable file that represents a tool-folder.

    Walks the full subtree (skipping ``IGNORE_DIRS`` plus ``_internal``,
    which PyInstaller one-folder bundles fill with helper DLLs/EXEs we
    shouldn't surface) and picks, in order:

      1. The file whose stem matches the folder name (case/punctuation
         insensitive). Catches PyInstaller layouts like
         ``BrokenURLFinder/dist/<bundle>/Broken_URL_Finder.exe``.
      2. The file whose stem matches a known entry (main, app, run, ...).
      3. If exactly one executable exists in the whole subtree, use it.
      4. If exactly one ``.exe`` exists, use it (prefer a prebuilt binary
         over scattered scripts).
    """
    ignore = IGNORE_DIRS | {"_internal"}
    norm = lambda s: re.sub(r"[^a-z0-9]", "", s.lower())

    executables: list[Path] = []
    stack: list[Path] = [folder]
    while stack:
        current = stack.pop()
        try:
            children = list(current.iterdir())
        except OSError:
            continue
        for p in children:
            if p.is_dir():
                if p.name not in ignore:
                    stack.append(p)
            elif (
                p.is_file()
                and p.suffix.lower() in TOOL_EXTENSIONS
                and p.name not in IGNORE_FILES
            ):
                executables.append(p)

    if not executables:
        return None

    # Prefer shallower paths when multiple candidates tie.
    executables.sort(key=lambda p: (len(p.parts), p.as_posix().lower()))

    norm_folder = norm(folder.name)
    for p in executables:
        if norm(p.stem) == norm_folder:
            return p

    for stem in ENTRY_STEMS:
        for p in executables:
            if p.stem.lower() == stem:
                return p

    if len(executables) == 1:
        return executables[0]

    exe_only = [p for p in executables if p.suffix.lower() == ".exe"]
    if len(exe_only) == 1:
        return exe_only[0]

    return None


# ---------------------------------------------------------------------------
# Tool discovery
# ---------------------------------------------------------------------------

def discover_tools_flat() -> list[dict]:
    """Discover all tools as a flat list. Categorisation is applied
    separately (see :func:`group_by_user_categories`) so the user can
    override the repo's folder structure with their own grouping.
    """
    if not REPO_DIR.exists():
        return []

    tools: list[dict] = []
    icon_idx = 0

    def add(file_path: Path) -> None:
        nonlocal icon_idx
        ext = file_path.suffix.lower()
        tool_name = snake_to_title(file_path.stem)
        description = read_py_description(file_path) if ext == ".py" else ""
        if not description:
            description = f"{ext.lstrip('.').upper()} tool"
        icon = ICON_CYCLE[icon_idx % len(ICON_CYCLE)]
        icon_idx += 1
        rel_path = file_path.relative_to(REPO_DIR).as_posix()
        tools.append({
            "name": tool_name,
            "description": description,
            "extension": ext,
            "icon": icon,
            "path": str(file_path),
            "rel_path": rel_path,
        })

    for entry in sorted(REPO_DIR.iterdir(), key=lambda p: p.name.lower()):
        if (
            entry.is_file()
            and entry.suffix.lower() in TOOL_EXTENSIONS
            and entry.name not in IGNORE_FILES
        ):
            add(entry)

    for sub in sorted(REPO_DIR.iterdir(), key=lambda p: p.name.lower()):
        if not sub.is_dir() or sub.name in IGNORE_DIRS:
            continue
        entry = find_entry_file(sub)
        if entry is not None:
            add(entry)

    return tools


def group_by_user_categories(tools: list[dict], settings: dict) -> dict:
    """Build the ``{categories, empty, repo_path}`` payload the UI
    consumes, using the user's chosen ``order``/``assignments`` from
    settings. Tools without an assignment (or with an assignment that
    points to a deleted category) land in the virtual *Ongecategoriseerd*
    bucket which always renders last.
    """
    cats = (settings or {}).get("categories") or {}
    order = [c for c in (cats.get("order") or []) if isinstance(c, str)]
    assignments = cats.get("assignments") or {}
    if not isinstance(assignments, dict):
        assignments = {}

    known = set(order)
    groups: dict[str, list[dict]] = {name: [] for name in order}
    uncategorised: list[dict] = []

    for tool in tools:
        target = assignments.get(tool["rel_path"])
        if target and target in known:
            tool = dict(tool)
            tool["category"] = target
            groups[target].append(tool)
        else:
            tool = dict(tool)
            tool["category"] = UNCATEGORIZED
            uncategorised.append(tool)

    payload: list[dict] = []
    for name in order:
        payload.append({
            "key": name,
            "name": name,
            "description": "",
            "tools": groups[name],
            "user_managed": True,
        })
    payload.append({
        "key": UNCATEGORIZED,
        "name": UNCATEGORIZED,
        "description": "Tools die nog geen categorie hebben.",
        "tools": uncategorised,
        "user_managed": False,
    })

    return {
        "categories": payload,
        "empty": not tools,
        "repo_path": str(REPO_DIR),
    }


def discover_tools(settings: dict | None = None) -> dict:
    return group_by_user_categories(discover_tools_flat(), settings or {})


# ---------------------------------------------------------------------------
# Category management helpers
# ---------------------------------------------------------------------------

def _normalise_category_name(raw: str) -> tuple[str, str | None]:
    """Return ``(clean_name, error)``. ``error`` is None when valid."""
    name = (raw or "").strip()
    if not name:
        return "", "Categorienaam mag niet leeg zijn."
    if len(name) > MAX_CATEGORY_NAME:
        return "", f"Categorienaam mag max. {MAX_CATEGORY_NAME} tekens zijn."
    if name.lower() == UNCATEGORIZED.lower():
        return "", f"De naam '{UNCATEGORIZED}' is gereserveerd."
    return name, None


# ---------------------------------------------------------------------------
# GitHub sync
# ---------------------------------------------------------------------------

class GitHubSync:
    def __init__(self, settings_provider):
        # settings_provider is a zero-arg callable that returns the live
        # settings dict, so toggling ssl_verify in settings.json takes
        # effect without restarting the app.
        self._settings_provider = settings_provider
        self._thread: threading.Thread | None = None
        self._lock = threading.Lock()
        self._log: list[str] = []
        self._running = False
        self._done = False
        self._success = False
        self._started_at = 0.0
        self._ssl_ctx: ssl.SSLContext | None = None

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
        with urllib.request.urlopen(req, timeout=30, context=self._ssl_ctx) as resp:
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
            settings = self._settings_provider() or {}
            verify = bool(settings.get("ssl_verify", True))
            self._ssl_ctx = build_ssl_context(verify=verify)
            imported = os.environ.get("DASHBOARD_WINCERTS_IMPORTED")
            if verify and imported:
                self._append(f"SSL: {imported} certs uit Windows store geladen.")
            elif not verify:
                self._append("SSL: verificatie UIT (ssl_verify: false in settings.json).")

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
            reason = str(exc.reason)
            self._append(f"FOUT: kan GitHub niet bereiken ({reason}).")
            if "CERTIFICATE_VERIFY_FAILED" in reason or "SSL" in reason.upper():
                self._append(
                    "! Corporate SSL-proxy gedetecteerd. Zet in settings.json "
                    "\"ssl_verify\": false en probeer opnieuw."
                )
        except Exception as exc:  # noqa: BLE001
            self._append(f"FOUT: {exc}")
        finally:
            with self._lock:
                self._running = False
                self._done = True
                self._success = ok


# ---------------------------------------------------------------------------
# GitHub push (upload local folder as a single commit)
# ---------------------------------------------------------------------------

SAFE_FOLDER_RE = re.compile(r"^[A-Za-z0-9._][A-Za-z0-9._ \-]{0,127}$")


class GitHubPush:
    """Uploads a local folder to the repo as a single commit via the
    Git Data API (blobs → tree → commit → update ref). All files land
    under ``<target_folder>/`` on the configured branch.
    """

    def __init__(self, settings_provider):
        self._settings_provider = settings_provider
        self._thread: threading.Thread | None = None
        self._lock = threading.Lock()
        self._log: list[str] = []
        self._running = False
        self._done = False
        self._success = False
        self._started_at = 0.0
        self._ssl_ctx: ssl.SSLContext | None = None
        self._error: str = ""

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
                "error": self._error,
            }

    def start(self, target_folder: str, files: list[dict], message: str) -> dict:
        settings = self._settings_provider() or {}
        token = (settings.get("github_token") or "").strip()
        if not token:
            token = os.environ.get("GITHUB_TOKEN", "").strip()
        if not token:
            return {
                "ok": False,
                "error": "Geen GitHub-token. Vul er een in bij Settings → "
                         "GitHub-authenticatie, of zet de omgevingsvariabele "
                         "GITHUB_TOKEN (scope: repo).",
            }
        target = (target_folder or "").strip().strip("/")
        if not target or not SAFE_FOLDER_RE.match(target):
            return {"ok": False, "error": "Ongeldige doelmap-naam."}
        if not isinstance(files, list) or not files:
            return {"ok": False, "error": "Geen bestanden geselecteerd."}

        clean_files: list[dict] = []
        for f in files:
            rel = (f.get("path") or "").replace("\\", "/").lstrip("/")
            if not rel or ".." in rel.split("/"):
                return {"ok": False, "error": f"Ongeldig bestandspad: {rel!r}"}
            content_b64 = f.get("content_b64") or ""
            try:
                base64.b64decode(content_b64, validate=True)
            except Exception:
                return {"ok": False, "error": f"Ongeldige base64 voor {rel}"}
            clean_files.append({"rel": rel, "b64": content_b64})

        with self._lock:
            if self._running:
                return {"ok": False, "error": "Push is al bezig."}
            self._log = []
            self._error = ""
            self._running = True
            self._done = False
            self._success = False
            self._started_at = time.time()

        msg = (message or "").strip() or f"Update {target} via dashboard"
        self._thread = threading.Thread(
            target=self._run,
            args=(token, target, clean_files, msg),
            daemon=True,
        )
        self._thread.start()
        return {"ok": True}

    def _api(self, method: str, url: str, body: dict | None, token: str) -> dict:
        data = json.dumps(body).encode("utf-8") if body is not None else None
        req = urllib.request.Request(url, data=data, method=method, headers={
            "User-Agent": "Dashboard-AI-Worx/1.0",
            "Accept": "application/vnd.github+json",
            "Authorization": f"Bearer {token}",
            "X-GitHub-Api-Version": "2022-11-28",
            "Content-Type": "application/json",
        })
        with urllib.request.urlopen(req, timeout=60, context=self._ssl_ctx) as resp:
            raw = resp.read()
        return json.loads(raw.decode("utf-8")) if raw else {}

    def _run(self, token: str, target: str, files: list[dict], message: str) -> None:
        ok = False
        err_msg = ""
        try:
            settings = self._settings_provider() or {}
            verify = bool(settings.get("ssl_verify", True))
            self._ssl_ctx = build_ssl_context(verify=verify)

            base = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}"
            self._append(f"Start push naar {GITHUB_OWNER}/{GITHUB_REPO}@{BRANCH}")
            self._append(f"Doelmap: {target}  ({len(files)} bestanden)")

            # Preflight: verify the token works at all before we spend time
            # base64-ing blobs. GitHub returns 401 for bad tokens and 403 for
            # tokens missing the ``repo`` scope.
            self._append("Preflight: token controleren via /user...")
            try:
                who = self._api("GET", "https://api.github.com/user", None, token)
                self._append(f"Token ok (user: {who.get('login', '?')})")
            except urllib.error.HTTPError as exc:
                if exc.code == 401:
                    raise RuntimeError(
                        "Token ongeldig of verlopen. Vernieuw in Settings → "
                        "GitHub-authenticatie."
                    ) from exc
                raise

            # Preflight #2: verify this token can actually WRITE to the repo.
            # Authenticated GET /repos/<o>/<r> returns a ``permissions`` block
            # that tells us whether the token has push access. Catches the
            # common "fine-grained PAT without Contents: Read/Write" case
            # before we upload any blobs.
            self._append(f"Preflight: schrijfrechten controleren op {GITHUB_OWNER}/{GITHUB_REPO}...")
            try:
                info = self._api("GET", base, None, token)
            except urllib.error.HTTPError as exc:
                if exc.code == 404:
                    raise RuntimeError(
                        f"Repo {GITHUB_OWNER}/{GITHUB_REPO} niet vindbaar met deze token. "
                        "Bij een fine-grained PAT moet je deze repo expliciet "
                        "selecteren onder 'Repository access'."
                    ) from exc
                raise
            perms = (info or {}).get("permissions") or {}
            if not perms.get("push"):
                raise RuntimeError(
                    "Token heeft geen schrijfrechten op "
                    f"{GITHUB_OWNER}/{GITHUB_REPO}. Oplossing: "
                    "bij een FINE-GRAINED PAT de permissie 'Contents: "
                    "Read and write' aanzetten én deze repo kiezen onder "
                    "'Repository access'. Bij een CLASSIC PAT de 'repo' "
                    "scope aanvinken. Maak daarna een nieuwe token aan en "
                    "zet die in Settings → GitHub-authenticatie."
                )
            self._append("Schrijfrechten ok.")

            ref = self._api("GET", f"{base}/git/ref/heads/{BRANCH}", None, token)
            base_sha = ref["object"]["sha"]
            self._append(f"Branch HEAD: {base_sha[:7]}")

            base_commit = self._api("GET", f"{base}/git/commits/{base_sha}", None, token)
            base_tree = base_commit["tree"]["sha"]

            tree_entries: list[dict] = []
            for idx, f in enumerate(files, 1):
                blob = self._api("POST", f"{base}/git/blobs", {
                    "content": f["b64"],
                    "encoding": "base64",
                }, token)
                tree_entries.append({
                    "path": f"{target}/{f['rel']}",
                    "mode": "100644",
                    "type": "blob",
                    "sha": blob["sha"],
                })
                if idx % 5 == 0 or idx == len(files):
                    self._append(f"[{idx}/{len(files)}] blob: {f['rel']}")

            tree = self._api("POST", f"{base}/git/trees", {
                "base_tree": base_tree,
                "tree": tree_entries,
            }, token)

            commit = self._api("POST", f"{base}/git/commits", {
                "message": message,
                "tree": tree["sha"],
                "parents": [base_sha],
            }, token)

            self._api("PATCH", f"{base}/git/refs/heads/{BRANCH}", {
                "sha": commit["sha"],
            }, token)

            elapsed = time.time() - self._started_at
            self._append(f"Klaar: commit {commit['sha'][:7]} in {elapsed:.1f}s.")
            ok = True
        except urllib.error.HTTPError as exc:
            try:
                body = exc.read().decode("utf-8", errors="replace")
            except Exception:
                body = ""
            # Prefer GitHub's own message (it's precise and actionable).
            gh_msg = ""
            try:
                gh_msg = (json.loads(body) or {}).get("message", "")
            except Exception:
                pass

            # The specific "fine-grained PAT without Contents:Write" signature
            # — map it to a concrete fix. GitHub returns this on 403 during
            # write ops even when /user and /repos both succeed (e.g. when
            # the token was just downgraded, or org SSO lapsed).
            if exc.code == 403 and "not accessible by personal access token" in gh_msg.lower():
                err_msg = (
                    "Token mag niet schrijven naar deze repo. "
                    "Fine-grained PAT: zet 'Contents: Read and write' aan en "
                    "voeg deze repo toe aan 'Repository access'. Classic PAT: "
                    "vink de 'repo' scope aan. Herlaad daarna de token in "
                    "Settings → GitHub-authenticatie."
                )
            elif exc.code == 401:
                err_msg = "Token ongeldig of verlopen. Vernieuw in Settings → GitHub-authenticatie."
            else:
                err_msg = f"GitHub HTTP {exc.code} — {gh_msg or exc.reason}"

            self._append(f"FOUT: {err_msg}")
            if body and body.strip() != gh_msg:
                self._append(body[:500])
        except urllib.error.URLError as exc:
            reason = str(exc.reason)
            err_msg = f"Kan GitHub niet bereiken ({reason})"
            self._append(f"FOUT: {err_msg}.")
            if "CERTIFICATE_VERIFY_FAILED" in reason or "SSL" in reason.upper():
                self._append(
                    "! Corporate SSL-proxy? Zet in Settings \"SSL-certificaten "
                    "verifiëren\" uit en probeer opnieuw."
                )
        except Exception as exc:  # noqa: BLE001
            err_msg = str(exc) or exc.__class__.__name__
            self._append(f"FOUT: {err_msg}")
        finally:
            with self._lock:
                self._running = False
                self._done = True
                self._success = ok
                if not ok and not self._error:
                    self._error = err_msg or "Onbekende fout — zie log."


# ---------------------------------------------------------------------------
# In-dashboard job runner
# ---------------------------------------------------------------------------

MAX_LINES_PER_JOB = 5000  # cap per-job log memory; keep the tail when exceeded


class JobRunner:
    """Runs child processes with captured stdout/stderr so output can be
    streamed to the UI instead of opening a new console window."""

    def __init__(self):
        self._jobs: dict[str, dict] = {}
        self._lock = threading.Lock()

    # ---- lifecycle ------------------------------------------------------

    def start(self, cmd: list[str], cwd: str, name: str,
              on_finish=None) -> str:
        # Force UTF-8 everywhere for child Python (and PyInstaller-frozen
        # Python) processes. Without this, scripts that print or write
        # non-cp1252 characters (emoji, arrows, accented forms) crash
        # with ``UnicodeEncodeError: 'charmap' codec can't encode ...``
        # because Windows defaults the console codec to cp1252 *and*
        # ``open()`` falls back to the locale encoding when called
        # without an ``encoding=`` argument.
        #
        # ``PYTHONIOENCODING`` covers stdout/stderr; ``PYTHONUTF8=1``
        # puts the entire interpreter in UTF-8 Mode (PEP 540) so
        # ``open()`` and friends default to UTF-8 too. Both env-vars
        # are honoured by PyInstaller-frozen apps because the bootloader
        # reads them at interpreter init. The ``:replace`` suffix on
        # PYTHONIOENCODING ensures a stray glyph never halts a tool.
        # Inherits the rest of the parent environment so PATH etc.
        # still resolve.
        child_env = os.environ.copy()
        child_env["PYTHONIOENCODING"] = "utf-8:replace"
        child_env["PYTHONUTF8"] = "1"

        popen_kwargs = {
            "cwd": cwd,
            "stdout": subprocess.PIPE,
            "stderr": subprocess.STDOUT,
            "stdin": subprocess.DEVNULL,
            "bufsize": 1,              # line-buffered
            "text": True,
            "encoding": "utf-8",
            "errors": "replace",
            "env": child_env,
        }
        if os.name == "nt":
            # CREATE_NO_WINDOW hides the would-be console of the child so
            # the only visible output is our captured stream.
            popen_kwargs["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0)

        proc = subprocess.Popen(cmd, **popen_kwargs)
        job_id = os.urandom(8).hex()
        state = {
            "id": job_id,
            "name": name,
            "cmd": " ".join(cmd),
            "cmd_list": list(cmd),   # kept for relaunch_externally
            "cwd": cwd,
            "proc": proc,
            "log": [f"$ {' '.join(cmd)}"],
            "started_at": time.time(),
            "finished_at": None,
            "running": True,
            "exit_code": None,
            "truncated": 0,              # lines dropped due to MAX_LINES_PER_JOB
            "stopped_by_user": False,    # differentiates kill-by-user from crashes
            "relaunched_externally": False,
            "on_finish": on_finish,      # called with a frozen snapshot when done
        }
        with self._lock:
            self._jobs[job_id] = state
        threading.Thread(target=self._read_loop, args=(state,), daemon=True).start()
        return job_id

    def _read_loop(self, state: dict) -> None:
        proc: subprocess.Popen = state["proc"]
        try:
            assert proc.stdout is not None
            for raw in proc.stdout:
                line = raw.rstrip("\r\n")
                with self._lock:
                    state["log"].append(line)
                    overflow = len(state["log"]) - MAX_LINES_PER_JOB
                    if overflow > 0:
                        state["log"] = state["log"][overflow:]
                        state["truncated"] += overflow
        except Exception as exc:  # noqa: BLE001
            with self._lock:
                state["log"].append(f"[reader error: {exc}]")
        finally:
            try:
                rc = proc.wait()
            except Exception:  # noqa: BLE001
                rc = -1
            with self._lock:
                state["running"] = False
                state["exit_code"] = rc
                state["finished_at"] = time.time()
                state["log"].append(f"[proces beëindigd, exit code {rc}]")
                cb = state.get("on_finish")
                snapshot = {
                    "id": state["id"],
                    "name": state["name"],
                    "cmd": state["cmd"],
                    "cwd": state["cwd"],
                    "started_at": state["started_at"],
                    "finished_at": state["finished_at"],
                    "exit_code": rc,
                    "stopped_by_user": state["stopped_by_user"],
                    "log": list(state["log"]),
                    "truncated": state["truncated"],
                } if cb else None
        # Run callback OUTSIDE the lock — it may do I/O (write run record,
        # send mail) and we mustn't block other JobRunner readers.
        if cb is not None:
            try:
                cb(snapshot)
            except Exception as exc:  # noqa: BLE001
                print(f"[jobrunner] on_finish error: {exc}", file=sys.stderr)


    def stop(self, job_id: str) -> dict:
        with self._lock:
            state = self._jobs.get(job_id)
            if state:
                state["stopped_by_user"] = True
        if not state:
            return {"ok": False, "error": "Onbekende job."}
        proc: subprocess.Popen = state["proc"]
        if proc.poll() is not None:
            return {"ok": True, "already_done": True}
        try:
            proc.terminate()
        except OSError as exc:
            return {"ok": False, "error": str(exc)}

        # If it doesn't respond to terminate, kill it after a grace period.
        def _kill_guard():
            time.sleep(2.0)
            if proc.poll() is None:
                try:
                    proc.kill()
                except OSError:
                    pass

        threading.Thread(target=_kill_guard, daemon=True).start()
        return {"ok": True}

    # ---- external fallback ---------------------------------------------

    def relaunch_externally(self, job_id: str) -> dict:
        """Re-run the job's command in a new, visible console window.

        Used as a fallback when a tool doesn't work inside the embedded
        console (e.g. needs a real TTY, prompts for input, relies on
        console colour, or crashes immediately because it can't detect
        stdin/stdout). On Windows the child gets its own cmd.exe window
        via CREATE_NEW_CONSOLE; on other OSes it inherits stdio.
        """
        with self._lock:
            state = self._jobs.get(job_id)
        if not state:
            return {"ok": False, "error": "Onbekende job."}

        cmd = list(state.get("cmd_list") or [])
        cwd = state.get("cwd") or None
        if not cmd:
            return {"ok": False, "error": "Geen commando beschikbaar."}

        popen_kwargs: dict = {"cwd": cwd}
        if os.name == "nt":
            # CREATE_NEW_CONSOLE pops a fresh cmd.exe window so the tool
            # gets a real console (vs. our captured pipes).
            popen_kwargs["creationflags"] = getattr(subprocess, "CREATE_NEW_CONSOLE", 0)
        # On POSIX we simply inherit stdio — the dashboard may be launched
        # from a terminal; if not, output is lost, but the tool still runs.

        try:
            subprocess.Popen(cmd, **popen_kwargs)
        except FileNotFoundError:
            return {"ok": False, "error": "Launcher niet gevonden op PATH."}
        except OSError as exc:
            return {"ok": False, "error": f"Kan niet openen in nieuw venster: {exc}"}

        with self._lock:
            state["relaunched_externally"] = True
            state["log"].append("[tool opnieuw gestart in nieuw venster]")
        return {"ok": True}

    # ---- queries --------------------------------------------------------

    def status(self, job_id: str, since: int = 0) -> dict:
        with self._lock:
            state = self._jobs.get(job_id)
            if not state:
                return {"ok": False, "error": "Onbekende job."}
            total = len(state["log"]) + state["truncated"]
            # Clamp "since" so the UI can recover if older lines were dropped.
            start = max(0, since - state["truncated"])
            log_slice = state["log"][start:]
            next_offset = total
            elapsed = (state["finished_at"] or time.time()) - state["started_at"]
            return {
                "ok": True,
                "id": state["id"],
                "name": state["name"],
                "cmd": state["cmd"],
                "running": state["running"],
                "exit_code": state["exit_code"],
                "started_at": state["started_at"],
                "elapsed": elapsed,
                "log": log_slice,
                "next_offset": next_offset,
                "truncated": state["truncated"],
                "stopped_by_user": state["stopped_by_user"],
                "relaunched_externally": state["relaunched_externally"],
            }

    def list_all(self) -> list[dict]:
        with self._lock:
            return [
                {
                    "id": s["id"],
                    "name": s["name"],
                    "running": s["running"],
                    "exit_code": s["exit_code"],
                    "started_at": s["started_at"],
                }
                for s in self._jobs.values()
            ]


# ---------------------------------------------------------------------------
# Persistent run history (survives dashboard restarts)
# ---------------------------------------------------------------------------

DEFAULT_RUNS_DIR = APP_DIR / "schedule_runs"
DEFAULT_MAX_RUNS_PER_SCHEDULE = 50
RUNS_CAP_MIN = 1
RUNS_CAP_MAX = 10000
SAFE_ID_RE = re.compile(r"^[A-Za-z0-9_-]{1,64}$")
SAFE_NAME_RE = re.compile(r"[^A-Za-z0-9_-]+")


def _safe_filename(name: str, max_len: int = 60) -> str:
    """Return a filesystem-safe slug derived from ``name``. Empty input
    becomes ``"unnamed"``.
    """
    cleaned = SAFE_NAME_RE.sub("_", (name or "").strip())
    cleaned = cleaned.strip("_") or "unnamed"
    return cleaned[:max_len]


class RunStore:
    """Per-schedule run archive on disk, dual-written as JSON + TXT.

    Layout::

        <root>/<schedule_id>/<YYYY-MM-DD_HH-MM-SS_<slug>_<status>>.json
        <root>/<schedule_id>/<YYYY-MM-DD_HH-MM-SS_<slug>_<status>>.txt

    The JSON holds full meta + log (consumed by the Runs UI). The TXT
    holds only stdout, suitable for hand-off to an external automation.

    The root path and per-schedule cap come from ``settings_provider``
    so changing them in the UI takes effect on the next write — no
    restart required. ``fallback_log`` is invoked once with a string
    when the configured path is unusable and we fall back to default.
    """

    def __init__(self, settings_provider, fallback_log=None):
        self._settings_provider = settings_provider
        self._fallback_log = fallback_log
        self._lock = threading.Lock()
        self._fallback_signalled = False

    # ---- config resolution --------------------------------------------

    def _root(self) -> Path:
        cfg = (self._settings_provider() or {}).get("runs_storage") or {}
        raw = (cfg.get("path") or "").strip()
        if not raw:
            return DEFAULT_RUNS_DIR
        candidate = Path(raw)
        try:
            candidate.mkdir(parents=True, exist_ok=True)
            if candidate.is_dir():
                self._fallback_signalled = False  # recovered
                return candidate
        except OSError as exc:
            self._signal_fallback(f"Kon {raw} niet gebruiken: {exc}. Val terug op default.")
            return DEFAULT_RUNS_DIR
        self._signal_fallback(f"Pad {raw} is geen map. Val terug op default.")
        return DEFAULT_RUNS_DIR

    def _signal_fallback(self, msg: str) -> None:
        if self._fallback_signalled:
            return
        self._fallback_signalled = True
        if self._fallback_log:
            try:
                self._fallback_log(msg)
            except Exception:  # noqa: BLE001
                pass

    def _max_per_schedule(self) -> int:
        cfg = (self._settings_provider() or {}).get("runs_storage") or {}
        try:
            n = int(cfg.get("max_per_schedule") or DEFAULT_MAX_RUNS_PER_SCHEDULE)
        except (TypeError, ValueError):
            n = DEFAULT_MAX_RUNS_PER_SCHEDULE
        return max(RUNS_CAP_MIN, min(RUNS_CAP_MAX, n))

    def resolved_path(self) -> Path:
        return self._root()

    # ---- internal helpers ---------------------------------------------

    def _dir_for(self, schedule_id: str, root: Path | None = None) -> Path:
        if not SAFE_ID_RE.match(schedule_id or ""):
            raise ValueError(f"Onveilig schedule_id: {schedule_id!r}")
        return (root or self._root()) / schedule_id

    @staticmethod
    def _stem(path: Path) -> str:
        # ".../<stem>.json" → "<stem>"
        return path.stem

    @staticmethod
    def _build_run_id(record: dict) -> str:
        ts = record.get("started_at") or time.time()
        when = datetime.fromtimestamp(ts).strftime("%Y-%m-%d_%H-%M-%S")
        slug = _safe_filename(record.get("schedule_name") or "")
        status = _safe_filename(record.get("status") or "?")
        return f"{when}_{slug}_{status}"

    # ---- public API ---------------------------------------------------

    def save(self, record: dict) -> str:
        """Write a run record. Returns the run_id (filename stem)."""
        sid = record.get("schedule_id") or ""
        target_dir = self._dir_for(sid)
        with self._lock:
            target_dir.mkdir(parents=True, exist_ok=True)
            run_id = self._build_run_id(record)
            json_path = target_dir / f"{run_id}.json"
            # Avoid collision (same schedule, same status, same second).
            attempt = 0
            while json_path.exists():
                attempt += 1
                run_id = f"{self._build_run_id(record)}_{os.urandom(2).hex()}"
                json_path = target_dir / f"{run_id}.json"
                if attempt > 5:
                    break
            txt_path = target_dir / f"{run_id}.txt"

            record = dict(record)
            record["run_id"] = run_id
            try:
                with open(json_path, "w", encoding="utf-8") as fh:
                    json.dump(record, fh, ensure_ascii=False)
            except OSError as exc:
                self._signal_fallback(f"Kan {json_path} niet schrijven: {exc}")
                raise

            # Dual-write the plain stdout dump for external pickers.
            try:
                with open(txt_path, "w", encoding="utf-8") as fh:
                    fh.write("\n".join(record.get("log") or []))
                    if record.get("log"):
                        fh.write("\n")
            except OSError as exc:
                # JSON is the source of truth — best-effort .txt only.
                self._signal_fallback(f"Kon {txt_path} niet schrijven: {exc}")

            self._trim(target_dir)
        return run_id

    def _trim(self, target_dir: Path) -> None:
        cap = self._max_per_schedule()
        try:
            json_files = list(target_dir.glob("*.json"))
        except OSError:
            return
        # Sort by record.started_at (with name as tie-breaker) since the
        # filename is no longer epoch-sortable.
        def _started(p: Path) -> float:
            try:
                with open(p, "r", encoding="utf-8") as fh:
                    return float(json.load(fh).get("started_at") or 0)
            except (OSError, json.JSONDecodeError, TypeError, ValueError):
                return 0.0
        json_files.sort(key=lambda p: (_started(p), p.name))
        excess = len(json_files) - cap
        for f in json_files[:max(0, excess)]:
            stem = self._stem(f)
            for sibling in (f, target_dir / f"{stem}.txt"):
                try:
                    sibling.unlink()
                except OSError:
                    pass

    def list(self, schedule_id: str | None = None, limit: int = 50) -> list[dict]:
        """Return run records (without ``log``) newest first.

        ``schedule_id=None`` walks every subdir — handy for the global
        "Alle runs" view.
        """
        out: list[dict] = []
        root = self._root()
        if not root.exists():
            return out
        with self._lock:
            try:
                if schedule_id is None:
                    sub_dirs = [p for p in root.iterdir() if p.is_dir()]
                else:
                    d = self._dir_for(schedule_id, root=root)
                    sub_dirs = [d] if d.exists() else []
            except (OSError, ValueError):
                return out
            for d in sub_dirs:
                try:
                    files = list(d.glob("*.json"))
                except OSError:
                    continue
                for f in files[:limit * 2]:  # over-read; we sort+cap later
                    try:
                        with open(f, "r", encoding="utf-8") as fh:
                            rec = json.load(fh)
                    except (OSError, json.JSONDecodeError):
                        continue
                    rec.pop("log", None)
                    out.append(rec)
            out.sort(key=lambda r: r.get("started_at") or 0, reverse=True)
            return out[:limit]

    def get(self, schedule_id: str, run_id: str) -> dict | None:
        target_dir = self._dir_for(schedule_id)
        path = target_dir / f"{run_id}.json"
        with self._lock:
            if not path.exists():
                return None
            try:
                with open(path, "r", encoding="utf-8") as fh:
                    return json.load(fh)
            except (OSError, json.JSONDecodeError):
                return None

    def clear(self, schedule_id: str) -> int:
        target_dir = self._dir_for(schedule_id)
        with self._lock:
            if not target_dir.exists():
                return 0
            removed = 0
            for pattern in ("*.json", "*.txt"):
                for f in target_dir.glob(pattern):
                    try:
                        f.unlink()
                        if pattern == "*.json":
                            removed += 1
                    except OSError:
                        pass
            try:
                target_dir.rmdir()
            except OSError:
                pass
            return removed


# ---------------------------------------------------------------------------
# Scheduler — fires tool runs at planned moments
# ---------------------------------------------------------------------------

SCHEDULER_TICK_SECONDS = 15
SCHEDULER_LOG_LINES = 200


def _parse_hhmm(raw: str) -> tuple[int, int]:
    """Parse 'HH:MM'. Raises ValueError on bad input."""
    parts = (raw or "").split(":")
    if len(parts) != 2:
        raise ValueError("verwacht HH:MM")
    h, m = int(parts[0]), int(parts[1])
    if not (0 <= h <= 23 and 0 <= m <= 59):
        raise ValueError("uur/minuut buiten bereik")
    return h, m


def compute_next_run(sched: dict, after: float) -> float | None:
    """Return the next epoch (>= after) when this schedule should fire,
    or ``None`` if it never will (invalid config, or a 'once' run that's
    already past).
    """
    typ = sched.get("type")
    after_dt = datetime.fromtimestamp(after)

    if typ == "once":
        raw = sched.get("datetime") or ""
        try:
            target = datetime.fromisoformat(raw)
        except ValueError:
            return None
        ts = target.timestamp()
        return ts if ts >= after else None

    if typ == "interval":
        try:
            mins = int(sched.get("interval_minutes") or 0)
        except (TypeError, ValueError):
            return None
        if mins <= 0:
            return None
        anchor = sched.get("last_run") or sched.get("created_at") or after
        step = mins * 60
        nxt = anchor + step
        if nxt < after:
            # Skip past missed runs in one jump so we don't fire a backlog.
            steps = int((after - anchor) // step) + 1
            nxt = anchor + steps * step
        return nxt

    if typ == "daily":
        try:
            h, m = _parse_hhmm(sched.get("time") or "")
        except (ValueError, AttributeError):
            return None
        cand = after_dt.replace(hour=h, minute=m, second=0, microsecond=0)
        if cand.timestamp() < after:
            cand += timedelta(days=1)
        return cand.timestamp()

    if typ == "weekly":
        try:
            h, m = _parse_hhmm(sched.get("time") or "")
        except (ValueError, AttributeError):
            return None
        weekdays_raw = sched.get("weekdays") or []
        try:
            weekdays = sorted({int(w) for w in weekdays_raw if 0 <= int(w) <= 6})
        except (TypeError, ValueError):
            return None
        if not weekdays:
            return None
        for offset in range(0, 8):
            cand = (after_dt + timedelta(days=offset)).replace(
                hour=h, minute=m, second=0, microsecond=0
            )
            if cand.weekday() in weekdays and cand.timestamp() >= after:
                return cand.timestamp()
        return None

    return None


def validate_schedule(payload: dict) -> tuple[dict, str | None]:
    """Sanitize a schedule dict from the UI. Returns ``(clean, error)``.
    On error, ``clean`` is empty.
    """
    if not isinstance(payload, dict):
        return {}, "Ongeldige payload."

    name = (payload.get("name") or "").strip()
    if not name:
        return {}, "Naam mag niet leeg zijn."
    if len(name) > MAX_SCHEDULE_NAME:
        return {}, f"Naam mag max. {MAX_SCHEDULE_NAME} tekens zijn."

    rel_path = (payload.get("rel_path") or "").strip().replace("\\", "/").lstrip("/")
    if not rel_path or ".." in rel_path.split("/"):
        return {}, "Ongeldig tool-pad."

    typ = (payload.get("type") or "").strip()
    if typ not in VALID_SCHEDULE_TYPES:
        return {}, f"Onbekend schedule-type: {typ!r}"

    clean: dict = {
        "name": name,
        "rel_path": rel_path,
        "type": typ,
        "enabled": bool(payload.get("enabled", True)),
        "skip_if_running": bool(payload.get("skip_if_running", True)),
    }

    if typ == "once":
        raw = (payload.get("datetime") or "").strip()
        try:
            datetime.fromisoformat(raw)
        except ValueError:
            return {}, "Ongeldige datum/tijd voor 'eenmalig'."
        clean["datetime"] = raw

    elif typ == "interval":
        try:
            mins = int(payload.get("interval_minutes") or 0)
        except (TypeError, ValueError):
            return {}, "Interval moet een getal in minuten zijn."
        if mins <= 0:
            return {}, "Interval moet groter dan 0 minuten zijn."
        clean["interval_minutes"] = mins

    elif typ == "daily":
        try:
            _parse_hhmm(payload.get("time") or "")
        except ValueError as exc:
            return {}, f"Ongeldige tijd: {exc}"
        clean["time"] = payload.get("time")

    elif typ == "weekly":
        try:
            _parse_hhmm(payload.get("time") or "")
        except ValueError as exc:
            return {}, f"Ongeldige tijd: {exc}"
        weekdays_raw = payload.get("weekdays") or []
        if not isinstance(weekdays_raw, list) or not weekdays_raw:
            return {}, "Kies minstens één weekdag."
        try:
            weekdays = sorted({int(w) for w in weekdays_raw if 0 <= int(w) <= 6})
        except (TypeError, ValueError):
            return {}, "Ongeldige weekdag-waarde."
        if not weekdays:
            return {}, "Kies minstens één weekdag."
        clean["time"] = payload.get("time")
        clean["weekdays"] = weekdays

    return clean, None


class Scheduler:
    """Fires due schedules in a background thread.

    The scheduler doesn't know about tools or jobs directly — it calls
    ``fire(sched)`` (provided by ``Api``) which actually launches the
    tool through the existing ``JobRunner``. Persistence happens via
    ``settings.json`` (the same lock the Api uses for mutations).
    """

    def __init__(self, settings_provider, save_fn, fire, lock):
        self._settings_provider = settings_provider
        self._save = save_fn
        self._fire = fire
        self._lock = lock
        self._stop = threading.Event()
        self._thread: threading.Thread | None = None
        # Heartbeat / observability — read by Api.get_scheduler_status.
        self._started_at: float | None = None
        self._last_tick_at: float = 0.0
        self._tick_count: int = 0
        self._fired_count: int = 0
        self._log: list[dict] = []   # ring buffer of recent events

    def _log_event(self, kind: str, message: str, **extra) -> None:
        entry = {"ts": time.time(), "kind": kind, "message": message}
        if extra:
            entry.update(extra)
        self._log.append(entry)
        overflow = len(self._log) - SCHEDULER_LOG_LINES
        if overflow > 0:
            self._log = self._log[overflow:]

    def status(self) -> dict:
        with self._lock:
            return {
                "started_at": self._started_at,
                "last_tick_at": self._last_tick_at,
                "tick_count": self._tick_count,
                "fired_count": self._fired_count,
                "tick_seconds": SCHEDULER_TICK_SECONDS,
                "log": list(self._log),
                "alive": bool(self._thread and self._thread.is_alive()),
            }

    def start(self) -> None:
        if self._thread and self._thread.is_alive():
            return
        with self._lock:
            self._started_at = time.time()
            self._log_event("started", "Scheduler gestart.")
        self._thread = threading.Thread(target=self._loop, daemon=True)
        self._thread.start()

    def stop(self) -> None:
        self._stop.set()

    def _loop(self) -> None:
        while not self._stop.is_set():
            try:
                self._tick()
            except Exception as exc:  # noqa: BLE001
                with self._lock:
                    self._log_event("error", f"Tick crashte: {exc}")
            self._stop.wait(SCHEDULER_TICK_SECONDS)

    def _tick(self) -> None:
        now = time.time()
        with self._lock:
            self._last_tick_at = now
            self._tick_count += 1

            settings = self._settings_provider()
            scheds = settings.get("schedules")
            if not isinstance(scheds, list):
                self._log_event("tick", "Tick: geen schedules-lijst.")
                return
            changed = False
            fired_this_tick = 0
            for s in scheds:
                if not isinstance(s, dict):
                    continue
                if not s.get("enabled"):
                    continue
                nxt = s.get("next_run")
                if not isinstance(nxt, (int, float)) or nxt <= 0:
                    nxt = compute_next_run(s, now)
                    if nxt:
                        s["next_run"] = nxt
                        changed = True
                if not nxt or nxt > now:
                    continue

                # Due. Try to fire.
                try:
                    result = self._fire(s) or {}
                except Exception as exc:  # noqa: BLE001
                    result = {"ok": False, "error": str(exc)}

                if result.get("skipped"):
                    s["last_status"] = "skipped"
                    s["last_message"] = result.get("error") or "Overgeslagen"
                    self._log_event(
                        "skipped",
                        f"{s.get('name')}: {s['last_message']}",
                        schedule_id=s.get("id"),
                        schedule_name=s.get("name"),
                    )
                elif result.get("ok"):
                    s["last_status"] = "ok"
                    s["last_message"] = f"job {(result.get('job_id') or '')[:8]}"
                    s["last_job_id"] = result.get("job_id") or ""
                    fired_this_tick += 1
                    self._fired_count += 1
                    self._log_event(
                        "fired",
                        f"{s.get('name')} → job {s['last_job_id'][:8]}",
                        schedule_id=s.get("id"),
                        schedule_name=s.get("name"),
                        job_id=s["last_job_id"],
                    )
                else:
                    s["last_status"] = "error"
                    s["last_message"] = str(result.get("error") or "Onbekende fout")
                    self._log_event(
                        "error",
                        f"{s.get('name')}: {s['last_message']}",
                        schedule_id=s.get("id"),
                        schedule_name=s.get("name"),
                    )

                s["last_run"] = now
                if s.get("type") == "once":
                    s["enabled"] = False
                    s["next_run"] = None
                else:
                    s["next_run"] = compute_next_run(s, now + 1)
                changed = True

            self._log_event(
                "tick",
                f"Tick: {len(scheds)} schedules, {fired_this_tick} gevuurd.",
            )
            if changed:
                self._save()


# ---------------------------------------------------------------------------
# API (same surface as the old PyWebView js_api)
# ---------------------------------------------------------------------------

class Api:
    def __init__(self):
        self.settings = load_settings()
        self.sync = GitHubSync(settings_provider=lambda: self.settings)
        self.push = GitHubPush(settings_provider=lambda: self.settings)
        self.jobs = JobRunner()
        self.runs = RunStore(
            settings_provider=lambda: self.settings,
            fallback_log=self._note_runstore_fallback,
        )
        self._sched_lock = threading.Lock()
        self.scheduler = Scheduler(
            settings_provider=lambda: self.settings,
            save_fn=lambda: save_settings(self.settings),
            fire=self._fire_schedule,
            lock=self._sched_lock,
        )

    def get_tools(self) -> dict:
        return discover_tools(self.settings)

    # ----- categories ---------------------------------------------------

    def _categories_payload(self) -> dict:
        cats = self.settings.setdefault(
            "categories", {"order": [], "assignments": {}}
        )
        return {
            "ok": True,
            "order": list(cats.get("order") or []),
            "assignments": dict(cats.get("assignments") or {}),
        }

    def add_category(self, name: str) -> dict:
        clean, err = _normalise_category_name(name)
        if err:
            return {"ok": False, "error": err}
        cats = self.settings.setdefault(
            "categories", {"order": [], "assignments": {}}
        )
        order = cats.setdefault("order", [])
        if any(c.lower() == clean.lower() for c in order):
            return {"ok": False, "error": f"Categorie '{clean}' bestaat al."}
        order.append(clean)
        save_settings(self.settings)
        return self._categories_payload()

    def rename_category(self, old: str, new: str) -> dict:
        clean, err = _normalise_category_name(new)
        if err:
            return {"ok": False, "error": err}
        cats = self.settings.setdefault(
            "categories", {"order": [], "assignments": {}}
        )
        order = cats.setdefault("order", [])
        if old not in order:
            return {"ok": False, "error": f"Onbekende categorie: '{old}'."}
        if old != clean and any(c.lower() == clean.lower() for c in order):
            return {"ok": False, "error": f"Categorie '{clean}' bestaat al."}
        order[order.index(old)] = clean
        assigns = cats.setdefault("assignments", {})
        for path, cat in list(assigns.items()):
            if cat == old:
                assigns[path] = clean
        save_settings(self.settings)
        return self._categories_payload()

    def delete_category(self, name: str) -> dict:
        cats = self.settings.setdefault(
            "categories", {"order": [], "assignments": {}}
        )
        order = cats.setdefault("order", [])
        if name not in order:
            return {"ok": False, "error": f"Onbekende categorie: '{name}'."}
        order.remove(name)
        assigns = cats.setdefault("assignments", {})
        for path, cat in list(assigns.items()):
            if cat == name:
                del assigns[path]
        save_settings(self.settings)
        return self._categories_payload()

    def reorder_categories(self, names) -> dict:
        if not isinstance(names, list):
            return {"ok": False, "error": "Verwacht een lijst met namen."}
        cats = self.settings.setdefault(
            "categories", {"order": [], "assignments": {}}
        )
        order = cats.setdefault("order", [])
        if sorted(map(str, names)) != sorted(order):
            return {
                "ok": False,
                "error": "Volgorde komt niet exact overeen met huidige categorieën.",
            }
        cats["order"] = [str(n) for n in names]
        save_settings(self.settings)
        return self._categories_payload()

    def set_tool_category(self, rel_path: str, category: str) -> dict:
        if not (rel_path or "").strip():
            return {"ok": False, "error": "Geen tool opgegeven."}
        cats = self.settings.setdefault(
            "categories", {"order": [], "assignments": {}}
        )
        assigns = cats.setdefault("assignments", {})
        target = (category or "").strip()
        if not target or target.lower() == UNCATEGORIZED.lower():
            assigns.pop(rel_path, None)
        else:
            order = cats.setdefault("order", [])
            if target not in order:
                return {
                    "ok": False,
                    "error": f"Categorie '{target}' bestaat niet — maak hem eerst aan.",
                }
            assigns[rel_path] = target
        save_settings(self.settings)
        return self._categories_payload()

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

    def set_ssl_verify(self, enabled: bool) -> dict:
        self.settings["ssl_verify"] = bool(enabled)
        save_settings(self.settings)
        return {"ok": True, "ssl_verify": self.settings["ssl_verify"]}

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

    def has_github_token(self) -> dict:
        if (self.settings.get("github_token") or "").strip():
            return {"present": True, "source": "settings"}
        if os.environ.get("GITHUB_TOKEN", "").strip():
            return {"present": True, "source": "env"}
        return {"present": False, "source": None}

    def set_github_token(self, token: str) -> dict:
        token = (token or "").strip()
        if token:
            self.settings["github_token"] = token
        else:
            self.settings.pop("github_token", None)
        save_settings(self.settings)
        return self.has_github_token()

    def push_folder(self, target_folder: str, files, message: str) -> dict:
        return self.push.start(target_folder, files or [], message or "")

    def get_push_status(self) -> dict:
        return self.push.status()

    def _resolve_launch(self, script_path: str) -> dict:
        """Resolve a tool path to an executable command + cwd + name.

        Shared by ``launch_tool`` (manual + scheduler) and
        ``spawn_tool_window`` (manual pop-up). Returns either
        ``{"ok": True, "cmd": [...], "cwd": str, "name": str, "launcher": str}``
        or ``{"ok": False, "error": str}`` so callers can short-circuit
        with the same error UX.
        """
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

        return {
            "ok": True,
            "cmd": cmd,
            "cwd": str(path.parent),
            "name": path.stem,
            "launcher": launcher,
        }

    def launch_tool(self, script_path: str, on_finish=None) -> dict:
        resolved = self._resolve_launch(script_path)
        if not resolved.get("ok"):
            return resolved

        cmd = resolved["cmd"]
        try:
            job_id = self.jobs.start(
                cmd,
                cwd=resolved["cwd"],
                name=resolved["name"],
                on_finish=on_finish,
            )
        except FileNotFoundError:
            return {
                "ok": False,
                "error": f"Launcher '{resolved['launcher']}' niet gevonden op PATH.",
            }
        except OSError as exc:
            return {"ok": False, "error": f"Kan tool niet starten: {exc}"}

        return {
            "ok": True,
            "job_id": job_id,
            "name": resolved["name"],
            "cmd": " ".join(cmd),
        }

    def spawn_tool_window(self, script_path: str) -> dict:
        """Start a tool and open a dedicated chromeless pop-up window
        showing its live output. The pop-up uses the same Edge/Chrome
        ``--app=`` shell as the main dashboard, so it gets the same
        look-and-feel without an address bar or browser chrome.

        Replaces the old embedded-console flow: every manual launch
        now goes through here. The pop-up polls ``get_job`` exactly
        like the embedded console used to.
        """
        resolved = self._resolve_launch(script_path)
        if not resolved.get("ok"):
            return resolved

        cmd = resolved["cmd"]
        try:
            job_id = self.jobs.start(
                cmd, cwd=resolved["cwd"], name=resolved["name"],
            )
        except FileNotFoundError:
            return {
                "ok": False,
                "error": f"Launcher '{resolved['launcher']}' niet gevonden op PATH.",
            }
        except OSError as exc:
            return {"ok": False, "error": f"Kan tool niet starten: {exc}"}

        # Build URL for the pop-up. Server reads the auth token + base
        # URL from class-level attributes set by ``main()`` once the
        # HTTP server is up. Window dimensions are kept compact —
        # console UX is vertical-scroll dominant.
        token = getattr(type(self), "_AUTH_TOKEN", "")
        port = getattr(type(self), "_PORT", 0)
        if not token or not port:
            return {"ok": True, "job_id": job_id, "name": resolved["name"], "window": False}

        from urllib.parse import quote
        url = (
            f"http://127.0.0.1:{port}/tool-window"
            f"?token={token}"
            f"&job={quote(job_id)}"
            f"&name={quote(resolved['name'])}"
        )
        try:
            launch_browser(url, size="960,640")
        except Exception as exc:  # noqa: BLE001 — never fail the launch over a window
            return {
                "ok": True,
                "job_id": job_id,
                "name": resolved["name"],
                "window": False,
                "window_error": str(exc),
            }

        return {
            "ok": True,
            "job_id": job_id,
            "name": resolved["name"],
            "window": True,
        }

    # ----- job control ---------------------------------------------------

    def get_job(self, job_id: str, since: int = 0) -> dict:
        return self.jobs.status(job_id, since)

    def stop_job(self, job_id: str) -> dict:
        return self.jobs.stop(job_id)

    def relaunch_externally(self, job_id: str) -> dict:
        return self.jobs.relaunch_externally(job_id)

    def list_jobs(self) -> dict:
        return {"jobs": self.jobs.list_all()}

    # ----- scheduler ----------------------------------------------------

    def _fire_schedule(self, sched: dict) -> dict:
        """Called by the scheduler when a schedule is due. Resolves the
        tool by ``rel_path`` against the live repo and launches it. The
        scheduler holds ``self._sched_lock`` for the duration, so we do
        NOT acquire it again here.

        Attaches an ``on_finish`` callback so the captured stdout, exit
        code and duration are persisted as a run record (JSON + TXT) on
        disk for later inspection.
        """
        rel_path = (sched.get("rel_path") or "").strip()
        if not rel_path:
            return {"ok": False, "error": "Geen tool gekoppeld."}

        if sched.get("skip_if_running"):
            last_job_id = sched.get("last_job_id") or ""
            if last_job_id:
                with self.jobs._lock:  # noqa: SLF001 — internal but safe
                    state = self.jobs._jobs.get(last_job_id)  # noqa: SLF001
                    if state and state.get("running"):
                        return {
                            "ok": False,
                            "skipped": True,
                            "error": "Vorige run draait nog (overgeslagen).",
                        }

        abs_path = REPO_DIR / rel_path
        if not abs_path.exists() or not abs_path.is_file():
            # Persist a stub run record so failures are visible too.
            self._record_synthetic_run(sched, "error", f"Tool niet gevonden: {rel_path}")
            return {"ok": False, "error": f"Tool niet gevonden: {rel_path}"}

        # Snapshot only the fields the callback actually needs — we
        # don't want it to hold a reference to a mutating dict from
        # settings.json.
        sched_snapshot = {
            "id": sched.get("id"),
            "name": sched.get("name"),
            "rel_path": sched.get("rel_path"),
        }

        def on_finish(snapshot):
            try:
                self._record_finished_run(sched_snapshot, snapshot)
            except Exception as exc:  # noqa: BLE001
                print(f"[api] _record_finished_run error: {exc}", file=sys.stderr)

        return self.launch_tool(str(abs_path), on_finish=on_finish)

    def _record_synthetic_run(self, sched: dict, status: str, message: str) -> None:
        """Persist a run record for a schedule that never actually ran a
        process (e.g. the tool file was missing). Keeps Runs UI honest.
        """
        sid = sched.get("id") or ""
        if not sid:
            return
        now = time.time()
        record = {
            "schedule_id": sid,
            "schedule_name": sched.get("name") or "",
            "rel_path": sched.get("rel_path") or "",
            "started_at": now,
            "finished_at": now,
            "exit_code": None,
            "status": status,
            "log": [message],
            "truncated": 0,
            "stopped_by_user": False,
        }
        try:
            self.runs.save(record)
        except (OSError, ValueError) as exc:
            print(f"[api] kan synthetisch run-record niet opslaan: {exc}", file=sys.stderr)

    def _record_finished_run(self, sched_snapshot: dict, run_snapshot: dict) -> None:
        """Called from a JobRunner thread after a scheduled tool exits.
        Writes the run record (JSON + TXT) to disk; nothing else.
        """
        sid = sched_snapshot.get("id") or ""
        if not sid:
            return
        rc = run_snapshot.get("exit_code")
        status = "ok" if rc == 0 else "error"
        record = {
            "schedule_id": sid,
            "schedule_name": sched_snapshot.get("name") or "",
            "rel_path": sched_snapshot.get("rel_path") or "",
            "started_at": run_snapshot.get("started_at"),
            "finished_at": run_snapshot.get("finished_at"),
            "exit_code": rc,
            "status": status,
            "log": run_snapshot.get("log") or [],
            "truncated": run_snapshot.get("truncated", 0),
            "stopped_by_user": run_snapshot.get("stopped_by_user", False),
        }
        try:
            self.runs.save(record)
        except (OSError, ValueError) as exc:
            print(f"[api] kan run-record niet opslaan: {exc}", file=sys.stderr)

    def _note_runstore_fallback(self, msg: str) -> None:
        """Surface a RunStore fallback (bad path, etc.) in the scheduler
        tick log so the user can see it in the Runs page.
        """
        try:
            with self._sched_lock:
                self.scheduler._log_event("error", f"Runs-opslag: {msg}")
        except Exception:  # noqa: BLE001
            pass

    def _schedule_view(self, sched: dict) -> dict:
        """Add UI-friendly fields (next_run, pending) without mutating
        the stored dict.

        Honours the stored ``next_run`` when it's still in the past — that
        run is pending and the scheduler will fire it on the next tick.
        Only when the stored value is in the future, or absent, do we
        fall back to a freshly computed value so the UI stays accurate
        across edits.
        """
        out = dict(sched)
        out.pop("last_job_id", None)  # internal — don't leak to UI
        now = time.time()
        stored = sched.get("next_run")
        enabled = bool(sched.get("enabled"))
        pending = False
        if not enabled:
            out["next_run"] = None
        elif isinstance(stored, (int, float)) and stored > 0 and stored <= now:
            # Past stored next_run = a missed/queued fire. Show that, not
            # a misleading "tomorrow".
            out["next_run"] = stored
            pending = True
        elif isinstance(stored, (int, float)) and stored > now:
            out["next_run"] = stored
        else:
            out["next_run"] = compute_next_run(sched, now)
        out["pending"] = pending
        return out

    def _schedules_payload(self) -> dict:
        scheds = self.settings.setdefault("schedules", [])
        return {
            "ok": True,
            "schedules": [self._schedule_view(s) for s in scheds if isinstance(s, dict)],
        }

    def list_schedules(self) -> dict:
        with self._sched_lock:
            return self._schedules_payload()

    def add_schedule(self, payload) -> dict:
        clean, err = validate_schedule(payload or {})
        if err:
            return {"ok": False, "error": err}
        with self._sched_lock:
            now = time.time()
            clean["id"] = os.urandom(6).hex()
            clean["created_at"] = now
            clean["last_run"] = None
            clean["last_status"] = None
            clean["last_message"] = ""
            clean["last_job_id"] = ""
            clean["next_run"] = compute_next_run(clean, now)
            self.settings.setdefault("schedules", []).append(clean)
            save_settings(self.settings)
            return self._schedules_payload()

    def update_schedule(self, schedule_id: str, payload) -> dict:
        clean, err = validate_schedule(payload or {})
        if err:
            return {"ok": False, "error": err}
        with self._sched_lock:
            scheds = self.settings.setdefault("schedules", [])
            for s in scheds:
                if s.get("id") == schedule_id:
                    # Preserve history fields; replace config fields.
                    s["name"] = clean["name"]
                    s["rel_path"] = clean["rel_path"]
                    s["type"] = clean["type"]
                    s["enabled"] = clean["enabled"]
                    s["skip_if_running"] = clean["skip_if_running"]
                    for k in ("datetime", "interval_minutes", "time", "weekdays"):
                        if k in clean:
                            s[k] = clean[k]
                        else:
                            s.pop(k, None)
                    s["next_run"] = compute_next_run(s, time.time())
                    save_settings(self.settings)
                    return self._schedules_payload()
            return {"ok": False, "error": "Onbekend schedule-id."}

    def delete_schedule(self, schedule_id: str) -> dict:
        with self._sched_lock:
            scheds = self.settings.setdefault("schedules", [])
            new = [s for s in scheds if s.get("id") != schedule_id]
            if len(new) == len(scheds):
                return {"ok": False, "error": "Onbekend schedule-id."}
            self.settings["schedules"] = new
            save_settings(self.settings)
            return self._schedules_payload()

    def toggle_schedule(self, schedule_id: str, enabled: bool) -> dict:
        with self._sched_lock:
            scheds = self.settings.setdefault("schedules", [])
            for s in scheds:
                if s.get("id") == schedule_id:
                    s["enabled"] = bool(enabled)
                    s["next_run"] = (
                        compute_next_run(s, time.time()) if s["enabled"] else None
                    )
                    save_settings(self.settings)
                    return self._schedules_payload()
            return {"ok": False, "error": "Onbekend schedule-id."}

    def run_schedule_now(self, schedule_id: str) -> dict:
        """Manually fire a schedule (useful for testing). Doesn't update
        ``last_run`` so the normal cadence isn't disturbed.
        """
        with self._sched_lock:
            scheds = self.settings.setdefault("schedules", [])
            target = next((s for s in scheds if s.get("id") == schedule_id), None)
            if not target:
                return {"ok": False, "error": "Onbekend schedule-id."}
            result = self._fire_schedule(target) or {}
            return result

    # ----- scheduler observability + run history -----------------------

    def get_scheduler_status(self) -> dict:
        return self.scheduler.status()

    def list_runs(self, schedule_id: str | None = None, limit: int = 50) -> dict:
        try:
            limit = int(limit or 50)
        except (TypeError, ValueError):
            limit = 50
        limit = max(1, min(500, limit))
        try:
            runs = self.runs.list(schedule_id, limit=limit)
        except (OSError, ValueError) as exc:
            return {"ok": False, "error": str(exc), "runs": []}
        return {"ok": True, "runs": runs}

    def get_run(self, schedule_id: str, run_id: str) -> dict:
        try:
            rec = self.runs.get(schedule_id, run_id)
        except (OSError, ValueError) as exc:
            return {"ok": False, "error": str(exc)}
        if not rec:
            return {"ok": False, "error": "Run niet gevonden."}
        return {"ok": True, "run": rec}

    def clear_runs(self, schedule_id: str) -> dict:
        try:
            removed = self.runs.clear(schedule_id)
        except (OSError, ValueError) as exc:
            return {"ok": False, "error": str(exc)}
        return {"ok": True, "removed": removed}

    # ----- runs storage (where output files land) ----------------------

    def get_runs_storage(self) -> dict:
        cfg = (self.settings.get("runs_storage") or {}).copy()
        resolved = self.runs.resolved_path()
        is_default = resolved == DEFAULT_RUNS_DIR
        configured = (cfg.get("path") or "").strip()
        return {
            "ok": True,
            "path": configured,
            "max_per_schedule": cfg.get("max_per_schedule", DEFAULT_MAX_RUNS_PER_SCHEDULE),
            "resolved_path": str(resolved),
            "default_path": str(DEFAULT_RUNS_DIR),
            "is_default": is_default,
            "fallback": bool(configured) and is_default,
            "cap_min": RUNS_CAP_MIN,
            "cap_max": RUNS_CAP_MAX,
        }

    def set_runs_storage(self, payload: dict) -> dict:
        if not isinstance(payload, dict):
            return {"ok": False, "error": "Ongeldige payload."}
        cfg = self.settings.setdefault(
            "runs_storage", dict(DEFAULT_SETTINGS["runs_storage"])
        )
        if "path" in payload:
            raw = (payload.get("path") or "").strip()
            if raw:
                p = Path(raw)
                if not p.is_absolute():
                    return {"ok": False, "error": "Pad moet absoluut zijn."}
                try:
                    p.mkdir(parents=True, exist_ok=True)
                except OSError as exc:
                    return {"ok": False, "error": f"Kan map niet aanmaken: {exc}"}
                if not p.is_dir():
                    return {"ok": False, "error": "Pad bestaat maar is geen map."}
            cfg["path"] = raw
        if "max_per_schedule" in payload:
            try:
                n = int(payload["max_per_schedule"])
            except (TypeError, ValueError):
                return {"ok": False, "error": "Maximum moet een getal zijn."}
            if not (RUNS_CAP_MIN <= n <= RUNS_CAP_MAX):
                return {
                    "ok": False,
                    "error": f"Maximum buiten bereik ({RUNS_CAP_MIN}–{RUNS_CAP_MAX}).",
                }
            cfg["max_per_schedule"] = n
        save_settings(self.settings)
        return self.get_runs_storage()

    def open_runs_folder(self) -> dict:
        path = self.runs.resolved_path()
        try:
            path.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            return {"ok": False, "error": f"Kan map niet openen: {exc}"}
        try:
            if os.name == "nt":
                os.startfile(str(path))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except (OSError, FileNotFoundError) as exc:
            return {"ok": False, "error": f"Kan map niet openen: {exc}"}
        return {"ok": True, "path": str(path)}

    # Called by the HTTP server when the UI navigates away / closes.
    def shutdown(self) -> dict:
        threading.Thread(target=lambda: (time.sleep(0.2), os._exit(0)), daemon=True).start()
        return {"ok": True}


# ---------------------------------------------------------------------------
# HTTP server (serves ui.html + /api/<method> JSON-RPC)
# ---------------------------------------------------------------------------

def build_handler(api: Api, ui_html: bytes, tool_window_html: bytes, auth_token: str):
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
            if path == "/tool-window":
                if not self._check_token():
                    self.send_error(403, "Forbidden")
                    return
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(tool_window_html)))
                self.send_header("Cache-Control", "no-store")
                self.end_headers()
                self.wfile.write(tool_window_html)
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


def launch_browser(url: str, size: str = "1280,820") -> subprocess.Popen | None:
    """Open ``url`` in a chromeless Edge/Chrome ``--app=`` window.

    ``size`` is forwarded as ``--window-size``; the default fits the main
    dashboard, while tool pop-ups pass a compacter ``"960,640"``.
    """
    browser = find_chromium_browser()
    profile_dir = APP_DIR / ".browser_profile"
    profile_dir.mkdir(exist_ok=True)
    if browser:
        args = [
            browser,
            f"--app={url}",
            f"--user-data-dir={profile_dir}",
            # ``--new-window`` defends against Chromium reusing a window
            # when two ``--app=`` invocations share the same profile and
            # arrive close in time. The unique ``?job=<id>`` in the URL
            # already differentiates them, but this is gratis insurance.
            "--new-window",
            f"--window-size={size}",
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
    if not TOOL_WINDOW_PATH.exists():
        raise SystemExit(f"tool_window.html ontbreekt (verwacht: {TOOL_WINDOW_PATH})")

    load_settings()  # ensure settings.json exists
    with open(UI_PATH, "rb") as fh:
        ui_html = fh.read()
    with open(TOOL_WINDOW_PATH, "rb") as fh:
        tool_window_html = fh.read()

    api = Api()
    api.scheduler.start()
    auth_token = os.urandom(16).hex()
    port = pick_free_port()
    # ``spawn_tool_window`` reads these to build pop-up URLs.
    Api._AUTH_TOKEN = auth_token
    Api._PORT = port
    server = ThreadingHTTPServer(
        ("127.0.0.1", port),
        build_handler(api, ui_html, tool_window_html, auth_token),
    )

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
