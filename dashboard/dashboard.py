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
        entry = find_entry_file(sub)
        if entry is None:
            # No obvious executable → treat folder as pure supporting
            # material (docs, assets, ...) and hide it from the dashboard.
            continue
        category_label = snake_to_title(sub.name)
        category_desc = read_readme_description(sub) or f"Tools in {category_label.lower()}."
        add_tool(sub.name, category_label, category_desc, entry)

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

    def start(self, cmd: list[str], cwd: str, name: str) -> str:
        popen_kwargs = {
            "cwd": cwd,
            "stdout": subprocess.PIPE,
            "stderr": subprocess.STDOUT,
            "stdin": subprocess.DEVNULL,
            "bufsize": 1,              # line-buffered
            "text": True,
            "encoding": "utf-8",
            "errors": "replace",
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
# API (same surface as the old PyWebView js_api)
# ---------------------------------------------------------------------------

class Api:
    def __init__(self):
        self.settings = load_settings()
        self.sync = GitHubSync(settings_provider=lambda: self.settings)
        self.push = GitHubPush(settings_provider=lambda: self.settings)
        self.jobs = JobRunner()

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
            job_id = self.jobs.start(cmd, cwd=str(path.parent), name=path.stem)
        except FileNotFoundError:
            return {"ok": False, "error": f"Launcher '{launcher}' niet gevonden op PATH."}
        except OSError as exc:
            return {"ok": False, "error": f"Kan tool niet starten: {exc}"}

        return {
            "ok": True,
            "job_id": job_id,
            "name": path.stem,
            "cmd": " ".join(cmd),
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
