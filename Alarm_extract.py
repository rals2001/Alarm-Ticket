"""
PMT Alarm – Closed Alarm downloader (3-step workflow, Playwright edition)

STEP 1  Retrieve Excel  -> login in Chromium, capture the export API call, download the data via requests
STEP 2  Filter Excel    -> stream the API JSON, keep the chosen Year-Month, write XLSX
STEP 3  Download Result -> hand the processed XLSX to the browser (no processing)

All state lives on disk in a per-session job folder (not only in st.session_state),
so it survives websocket reconnects and page refreshes (the job id is kept in the URL).
"""
import faulthandler
import json
import logging
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import traceback
import uuid
from urllib.parse import urlparse
from datetime import date, datetime, timedelta, timezone
from itertools import chain, islice
from pathlib import Path

import streamlit as st
import ijson
import requests
from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel
from playwright.sync_api import Error as PWError
from playwright.sync_api import TimeoutError as PWTimeout
from playwright.sync_api import sync_playwright

try:
    import psutil
except ImportError:  # diagnostics degrade gracefully
    psutil = None

# ===================== CONFIGURATION =====================
LOGIN_URL = "https://pmt-alarm.komdigi.go.id/auth/login"
ALARM_URL = "https://pmt-alarm.komdigi.go.id/dashboard/alarm"
DOWNLOAD_TIMEOUT = 900        # seconds to wait for the big export (15 min)
DEFAULT_TIMEOUT = 40          # seconds for normal waits
JOBS_ROOT = Path(tempfile.gettempdir()) / "pmt_alarm_jobs"
STALE_AFTER = 300             # a "running" step with no heartbeat for 5 min = interrupted
JOB_MAX_AGE_HOURS = 12
POLL_SECONDS = 3
MEM_GUARD_FRACTION = 0.85     # kill Chromium when the container reaches 85 % of its memory limit
MEM_LIMIT_MB_OVERRIDE = float(os.environ.get("MEM_LIMIT_MB", 0)) or None  # force a limit if auto-detect fails
MONITOR_INTERVAL = 2          # seconds between diagnostic samples
BLOCK_IMAGES = True           # set False to test whether request interception matters

# ---- PMT API (discovered from the captured network log) ----
API_HOST = "pmt-api.komdigi.go.id"
# the request fired by Export > All Page:  POST https://pmt-api.komdigi.go.id/api/v2/en/alarm
API_EXPORT_URL_RE = re.compile(r"^https://pmt-api\.komdigi\.go\.id/api/v\d+/[^/]+/alarm/?(\?.*)?$")
API_DATE_FIELD = None         # e.g. "created_at". None = auto-detect (see Diagnostics after Step 1)
API_RECORDS_PATH = None       # e.g. "data.item". None = auto-detect
COLUMN_LABELS = {}            # e.g. {"created_at": "Data Created"} to rename Excel headers
SITE_UTC_OFFSET_HOURS = 7     # month boundaries are evaluated in WIB (UTC+7)
MONTHS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
# =========================================================

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger("pmt_alarm")
faulthandler.enable()  # prints a Python traceback to the app log on fatal signals (SIGSEGV, SIGABRT...)

_INSTALL_LOCK = threading.Lock()


# =====================================================================
#  Job folder, status files and log (shared by the UI and the workers)
# =====================================================================
def write_json_atomic(path: Path, data: dict):
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text(json.dumps(data), encoding="utf-8")
    os.replace(tmp, path)


def job_log(job_dir: Path, msg: str, level: str = "INFO"):
    line = f"{datetime.now():%H:%M:%S} [{level}] {msg}"
    try:
        with open(job_dir / "job.log", "a", encoding="utf-8") as fh:
            fh.write(line + "\n")
    except OSError:
        pass
    log.log(getattr(logging, level, logging.INFO), "[%s] %s", job_dir.name, msg)


class StepReporter:
    """Writes step1.json / step2.json so the UI can poll progress."""

    def __init__(self, job_dir: Path, step: int):
        self.job_dir, self.step, self.t0 = job_dir, step, time.time()

    def _write(self, state, message, progress=None, **extra):
        write_json_atomic(
            self.job_dir / f"step{self.step}.json",
            {"state": state, "message": message, "progress": progress,
             "started": self.t0, "updated": time.time(), **extra},
        )

    def running(self, message, progress=None, quiet=False):
        if not quiet:  # quiet = heartbeat only, don't spam the activity log
            job_log(self.job_dir, f"[step {self.step}] {message}")
        self._write("running", message, progress)

    def done(self, message, **extra):
        job_log(self.job_dir, f"[step {self.step}] DONE: {message}")
        self._write("done", message, 1.0, **extra)

    def fail(self, message):
        job_log(self.job_dir, f"[step {self.step}] FAILED: {message}", "ERROR")
        self._write("error", message)


def get_status(job_dir: Path, step: int) -> dict:
    path = job_dir / f"step{step}.json"
    if not path.exists():
        return {"state": "idle", "message": ""}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {"state": "idle", "message": ""}
    if data.get("state") == "running" and time.time() - data.get("updated", 0) > STALE_AFTER:
        data["state"] = "error"
        data["message"] = "This step was interrupted (the app may have restarted). Please run it again."
    return data


def cleanup_old_jobs():
    if not JOBS_ROOT.exists():
        return
    cutoff = time.time() - JOB_MAX_AGE_HOURS * 3600
    for d in JOBS_ROOT.iterdir():
        try:
            if d.is_dir() and d.stat().st_mtime < cutoff:
                shutil.rmtree(d, ignore_errors=True)
        except OSError:
            pass


def source_path(job_dir: Path) -> Path:
    return job_dir / "source.json"


def source_ready(job_dir: Path) -> bool:
    p = source_path(job_dir)
    return p.exists() and p.stat().st_size > 0


# =====================================================================
#  Diagnostics: container memory, process tree, disk, browser events
# =====================================================================
def _read_int(path):
    try:
        raw = Path(path).read_text().strip()
        return None if raw == "max" else int(raw)
    except Exception:
        return None


def _mb(v):
    return None if v is None else v / 1024 / 1024


def cgroup_memory() -> dict:
    """Container memory from cgroup v2 (or v1). This is the number the platform enforces."""
    used = _read_int("/sys/fs/cgroup/memory.current")
    limit = _read_int("/sys/fs/cgroup/memory.max")
    peak = _read_int("/sys/fs/cgroup/memory.peak")
    if used is None:  # cgroup v1
        used = _read_int("/sys/fs/cgroup/memory/memory.usage_in_bytes")
        limit = _read_int("/sys/fs/cgroup/memory/memory.limit_in_bytes")
        peak = _read_int("/sys/fs/cgroup/memory/memory.max_usage_in_bytes")
    if limit is not None and limit > (1 << 40):  # v1 "unlimited"
        limit = None
    oom = None
    try:
        for line in Path("/sys/fs/cgroup/memory.events").read_text().splitlines():
            if line.startswith("oom_kill "):
                oom = int(line.split()[1])
    except Exception:
        pass
    return {"used_mb": _mb(used), "limit_mb": MEM_LIMIT_MB_OVERRIDE or _mb(limit),
            "peak_mb": _mb(peak), "oom_kills": oom}


def process_tree():
    """[(pid, role, status, rss_mb)] for this process and every child (node driver, Chromium...)."""
    out = []
    if psutil is None:
        return out
    try:
        me = psutil.Process(os.getpid())
        for p in [me] + me.children(recursive=True):
            try:
                cmd = " ".join(p.cmdline())
                name = p.name()
                m = re.search(r"--type=([\w-]+)", cmd)
                if p.pid == me.pid:
                    role = "streamlit"
                elif name.startswith("node"):
                    role = "playwright-node"
                elif "chrom" in name.lower() or "chrom" in cmd.lower():
                    role = "chromium-" + (m.group(1) if m else "browser")
                else:
                    role = name
                out.append((p.pid, role, p.status(), p.memory_info().rss / 1048576))
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
    except Exception:
        pass
    return out


def kill_browser_processes(job_dir: Path):
    """Last-resort: kill the node driver + Chromium so the container itself is not OOM-killed."""
    if psutil is None:
        return
    try:
        for p in psutil.Process(os.getpid()).children(recursive=True):
            try:
                p.kill()
            except Exception:
                pass
        job_log(job_dir, "Killed all browser child processes (memory guard).", "ERROR")
    except Exception as exc:
        job_log(job_dir, f"Could not kill browser processes: {exc}", "ERROR")


def collect_snapshot(dl_dir: Path) -> dict:
    cg = cgroup_memory()
    tree = process_tree()
    chrom = [t for t in tree if t[1].startswith("chromium")]
    used = cg["used_mb"]
    if used is None and tree:  # no cgroup files: approximate with process RSS
        used = sum(t[3] for t in tree)
    files = []
    try:
        files = [(p.name, p.stat().st_size / 1048576) for p in dl_dir.rglob("*") if p.is_file()]
    except OSError:
        pass
    try:
        disk_tmp = shutil.disk_usage(tempfile.gettempdir()).free / 1048576
    except OSError:
        disk_tmp = None
    try:
        disk_shm = shutil.disk_usage("/dev/shm").free / 1048576
    except OSError:
        disk_shm = None
    return {
        "used": used, "limit": cg["limit_mb"], "peak": cg["peak_mb"], "oom": cg["oom_kills"],
        "chrom_n": len(chrom), "chrom_rss": sum(t[3] for t in chrom),
        "chrom_top": sorted(chrom, key=lambda t: -t[3])[:4],
        "tree": tree, "files": files,
        "dl_mb": max((f[1] for f in files), default=0.0),
        "disk_tmp": disk_tmp, "disk_shm": disk_shm,
    }


def format_snapshot(s: dict) -> str:
    def f(v, unit="MB"):
        return "?" if v is None else f"{v:.0f}{unit}"

    pct = f" ({s['used'] / s['limit']:.0%})" if s["used"] is not None and s["limit"] else ""
    top = ", ".join(f"{t[1].replace('chromium-', '')} {t[3]:.0f}" for t in s["chrom_top"]) or "none"
    return (
        f"mem {f(s['used'])}/{f(s['limit'])}{pct} peak={f(s['peak'])} oom_kills={s['oom']} | "
        f"chromium procs={s['chrom_n']} rss={f(s['chrom_rss'])} [{top}] | "
        f"download files={len(s['files'])} biggest={s['dl_mb']:.1f}MB | "
        f"free: tmp={f(s['disk_tmp'])} shm={f(s['disk_shm'])}"
    )


def log_system_info(job_dir: Path):
    try:
        from importlib.metadata import version

        pw_version = version("playwright")
    except Exception:
        pw_version = "?"
    cg = cgroup_memory()
    total = f"{psutil.virtual_memory().total / 1048576:.0f}MB" if psutil else "?"
    job_log(
        job_dir,
        f"SYSTEM: python={sys.version.split()[0]} playwright={pw_version} cpus={os.cpu_count()} "
        f"host_RAM={total} container_limit={cg['limit_mb'] and round(cg['limit_mb'])}MB "
        f"chromium={shutil.which('chromium') or shutil.which('chromium-browser')} psutil={'yes' if psutil else 'NO'}",
    )


class DiagnosticsMonitor(threading.Thread):
    """
    Samples container memory / process tree / download dir every few seconds, writes them to the log
    (stdout + job.log), refreshes the progress message, and kills Chromium before the container hits
    its memory limit so the app survives and shows a clear error.
    """

    def __init__(self, job_dir: Path, dl_dir: Path, rep: StepReporter):
        super().__init__(daemon=True)
        self.job_dir, self.dl_dir, self.rep = job_dir, dl_dir, rep
        self.stop_evt = threading.Event()
        self.t0 = time.time()
        self.phase = "starting"
        self.tripped = None
        self.peak_used = 0.0
        self.peak_chromium = 0.0
        self._seen_chromium = False
        self._gone_logged = False
        self._last_logged = 0.0
        self._last_used = 0.0

    def run(self):
        while not self.stop_evt.wait(MONITOR_INTERVAL):
            try:
                self._tick()
            except Exception as exc:
                log.warning("monitor error: %s", exc)

    def stop(self):
        self.stop_evt.set()

    def _tick(self):
        now = time.time()
        elapsed = now - self.t0
        s = collect_snapshot(self.dl_dir)
        used = s["used"] or 0.0
        self.peak_used = max(self.peak_used, used)
        self.peak_chromium = max(self.peak_chromium, s["chrom_rss"])

        if s["chrom_n"]:
            self._seen_chromium = True
        elif self._seen_chromium and not self._gone_logged:
            self._gone_logged = True
            job_log(self.job_dir, "!!! ALL CHROMIUM PROCESSES ARE GONE (browser exited or was killed). "
                                  f"Last memory state: {format_snapshot(s)}", "ERROR")

        # log every tick for the first 90 s, then every 10 s, and on any +100 MB jump
        if elapsed < 90 or now - self._last_logged >= 10 or used - self._last_used >= 100:
            self._last_logged, self._last_used = now, used
            job_log(self.job_dir, f"MON +{int(elapsed)}s [{self.phase}] {format_snapshot(s)}")

        if self.phase == "api-download" and not self.tripped:
            mins, secs = divmod(int(elapsed), 60)
            lim = f"/{s['limit']:.0f}" if s["limit"] else ""
            self.rep.running(
                f"Downloading data from the API (server may take several minutes to respond)... {mins}m {secs:02d}s elapsed, "
                f"{s['dl_mb']:.1f} MB received · memory {used:.0f}{lim} MB",
                0.50, quiet=True,
            )

        limit = s["limit"]
        if limit and used > limit * MEM_GUARD_FRACTION and not self.tripped:
            self.tripped = (
                f"Memory guard stopped Chromium: the container reached {used:.0f} MB of its {limit:.0f} MB limit "
                f"({used / limit:.0%}) during the export. Chromium used {s['chrom_rss']:.0f} MB."
            )
            job_log(self.job_dir, f"!!! {self.tripped}", "ERROR")
            for pid, role, status, rss in sorted(s["tree"], key=lambda t: -t[3])[:8]:
                job_log(self.job_dir, f"    pid={pid} {role} status={status} rss={rss:.0f}MB", "ERROR")
            kill_browser_processes(self.job_dir)


def attach_diagnostics(job_dir: Path, browser, page):
    """Log Chromium crashes, console errors, failed requests, and the network calls around the export."""
    counters = {"console": 0, "net": 0}

    def guarded(fn):
        def wrapper(*args):
            try:
                fn(*args)
            except Exception:
                pass
        return wrapper

    @guarded
    def on_console(msg):
        if msg.type in ("error", "warning") and counters["console"] < 200:
            counters["console"] += 1
            job_log(job_dir, f"[console.{msg.type}] {msg.text[:300]}", "WARNING")

    @guarded
    def on_response(resp):
        req = resp.request
        if req.resource_type not in ("xhr", "fetch", "document") or API_HOST in resp.url:
            return
        counters["net"] += 1
        n = counters["net"]
        if n <= 40 or n % 25 == 0:
            h = resp.headers
            job_log(job_dir, f"[net #{n}] {resp.status} {req.method} {req.resource_type} "
                             f"{resp.url.split('?')[0][:140]} len={h.get('content-length', '?')} "
                             f"type={h.get('content-type', '?')[:40]}")

    browser.on("disconnected", guarded(lambda *_: job_log(
        job_dir, "!!! BROWSER DISCONNECTED: the Chromium process exited or crashed", "ERROR")))
    page.on("crash", guarded(lambda *_: job_log(
        job_dir, "!!! PAGE CRASHED: the renderer process died (usually out of memory)", "ERROR")))
    page.on("close", guarded(lambda *_: job_log(job_dir, "[page] closed", "WARNING")))
    page.on("console", on_console)
    page.on("pageerror", guarded(lambda err: job_log(job_dir, f"[pageerror] {str(err)[:300]}", "WARNING")))
    page.on("requestfailed", guarded(lambda r: job_log(
        job_dir, f"[requestfailed] {r.method} {r.url[:150]} -> {r.failure}", "WARNING")))
    page.on("response", on_response)
    page.on("download", guarded(lambda d: job_log(
        job_dir, f"[download event] file={d.suggested_filename} url={d.url[:150]}")))
    page.on("popup", guarded(lambda p: job_log(job_dir, f"[popup] {p.url[:150]}", "WARNING")))


def take_shot(page, job_dir: Path, name: str):
    try:
        page.screenshot(path=str(job_dir / f"{name}.png"), timeout=15000)
        job_log(job_dir, f"Screenshot saved: {name}.png")
    except Exception as exc:
        job_log(job_dir, f"Screenshot {name} failed: {str(exc)[:150]}", "WARNING")


# =====================================================================
#  Playwright helpers
# =====================================================================
def launch_browser(pw, job_dir: Path, dl_dir: Path):
    """
    Launch headless Chromium.
    1) system Chromium (Streamlit Cloud via packages.txt) if available,
    2) otherwise Playwright's own Chromium, installing it on first use.
    """
    args = [
        "--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu", "--mute-audio",
        # fewer processes / less memory
        "--renderer-process-limit=1", "--no-zygote",
        "--disable-features=IsolateOrigins,site-per-process", "--disable-site-isolation-trials",
        "--disable-extensions", "--disable-background-networking", "--disable-sync",
        "--disable-component-update", "--disable-breakpad", "--metrics-recording-only", "--no-first-run",
    ]
    limit = cgroup_memory()["limit_mb"]
    if limit:  # make a runaway page die inside Chromium (page crash) instead of OOM-killing the container
        args.append(f"--js-flags=--max-old-space-size={int(max(256, limit * 0.45))}")

    kwargs = dict(headless=True, downloads_path=str(dl_dir), args=args)

    system_chromium = shutil.which("chromium") or shutil.which("chromium-browser")
    if system_chromium:
        try:
            browser = pw.chromium.launch(executable_path=system_chromium, **kwargs)
            job_log(job_dir, f"Using system Chromium {system_chromium} (version {browser.version})")
            return browser
        except PWError as exc:
            job_log(job_dir, f"System Chromium failed ({str(exc)[:300]}); trying Playwright's own.", "WARNING")

    try:
        browser = pw.chromium.launch(**kwargs)
        job_log(job_dir, f"Using Playwright Chromium (version {browser.version})")
        return browser
    except PWError as exc:
        if "playwright install" not in str(exc) and "Executable doesn't exist" not in str(exc):
            raise
    with _INSTALL_LOCK:
        job_log(job_dir, "Installing Playwright Chromium (first run only, may take a minute)...")
        proc = subprocess.run(
            [sys.executable, "-m", "playwright", "install", "chromium"],
            capture_output=True, text=True, timeout=900,
        )
        if proc.returncode != 0:
            raise RuntimeError(f"'playwright install chromium' failed: {proc.stderr[-500:] or proc.stdout[-500:]}")
    return pw.chromium.launch(**kwargs)


def click_first(page, selectors, description, total_timeout=DEFAULT_TIMEOUT):
    """Try several selectors in turn; click the first visible one (JS-click fallback)."""
    per = max(5000, int(total_timeout * 1000 / max(1, len(selectors))))
    last_exc = None
    for sel in selectors:
        loc = page.locator(sel).first
        try:
            loc.wait_for(state="visible", timeout=per)
            loc.scroll_into_view_if_needed(timeout=per)
            try:
                loc.click(timeout=per)
            except PWError:
                loc.evaluate("el => el.click()")
            return
        except PWError as exc:
            last_exc = exc
            log.warning("Selector failed for %s: %s", description, sel)
    raise RuntimeError(f"Could not click {description}: {str(last_exc)[:200]}")


UPPER = "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ')"


def pw_login(page, username, password):
    page.goto(LOGIN_URL, wait_until="load", timeout=60000)
    user = page.locator(
        "input[placeholder='Username'], input[name='username'], input#username"
    ).first
    user.wait_for(state="visible")
    user.fill(username)
    pwd = page.locator("input[type='password']").first
    pwd.wait_for(state="visible")
    pwd.fill(password)

    click_first(
        page,
        [
            "xpath=//button[normalize-space()='LOGIN' or normalize-space()='Login']",
            "xpath=//button[contains(translate(., 'login', 'LOGIN'), 'LOGIN')]",
            "xpath=//input[@type='submit']",
            "css=button[type='submit']",
        ],
        "LOGIN button",
    )
    try:
        page.wait_for_url(lambda url: "/auth/login" not in url, timeout=60000)
    except PWTimeout:
        raise RuntimeError(
            "Login failed: still on the login page. Check username/password, "
            "or the site may require CAPTCHA/OTP or block this server's IP."
        )


def pw_open_alarm_page(page):
    page.goto(ALARM_URL, wait_until="load", timeout=60000)
    if "/auth/login" in page.url:
        raise RuntimeError("Redirected back to login page; session not authenticated.")
    page.wait_for_selector("body", state="attached")
    page.wait_for_timeout(2000)


def pw_scroll_to_alarm_section(page):
    section = page.locator(
        "xpath=//*[self::h1 or self::h2 or self::h3 or self::h4 or self::h5 or self::div or self::span]"
        "[normalize-space()='Alarm' or normalize-space()='ALARM']"
    ).first
    try:
        section.wait_for(state="attached")
        section.evaluate("el => el.scrollIntoView({block: 'start'})")
    except PWError:
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
    page.wait_for_timeout(1000)


def pw_click_closed_alarm_tab(page):
    click_first(
        page,
        [
            "xpath=//*[(self::button or self::a or self::li or self::div or self::span or @role='tab')]"
            "[normalize-space()='CLOSED ALARM' or normalize-space()='Closed Alarm']",
            f"xpath=//*[contains({UPPER}, 'CLOSED ALARM') and (self::button or self::a or @role='tab' or self::li)]",
        ],
        "CLOSED ALARM tab",
    )
    page.wait_for_selector("table", state="attached", timeout=60000)
    try:
        page.wait_for_selector("table tbody tr", state="attached", timeout=DEFAULT_TIMEOUT * 1000)
    except PWTimeout:
        pass
    try:
        page.wait_for_selector(
            ".spinner, .loading, .loader, .spinner-border, [class*='loading'], [class*='spinner']",
            state="hidden", timeout=10000,
        )
    except PWError:
        pass
    page.wait_for_timeout(1000)


def pw_open_download_menu(page):
    click_first(
        page,
        [
            f"xpath=//button[contains({UPPER}, 'DOWNLOAD ALARM')]",
            f"xpath=//*[contains({UPPER}, 'DOWNLOAD ALARM') and (self::a or self::button or @role='button')]",
            "css=button.btn-success",
        ],
        "Download Alarm button",
    )


def pw_click_all_page(page):
    click_first(
        page,
        [
            "xpath=//*[(self::a or self::button or self::li or self::span or self::div "
            "or @role='menuitem') and (normalize-space()='All Page' or normalize-space()='All page')]",
            f"xpath=//*[contains({UPPER}, 'ALL PAGE') and (self::a or self::button or self::li or @role='menuitem')]",
        ],
        "All Page option",
    )


# =====================================================================
#  API discovery + capture (replaces the browser export)
# =====================================================================
SENSITIVE_RE = re.compile(r"pass|pwd|secret|otp|token", re.I)


def redact_headers(headers: dict) -> dict:
    out = {}
    for k, v in headers.items():
        lk, v = k.lower(), str(v)
        if lk in ("authorization", "cookie", "proxy-authorization") or "token" in lk or "api-key" in lk:
            out[k] = f"{v[:12]}…(len {len(v)})"
        else:
            out[k] = v[:150]
    return out


def redact_payload(body) -> str:
    if not body:
        return ""
    text = body.decode("utf-8", "replace") if isinstance(body, (bytes, bytearray)) else str(body)

    def red(o):
        if isinstance(o, dict):
            return {k: ("***" if SENSITIVE_RE.search(str(k)) else red(v)) for k, v in o.items()}
        if isinstance(o, list):
            return [red(x) for x in o[:20]]
        return o

    try:
        return json.dumps(red(json.loads(text)), ensure_ascii=False)[:1500]
    except ValueError:
        return re.sub(r"(?i)([^&=]*(?:pass|pwd|secret|token|otp)[^&=]*=)[^&]*", r"\1***", text)[:1500]


def find_records_path(obj, max_depth=6):
    """Return (path, records): the biggest list of dicts inside a JSON document."""
    best = ([], [])

    def walk(o, path, depth):
        nonlocal best
        if depth > max_depth:
            return
        if isinstance(o, list):
            if o and isinstance(o[0], dict) and len(o) > len(best[1]):
                best = (list(path), o)
        elif isinstance(o, dict):
            for k, v in o.items():
                walk(v, path + [str(k)], depth + 1)

    walk(obj, [], 0)
    return best


def prefix_from_path(path) -> str:
    return ".".join(list(path) + ["item"])


def describe_json(obj) -> dict:
    info = {"type": type(obj).__name__}
    if isinstance(obj, dict):
        info["top_keys"] = list(obj.keys())[:30]
        info["nested_keys"] = {k: list(v.keys())[:15] for k, v in list(obj.items())[:10] if isinstance(v, dict)}
    path, recs = find_records_path(obj)
    if recs:
        first = recs[0]
        info["records_path"] = ".".join(path) or "(root list)"
        info["records_in_response"] = len(recs)
        info["fields"] = list(first.keys())[:80]
        info["sample_record"] = {k: str(v)[:40] for k, v in list(first.items())[:80]}
    return info


def url_key(url: str) -> str:
    return url.split("?")[0]


class ApiCapture:
    """
    Records the site's API traffic (redacted) and, once armed, CAPTURES the export request and
    ABORTS it inside the browser so Chromium never receives (or renders) the huge response.
    The real headers/body are kept in memory only, never written to disk or logs.
    """

    def __init__(self, job_dir: Path):
        self.job_dir = job_dir
        self.export_armed = False
        self.export_requests = []   # real data, in memory only
        self.calls = []             # redacted, saved to api_capture.json
        self.structures = {}        # url_key -> {"prefix":..., "describe":...}
        self.login = None
        self.last_listing = None
        self._n = 0

    # ---- wiring
    def attach(self, page):
        page.route(API_EXPORT_URL_RE, self._on_route)
        page.on("request", self._guard(self._on_request))
        page.on("response", self._guard(self._on_response))

    @staticmethod
    def _guard(fn):
        def wrapper(*args):
            try:
                fn(*args)
            except Exception as exc:
                log.debug("api-capture handler error: %s", exc)
        return wrapper

    def _record(self, kind, req, body):
        rec = {
            "kind": kind, "method": req.method, "url": req.url, "resource_type": req.resource_type,
            "headers": redact_headers(req.headers), "payload": redact_payload(body),
            "payload_bytes": len(body) if body else 0,
        }
        if len(self.calls) < 150:
            self.calls.append(rec)
        return rec

    # ---- handlers
    def _on_route(self, route, *_):
        req = route.request
        if self.export_armed:
            body = req.post_data_buffer
            self.export_requests.append(
                {"method": req.method, "url": req.url, "headers": dict(req.headers), "body": body}
            )
            rec = self._record("EXPORT", req, body)
            job_log(self.job_dir, f"[EXPORT-REQUEST CAPTURED] {req.method} {url_key(req.url)} "
                                  f"payload_bytes={rec['payload_bytes']} payload={rec['payload'][:600]} "
                                  f"headers={list(rec['headers'].keys())}")
            route.abort()  # Chromium must NOT process the multi-hundred-MB response
        else:
            route.fallback()

    def _on_request(self, req):
        if API_HOST not in req.url or req.resource_type not in ("xhr", "fetch"):
            return
        if self.export_armed and API_EXPORT_URL_RE.match(req.url):
            return  # logged by _on_route
        body = req.post_data_buffer if req.method.upper() != "GET" else None
        self._n += 1
        rec = self._record("api", req, body)
        if API_EXPORT_URL_RE.match(req.url):
            self.last_listing = rec
        if re.search(r"login|signin|sign-in|auth|token", req.url, re.I) and req.method.upper() == "POST":
            self.login = rec
        if self._n <= 60:
            job_log(self.job_dir, f"[api-req #{self._n}] {req.method} {url_key(req.url)} "
                                  f"auth={'yes' if 'authorization' in req.headers else 'no'} "
                                  f"payload={rec['payload'][:300]}")

    def _on_response(self, resp):
        req = resp.request
        if API_HOST not in resp.url or req.resource_type not in ("xhr", "fetch"):
            return
        ctype = resp.headers.get("content-type", "")
        length = int(resp.headers.get("content-length", "0") or 0)
        job_log(self.job_dir, f"[api-resp] {resp.status} {req.method} {url_key(resp.url)} "
                              f"type={ctype[:30]} len={length}")
        if "json" in ctype and length < 2_000_000 and len(self.structures) < 15:
            data = json.loads(resp.body())
            info = describe_json(data)
            path, recs = find_records_path(data)
            entry = {"describe": info, "prefix": prefix_from_path(path) if recs else None}
            self.structures.setdefault(url_key(resp.url), entry)
            if self.login is not None and req.url == self.login["url"]:
                entry["describe"].pop("sample_record", None)  # never keep login response values
            job_log(self.job_dir, f"[api-struct] {url_key(resp.url)} top_keys={info.get('top_keys')} "
                                  f"records_path={info.get('records_path')} "
                                  f"records={info.get('records_in_response')} fields={info.get('fields')}")

    # ---- output
    def save(self):
        out = {
            "login_request": self.login,
            "listing_request_before_export": self.last_listing,
            "export_request": next((c for c in self.calls if c["kind"] == "EXPORT"), None),
            "structures": self.structures,
            "all_calls": self.calls,
        }
        (self.job_dir / "api_capture.json").write_text(json.dumps(out, indent=2, ensure_ascii=False), encoding="utf-8")


# =====================================================================
#  Direct API download (no browser)
# =====================================================================
DROP_HEADERS = {"host", "content-length", "connection", "accept-encoding", "cookie",
                "transfer-encoding", "keep-alive", "upgrade-insecure-requests"}


def build_replay_headers(export: dict) -> dict:
    headers = {k: v for k, v in export["headers"].items()
               if k.lower() not in DROP_HEADERS and not k.startswith(":") and not k.lower().startswith("sec-")}
    present = {k.lower() for k in headers}
    origin = "{0.scheme}://{0.netloc}".format(urlparse(ALARM_URL))
    if "user-agent" not in present:
        headers["User-Agent"] = export.get("user_agent") or "Mozilla/5.0 (X11; Linux x86_64) Chrome/124.0 Safari/537.36"
    if "origin" not in present:
        headers["Origin"] = origin
    if "referer" not in present:
        headers["Referer"] = ALARM_URL
    headers["Accept-Encoding"] = "gzip, deflate"
    cookies = export.get("cookies") or []
    if cookies:
        headers["Cookie"] = "; ".join(f"{c['name']}={c['value']}" for c in cookies)
    return headers


def download_via_api(job_dir: Path, export: dict, dest: Path, monitor) -> int:
    """Replay the captured export request with `requests` and stream the body to disk."""
    headers = build_replay_headers(export)
    job_log(job_dir, f"API replay: {export['method']} {url_key(export['url'])} "
                     f"headers={redact_headers(headers)} body_bytes={len(export['body'] or b'')}")
    with requests.Session() as sess:
        resp = sess.request(
            export["method"], export["url"], headers=headers, data=export["body"],
            stream=True, timeout=(30, DOWNLOAD_TIMEOUT),
        )
        job_log(job_dir, f"API replay response: HTTP {resp.status_code} "
                         f"type={resp.headers.get('content-type')} length={resp.headers.get('content-length')} "
                         f"encoding={resp.headers.get('content-encoding')}")
        if resp.status_code >= 400:
            snippet = next(resp.iter_content(1024), b"")[:300].decode("utf-8", "replace")
            hint = (" The token/cookie captured from the browser was not accepted; "
                    "see api_capture.json for the headers that were sent."
                    if resp.status_code in (401, 403) else "")
            raise RuntimeError(f"API rejected the export request (HTTP {resp.status_code}): {snippet}{hint}")

        written, first = 0, b""
        with open(dest, "wb") as fh:
            for chunk in resp.iter_content(chunk_size=1 << 20):
                if monitor is not None and monitor.tripped:
                    raise RuntimeError(monitor.tripped)
                if not chunk:
                    continue
                if not first:
                    first = chunk[:200]
                fh.write(chunk)
                written += len(chunk)

    if not first.lstrip(b"\xef\xbb\xbf \r\n\t")[:1] in (b"{", b"["):
        raise RuntimeError("The API did not return JSON. First bytes: "
                           f"{first[:120]!r} (content-type {resp.headers.get('content-type')}).")
    return written


def detect_prefix_from_file(path: Path):
    """Fallback: first array of objects found while streaming the file."""
    last_array = None
    with open(path, "rb") as fh:
        for i, (prefix, event, _value) in enumerate(ijson.parse(fh)):
            if event == "start_array":
                last_array = prefix
            elif event == "start_map" and last_array is not None and prefix == (f"{last_array}.item" if last_array else "item"):
                return prefix
            elif i > 3_000_000:
                break
    return None


# =====================================================================
#  STEP 1 worker – capture the export request, then download via API
# =====================================================================
def retrieve_worker(job_dir: Path, username: str, password: str):
    rep = StepReporter(job_dir, 1)
    dl_dir = job_dir / "dl"
    part = dl_dir / "alarm_export.json.part"
    shutil.rmtree(dl_dir, ignore_errors=True)
    dl_dir.mkdir(parents=True, exist_ok=True)
    for pattern in ("shot_*.png", "api_capture.json", "source_meta.json"):
        for old in job_dir.glob(pattern):
            old.unlink(missing_ok=True)

    monitor = None
    export = None
    capture = ApiCapture(job_dir)
    try:
        log_system_info(job_dir)
        rep.running("Starting browser...", 0.05)
        monitor = DiagnosticsMonitor(job_dir, dl_dir, rep)
        monitor.start()

        # ---------------- browser: login + capture only ----------------
        with sync_playwright() as pw:
            browser = launch_browser(pw, job_dir, dl_dir)
            page = None
            try:
                context = browser.new_context(accept_downloads=False, viewport={"width": 1600, "height": 900})
                if BLOCK_IMAGES:
                    context.route(
                        "**/*",
                        lambda route: route.abort()
                        if route.request.resource_type in ("image", "media")
                        else route.continue_(),
                    )
                page = context.new_page()
                page.set_default_timeout(DEFAULT_TIMEOUT * 1000)
                attach_diagnostics(job_dir, browser, page)
                capture.attach(page)

                monitor.phase = "login"
                rep.running("Logging in...", 0.15)
                pw_login(page, username, password)

                monitor.phase = "alarm-page"
                rep.running("Opening alarm page...", 0.25)
                pw_open_alarm_page(page)
                pw_scroll_to_alarm_section(page)

                rep.running("Opening CLOSED ALARM tab...", 0.30)
                pw_click_closed_alarm_tab(page)
                job_log(job_dir, "=== BEFORE export click === " + format_snapshot(collect_snapshot(dl_dir)))
                take_shot(page, job_dir, "shot_1_before_export")

                monitor.phase = "capture"
                rep.running("Capturing the export API request (the heavy browser export is blocked)...", 0.35)
                pw_open_download_menu(page)
                capture.export_armed = True       # from now on the alarm API call is captured + aborted
                pw_click_all_page(page)

                end = time.time() + 60
                while not capture.export_requests and time.time() < end:
                    if not browser.is_connected() or page.is_closed():
                        raise RuntimeError("Chromium exited while capturing the export request.")
                    page.wait_for_timeout(500)
                if not capture.export_requests:
                    raise RuntimeError("Clicking Export > All Page did not send a request to the alarm API. "
                                       "Check the [api-req] lines and api_capture.json in Diagnostics.")
                page.wait_for_timeout(1500)  # collect any additional calls triggered by the click
                job_log(job_dir, f"Captured {len(capture.export_requests)} export request(s); using the first one.")
                export = capture.export_requests[0]
                export["cookies"] = context.cookies([export["url"]])
                try:
                    export["user_agent"] = page.evaluate("navigator.userAgent").replace("HeadlessChrome", "Chrome")
                except Exception:
                    export["user_agent"] = None
                take_shot(page, job_dir, "shot_2_after_export")
                job_log(job_dir, "=== AFTER capture === " + format_snapshot(collect_snapshot(dl_dir)))
            except Exception:
                job_log(job_dir, "State at failure: " + format_snapshot(collect_snapshot(dl_dir)), "ERROR")
                if page is not None:
                    try:
                        page.screenshot(path=str(job_dir / "error.png"), timeout=10000)
                    except Exception:
                        pass
                raise
            finally:
                capture.save()
                try:
                    browser.close()
                except Exception:
                    pass

        # ---------------- no browser from here on ----------------
        job_log(job_dir, "=== BROWSER CLOSED === " + format_snapshot(collect_snapshot(dl_dir)))
        monitor.phase = "api-download"
        rep.running("Browser closed. Downloading the data directly from the API...", 0.45)
        written = download_via_api(job_dir, export, part, monitor)
        export = None  # drop credentials from memory
        job_log(job_dir, f"API download finished: {written / 1048576:.1f} MB (decompressed JSON)")

        capture_url = capture_export_url(capture)
        structure = capture.structures.get(url_key(capture_url))
        prefix = API_RECORDS_PATH or (structure or {}).get("prefix") or detect_prefix_from_file(part)
        if not prefix:
            raise RuntimeError("Could not find the list of alarm records inside the API response. "
                               "See api_capture.json and set API_RECORDS_PATH.")
        job_log(job_dir, f"Records path: {prefix}")

        (job_dir / "source_meta.json").write_text(
            json.dumps({"records_prefix": prefix, "export_url": url_key(capture_url), "bytes": written}), encoding="utf-8")
        os.replace(part, source_path(job_dir))
        size_mb = source_path(job_dir).stat().st_size / 1024 / 1024
        rep.done(f"Download complete: {size_mb:.1f} MB of alarm data saved on the server.",
                 size_mb=round(size_mb, 1), records_prefix=prefix)
    except Exception as exc:
        job_log(job_dir, traceback.format_exc(), "ERROR")
        if monitor is not None and monitor.tripped:
            message = monitor.tripped
        else:
            message = str(exc).strip().splitlines()[0] if str(exc).strip() else type(exc).__name__
        rep.fail(message)
    finally:
        if monitor is not None:
            monitor.phase = "finished"
            monitor.stop()
            job_log(job_dir, f"SUMMARY: peak container memory={monitor.peak_used:.0f}MB, "
                             f"peak Chromium RSS={monitor.peak_chromium:.0f}MB, "
                             f"memory guard tripped={'YES' if monitor.tripped else 'no'}")
        shutil.rmtree(dl_dir, ignore_errors=True)


def capture_export_url(capture: ApiCapture) -> str:
    for c in capture.calls:
        if c["kind"] == "EXPORT":
            return c["url"]
    return ""


# =====================================================================
#  STEP 2 worker – streaming filter of the API JSON (low memory)
# =====================================================================
WIB = timezone(timedelta(hours=SITE_UTC_OFFSET_HOURS))
MAX_COLUMNS = 150
_DATE_FORMATS = [
    "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M:%S.%f", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
    "%d/%m/%Y %H:%M:%S", "%d-%m-%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d-%m-%Y %H:%M", "%d/%m/%Y", "%d-%m-%Y",
]


def _to_site_time(dt: datetime) -> datetime:
    return dt.astimezone(WIB).replace(tzinfo=None) if dt.tzinfo is not None else dt


class DateParser:
    """Row-by-row date parsing: ISO (with timezone), epoch s/ms, Excel serials, d/m/Y text."""

    def __init__(self):
        self.fmt = None

    def __call__(self, value):
        if value is None or isinstance(value, bool):
            return None
        if isinstance(value, datetime):
            return _to_site_time(value)
        if isinstance(value, date):
            return datetime(value.year, value.month, value.day)
        if isinstance(value, (int, float)):
            try:
                if value > 1e11:
                    return datetime.fromtimestamp(value / 1000, tz=WIB).replace(tzinfo=None)
                if value > 1e8:
                    return datetime.fromtimestamp(value, tz=WIB).replace(tzinfo=None)
                return from_excel(value)
            except Exception:
                return None
        s = str(value).strip()
        if not s:
            return None
        if self.fmt:
            try:
                return datetime.strptime(s, self.fmt)
            except ValueError:
                pass
        if "T" in s or s.endswith("Z"):
            try:
                return _to_site_time(datetime.fromisoformat(s.replace("Z", "+00:00")))
            except ValueError:
                pass
        for fmt in _DATE_FORMATS:
            try:
                parsed = datetime.strptime(s, fmt)
                self.fmt = fmt
                return parsed
            except ValueError:
                continue
        try:
            import pandas as pd

            iso = len(s) >= 5 and s[:4].isdigit() and s[4] in "-/"
            ts = pd.to_datetime(s, errors="coerce", dayfirst=not iso)
            return None if pd.isna(ts) else ts.to_pydatetime()
        except Exception:
            return None


def _clean(v):
    return ILLEGAL_CHARACTERS_RE.sub("", v) if isinstance(v, str) else v


def flatten(rec: dict, parent: str = "", out: dict = None) -> dict:
    if out is None:
        out = {}
    for k, v in rec.items():
        key = f"{parent}.{k}" if parent else str(k)
        if isinstance(v, dict):
            flatten(v, key, out)
        elif isinstance(v, list):
            out[key] = json.dumps(v, ensure_ascii=False)
        else:
            out[key] = v
    return out


_PREFERRED_DATE_FIELDS = ["created_at", "createdat", "date_created", "data_created", "created", "created_date",
                          "createddate", "datetime_created", "create_at", "created_time"]


def pick_date_field(columns, sample_rows, parser):
    if API_DATE_FIELD:
        if API_DATE_FIELD not in columns:
            raise RuntimeError(f"API_DATE_FIELD '{API_DATE_FIELD}' not found. Available fields: {columns[:60]}")
        return API_DATE_FIELD
    lower = {c: c.lower().replace(" ", "_") for c in columns}
    ordered = [c for pref in _PREFERRED_DATE_FIELDS for c in columns if lower[c].split(".")[-1] == pref]
    ordered += [c for c in columns if "creat" in lower[c] and c not in ordered]
    ordered += [c for c in columns if any(t in lower[c] for t in ("date", "time", "_at")) and c not in ordered]
    for col in ordered:
        vals = [r.get(col) for r in sample_rows if r.get(col) not in (None, "")][:200]
        if vals and sum(parser(v) is not None for v in vals) / len(vals) >= 0.8:
            return col
    raise RuntimeError("Could not detect the creation-date field automatically. "
                       f"Set API_DATE_FIELD at the top of app.py. Available fields: {columns[:60]}")


def stream_filter_json(src: Path, dst: Path, year: int, month: int, prefix: str, on_progress):
    """
    Stream records from the API JSON (ijson) -> keep the chosen Year-Month -> write XLSX row by row.
    Memory stays small even if the JSON is hundreds of MB.
    """
    parser = DateParser()
    with open(src, "rb") as fh:
        records = ijson.items(fh, prefix, use_float=True)
        head = [flatten(r) for r in islice(records, 500)]
        if not head:
            raise RuntimeError(f"No records found at path '{prefix}' in the downloaded JSON.")

        columns, seen = [], set()
        for row in head:
            for key in row:
                if key not in seen:
                    seen.add(key)
                    columns.append(key)
        columns = columns[:MAX_COLUMNS]
        date_field = pick_date_field(columns, head, parser)
        date_idx = columns.index(date_field)
        job_log_hint = f"columns={len(columns)} date_field={date_field}"

        wb_out = Workbook(write_only=True)
        ws_out = wb_out.create_sheet("Closed Alarm")
        ws_out.freeze_panes = "A2"
        labels = [COLUMN_LABELS.get(c, c) for c in columns]
        for i, label in enumerate(labels, start=1):
            ws_out.column_dimensions[get_column_letter(i)].width = min(max(len(label) + 4, 14), 40)
        ws_out.append(labels)

        total = kept = invalid = 0
        last_report = time.time()
        all_rows = chain(head, (flatten(r) for r in records))
        for flat in all_rows:
            total += 1
            dt = parser(flat.get(date_field))
            if dt is None:
                invalid += 1
            elif dt.year == year and dt.month == month:
                row = [_clean(flat.get(c)) for c in columns]
                row[date_idx] = dt
                ws_out.append(row)
                kept += 1
            if total % 2000 == 0 and time.time() - last_report >= 2:
                last_report = time.time()
                on_progress(total, kept)

        wb_out.save(dst)
    return {"total": total, "kept": kept, "invalid": invalid, "date_col": date_field, "info": job_log_hint}


def filter_worker(job_dir: Path, year: int, month: int):
    rep = StepReporter(job_dir, 2)
    label = f"{MONTHS[month - 1]} {year}"
    part = job_dir / "result.part"
    try:
        if not source_ready(job_dir):
            raise RuntimeError("Source data not found. Run Step 1 first.")
        meta = json.loads((job_dir / "source_meta.json").read_text(encoding="utf-8"))

        for old in job_dir.glob("result_*.xlsx"):
            old.unlink(missing_ok=True)

        rep.running(f"Reading source data and filtering for {label}...", None)

        def progress(total, kept):
            rep.running(f"Filtering {label}: {total:,} records scanned, {kept:,} kept so far...", None, quiet=True)

        stats = stream_filter_json(source_path(job_dir), part, year, month, meta["records_prefix"], progress)
        job_log(job_dir, f"[step 2] {stats['info']}")

        result_name = f"result_{year}-{month:02d}.xlsx"
        os.replace(part, job_dir / result_name)
        stats.pop("info", None)
        rep.done(
            f"Filtering complete for {label}: {stats['kept']:,} of {stats['total']:,} records kept "
            f"(date field: {stats['date_col']}).",
            result_file=result_name, label=label, year=year, month=month, **stats,
        )
    except Exception as exc:
        job_log(job_dir, traceback.format_exc(), "ERROR")
        rep.fail(str(exc) or type(exc).__name__)
    finally:
        part.unlink(missing_ok=True)


# =====================================================================
#  UI
# =====================================================================
def start_thread(target, *args):
    threading.Thread(target=target, args=args, daemon=True).start()


def get_job_id() -> str:
    jid = st.query_params.get("job", "")
    if not re.fullmatch(r"[a-f0-9]{12}", jid or ""):
        jid = uuid.uuid4().hex[:12]
        st.query_params["job"] = jid
    return jid


def fmt_elapsed(started):
    secs = int(time.time() - started)
    return f"{secs // 60}m {secs % 60:02d}s"


def show_status(status: dict):
    state = status.get("state", "idle")
    msg = status.get("message", "")
    if state == "running":
        if status.get("progress") is not None:
            st.progress(min(max(status["progress"], 0.0), 1.0))
        st.info(f"⏳ {msg}  \n_Total time: {fmt_elapsed(status.get('started', time.time()))}_")
    elif state == "done":
        st.success(f"✅ {msg}")
    elif state == "error":
        st.error(f"❌ {msg}")


def render_workflow(job_dir: Path, job_id: str):
    s1, s2 = get_status(job_dir, 1), get_status(job_dir, 2)
    busy = s1["state"] == "running" or s2["state"] == "running"

    # when a background step just finished, do one full rerun to stop polling
    was_busy = st.session_state.get("_was_busy", False)
    st.session_state["_was_busy"] = busy
    if was_busy and not busy:
        st.rerun()

    src_ok = source_ready(job_dir)

    # ------------------------------------------------------------ STEP 1
    st.subheader("Step 1 · Retrieve Excel")
    st.caption("Logs in, captures the Export API request (the heavy in-browser export is blocked), then downloads "
               "the data directly from the PMT API. No filtering happens here.")
    if src_ok and s1["state"] != "running":
        size_mb = source_path(job_dir).stat().st_size / 1024 / 1024
        st.success(f"✅ Source data is ready on the server ({size_mb:.1f} MB). No need to download again.")
    show_status(s1 if s1["state"] in ("running", "error") else {"state": "idle"})
    if s1["state"] == "error" and (job_dir / "error.png").exists():
        st.image(str(job_dir / "error.png"), caption="Browser screenshot when the error happened")

    with st.form("step1_form"):
        username = st.text_input("Username", disabled=busy)
        password = st.text_input("Password", type="password", disabled=busy)
        force = st.checkbox("Replace the existing file (download again)", disabled=busy or not src_ok)
        go1 = st.form_submit_button("Retrieve Excel", type="primary", disabled=busy, use_container_width=True)

    if go1:
        if src_ok and not force:
            st.info("The source data already exists, so the download was skipped. "
                    "Tick “Replace the existing file” to download a new copy.")
        elif not username or not password:
            st.error("Please enter both username and password.")
        else:
            # a new source invalidates old step-2 results
            source_path(job_dir).unlink(missing_ok=True)
            (job_dir / "step2.json").unlink(missing_ok=True)
            (job_dir / "error.png").unlink(missing_ok=True)
            for old in job_dir.glob("result_*.xlsx"):
                old.unlink(missing_ok=True)
            StepReporter(job_dir, 1).running("Queued...", 0.0)
            start_thread(retrieve_worker, job_dir, username, password)
            st.session_state["_was_busy"] = True
            st.rerun()

    st.divider()

    # ------------------------------------------------------------ STEP 2
    st.subheader("Step 2 · Filter Excel")
    st.caption("Streams the downloaded data and keeps only the chosen Year-Month, then builds the Excel file. Nothing is downloaded here.")
    ready2 = src_ok and s1["state"] != "running"

    now = datetime.now()
    years = list(range(now.year, 2019, -1))
    c1, c2 = st.columns(2)
    year = c1.selectbox("Year", years, key="sel_year", disabled=not ready2 or busy)
    month_name = c2.selectbox("Month", MONTHS, index=now.month - 1, key="sel_month", disabled=not ready2 or busy)

    if not ready2:
        st.caption("🔒 Complete Step 1 first.")
    go2 = st.button("Filter Excel", type="primary", disabled=(not ready2) or busy, use_container_width=True)
    if go2:
        for old in job_dir.glob("result_*.xlsx"):
            old.unlink(missing_ok=True)
        StepReporter(job_dir, 2).running("Queued...", 0.0)
        start_thread(filter_worker, job_dir, int(year), MONTHS.index(month_name) + 1)
        st.session_state["_was_busy"] = True
        st.rerun()

    show_status(s2)
    if s2["state"] == "done" and s2.get("invalid"):
        st.warning(f"{s2['invalid']:,} rows had an empty/unreadable date and were excluded.")

    st.divider()

    # ------------------------------------------------------------ STEP 3
    st.subheader("Step 3 · Download Result")
    st.caption("Just hands you the processed file from Step 2. No processing happens here.")
    result_path = job_dir / s2["result_file"] if s2.get("state") == "done" and s2.get("result_file") else None
    if result_path is not None and result_path.exists():
        st.download_button(
            "Download Result",
            data=result_path.read_bytes(),
            file_name=f"PMT_Alarm_Closed_{s2['year']}-{int(s2['month']):02d}.xlsx",
            mime=XLSX_MIME,
            type="primary",
            use_container_width=True,
            on_click="ignore",  # don't rerun the app when downloading
        )
        st.caption(f"{s2['label']} · {result_path.stat().st_size / 1024 / 1024:.2f} MB")
    else:
        st.button("Download Result", disabled=True, use_container_width=True, key="dl_disabled")
        st.caption("🔒 Complete Step 2 first.")

    # ------------------------------------------------------------ extras
    st.divider()
    with st.expander("Diagnostics (log, memory, screenshots)"):
        log_file = job_dir / "job.log"
        all_lines = log_file.read_text(encoding="utf-8").splitlines() if log_file.exists() else []
        st.code("\n".join(all_lines[-150:]) or "(empty)", language=None)
        if all_lines:
            st.download_button("Download full log (job.log)", data="\n".join(all_lines),
                               file_name=f"job_{job_id}.log", mime="text/plain", on_click="ignore", key="dl_log")
        cap_file = job_dir / "api_capture.json"
        if cap_file.exists():
            st.download_button("Download api_capture.json (redacted API discovery)", data=cap_file.read_bytes(),
                               file_name="api_capture.json", mime="application/json", on_click="ignore", key="dl_cap")
        for shot in sorted(job_dir.glob("shot_*.png")):
            st.image(str(shot), caption=shot.stem)
    if st.button("🗑️ Delete my files & start over", disabled=busy):
        shutil.rmtree(job_dir, ignore_errors=True)
        st.query_params.clear()
        st.session_state.clear()
        st.rerun()


def main():
    st.set_page_config(page_title="PMT Alarm Downloader", page_icon="🚨", layout="centered")
    st.title("🚨 PMT Alarm – Closed Alarm")
    cleanup_old_jobs()

    job_id = get_job_id()
    job_dir = JOBS_ROOT / job_id
    job_dir.mkdir(parents=True, exist_ok=True)
    st.caption(
        f"Session `{job_id}` – the page URL contains this id. If the page reloads or disconnects "
        "during a long step, reopen the same URL to continue."
    )

    busy = any(get_status(job_dir, s)["state"] == "running" for s in (1, 2))
    # poll only while a background step is running
    st.fragment(run_every=POLL_SECONDS if busy else None)(render_workflow)(job_dir, job_id)

    st.caption("Your credentials are used only to start Step 1 and are never written to disk or logs.")


if __name__ == "__main__":
    main()
