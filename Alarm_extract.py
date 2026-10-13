"""
PMT Alarm – Closed Alarm downloader (Playwright edition) – PER-PAGE CYCLE EDITION

WHY THIS VERSION
  The site resets "Items per page" (and possibly the dates) when it moves to the next page.
  So every page is now fetched in its OWN clean cycle:

    for every page N:
      1. reload the alarm page (re-login only if the session was lost)
      2. open the CLOSED ALARM tab
      3. type Start Date / End Date and click Apply
      4. set Items per page to 500 (waits until the table really shows it)
      5. go to page N (page-number buttons / next)  -- skipped for page 1
      6. verify the table really shows page N at full page size
      7. Download Alarm > By Page
      8. last page? -> stop, otherwise repeat for N+1

  All page files are compiled into ONE Excel file.
STEP 3  Download -> hand the compiled XLSX to the browser (no processing)

MEMORY MANAGEMENT (kept from the previous version)
  * Memory guard on the container WORKING SET (memory.current - inactive_file), kills ONLY Chromium.
  * Downloads are deleted from Playwright's own store after save_as.
  * JS handles are disposed, tiny JS signatures are used for change detection.
  * Browser recycling on memory pressure / crashes: because every page is a fresh cycle, a restart simply
    logs in again and continues with the next page (no skipping needed).
  * While a step is running the UI only shows the last 40 log lines.

All state lives on disk in a per-session job folder, so it survives websocket reconnects and page refreshes.
"""
import csv
import faulthandler
import hashlib
import io
import json
import logging
import math
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
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.utils import get_column_letter
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
DEFAULT_TIMEOUT = 40          # seconds for normal waits
JOBS_ROOT = Path(tempfile.gettempdir()) / "pmt_alarm_jobs"
STALE_AFTER = 600             # a "running" step with no heartbeat for 10 min = interrupted
JOB_MAX_AGE_HOURS = 12
POLL_SECONDS = 3
MONITOR_INTERVAL = 2          # seconds between diagnostic samples
BLOCK_IMAGES = True           # also blocks media + fonts
MEM_LIMIT_MB_OVERRIDE = float(os.environ.get("MEM_LIMIT_MB", 0)) or None

# ---- memory management -------------------------------------------------
MEM_GUARD_FRACTION = 0.93     # HARD guard: working set above 93 % of the limit ...
MEM_GUARD_SAMPLES = 3         # ... for this many consecutive samples (x MONITOR_INTERVAL s) kills Chromium
SOFT_RECYCLE_FRACTION = 0.80  # SOFT limit: after a page, recycle the browser if working set is above this
CHROMIUM_GROWTH_MB = 450      # ... or if Chromium RSS grew this much since the first page of the session
RECYCLE_MIN_PAGES = 3         # never recycle before this many pages were downloaded in a session
RECYCLE_EVERY_PAGES = 0       # >0 = also recycle the browser every N pages (0 = only on memory pressure)
MAX_RECYCLES = 200            # planned recycles (memory pressure) allowed per job
MAX_CRASH_RESTARTS = 6        # unplanned restarts (crash / guard trip / error) that made progress
PAGE_ATTEMPTS = 3             # how many times one page cycle is retried before the session fails
PROBE_FIRST_PAGES = 3         # verify the real row count inside the first N downloaded files ...
PROBE_EVERY = 25              # ... and then every N pages

SITE_UTC_OFFSET_HOURS = 7     # the site works in WIB (UTC+7)
DATE_INPUT_FORMAT = "%d-%m-%Y"        # how the Start/End Date boxes display dates (02-10-2026)
PAGE_SIZE_CHOICES = ("500", "250", "100", "50")  # tried in this order in the "Items" dropdown
DEFAULT_PAGE_ROWS = 10        # rows the site shows before a bigger page size is chosen
MAX_PAGES = 3000              # safety cap for the pagination loop
DOWNLOAD_TIMEOUT = 240        # seconds to wait for ONE "By Page" file
PAGE_SIZE_LOAD_TIMEOUT = 240  # seconds to wait for the table to grow after choosing a bigger page size
PAGE_SIZE_STABLE_SECONDS = 6  # a partial table (fewer rows than chosen size) must stay unchanged this long
APPLY_LOAD_TIMEOUT_FIRST = 90 # seconds to wait for the table to reload after Apply (first page)
APPLY_LOAD_TIMEOUT_NEXT = 25  # ... on later cycles (the table may legitimately look identical)
GOTO_STEP_TIMEOUT = 60        # seconds to wait for the table to change after ONE navigation click

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
# =========================================================

WIB = timezone(timedelta(hours=SITE_UTC_OFFSET_HOURS))
QUIET_PHASES = ("finished", "closing", "recycling", "compile", "idle")

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger("pmt_alarm")
faulthandler.enable()

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
    """Writes step1.json so the UI can poll progress."""

    def __init__(self, job_dir: Path, step: int):
        self.job_dir, self.step, self.t0 = job_dir, step, time.time()

    def _write(self, state, message, progress=None, **extra):
        write_json_atomic(
            self.job_dir / f"step{self.step}.json",
            {"state": state, "message": message, "progress": progress,
             "started": self.t0, "updated": time.time(), **extra},
        )

    def running(self, message, progress=None, quiet=False):
        if not quiet:
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


def _cg_stat() -> dict:
    """memory.stat of the container cgroup (v2 or v1)."""
    for p in ("/sys/fs/cgroup/memory.stat", "/sys/fs/cgroup/memory/memory.stat"):
        try:
            out = {}
            for line in Path(p).read_text().splitlines():
                parts = line.split()
                if len(parts) == 2:
                    try:
                        out[parts[0]] = int(parts[1])
                    except ValueError:
                        pass
            if out:
                return out
        except Exception:
            continue
    return {}


def cgroup_memory() -> dict:
    used = _read_int("/sys/fs/cgroup/memory.current")
    limit = _read_int("/sys/fs/cgroup/memory.max")
    peak = _read_int("/sys/fs/cgroup/memory.peak")
    if used is None:  # cgroup v1
        used = _read_int("/sys/fs/cgroup/memory/memory.usage_in_bytes")
        limit = _read_int("/sys/fs/cgroup/memory/memory.limit_in_bytes")
        peak = _read_int("/sys/fs/cgroup/memory/memory.max_usage_in_bytes")
    if limit is not None and limit > (1 << 40):
        limit = None
    oom = None
    try:
        for line in Path("/sys/fs/cgroup/memory.events").read_text().splitlines():
            if line.startswith("oom_kill "):
                oom = int(line.split()[1])
    except Exception:
        pass
    st_ = _cg_stat()
    inactive = st_.get("inactive_file", st_.get("total_inactive_file"))
    # working set = what the kernel can NOT simply drop (this is what really leads to an OOM kill)
    ws = used - inactive if (used is not None and inactive is not None) else used
    return {
        "used_mb": _mb(used), "ws_mb": _mb(ws),
        "limit_mb": MEM_LIMIT_MB_OVERRIDE or _mb(limit),
        "peak_mb": _mb(peak), "oom_kills": oom,
        "anon_mb": _mb(st_.get("anon", st_.get("total_rss"))),
        "file_mb": _mb(st_.get("file", st_.get("total_cache"))),
        "shmem_mb": _mb(st_.get("shmem", st_.get("total_shmem"))),
    }


def process_tree():
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
    """Kill ONLY Chromium processes (never the Playwright node driver)."""
    if psutil is None:
        return
    killed = 0
    try:
        for p in psutil.Process(os.getpid()).children(recursive=True):
            try:
                ident = (p.name() + " " + " ".join(p.cmdline())).lower()
                if "chrom" in ident and not p.name().startswith("node"):
                    p.kill()
                    killed += 1
            except Exception:
                pass
        job_log(job_dir, f"Killed {killed} Chromium process(es).", "ERROR")
    except Exception as exc:
        job_log(job_dir, f"Could not kill browser processes: {exc}", "ERROR")


def collect_snapshot(dl_dir: Path) -> dict:
    cg = cgroup_memory()
    tree = process_tree()
    chrom = [t for t in tree if t[1].startswith("chromium")]
    used = cg["used_mb"]
    if used is None and tree:
        used = sum(t[3] for t in tree)
    ws = cg["ws_mb"] if cg["ws_mb"] is not None else used
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
        "used": used, "ws": ws, "limit": cg["limit_mb"], "peak": cg["peak_mb"], "oom": cg["oom_kills"],
        "anon": cg["anon_mb"], "file": cg["file_mb"], "shmem": cg["shmem_mb"],
        "py_rss": sum(t[3] for t in tree if t[1] == "streamlit"),
        "node_rss": sum(t[3] for t in tree if t[1] == "playwright-node"),
        "chrom_n": len(chrom), "chrom_rss": sum(t[3] for t in chrom),
        "chrom_top": sorted(chrom, key=lambda t: -t[3])[:4],
        "tree": tree, "files": files,
        "dl_mb": max((f[1] for f in files), default=0.0),
        "disk_tmp": disk_tmp, "disk_shm": disk_shm,
    }


def format_snapshot(s: dict) -> str:
    def f(v, unit="MB"):
        return "?" if v is None else f"{v:.0f}{unit}"

    pct = f" ({s['ws'] / s['limit']:.0%})" if s.get("ws") is not None and s["limit"] else ""
    top = ", ".join(f"{t[1].replace('chromium-', '')} {t[3]:.0f}" for t in s["chrom_top"]) or "none"
    return (
        f"WS={f(s.get('ws'))}/{f(s['limit'])}{pct} used={f(s['used'])} anon={f(s.get('anon'))} "
        f"file={f(s.get('file'))} shmem={f(s.get('shmem'))} peak={f(s['peak'])} oom_kills={s['oom']} | "
        f"streamlit={f(s.get('py_rss'))} node={f(s.get('node_rss'))} "
        f"chromium procs={s['chrom_n']} rss={f(s['chrom_rss'])} [{top}] | "
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
    lim = cg["limit_mb"]
    if lim:
        job_log(job_dir, f"MEMORY POLICY: hard guard at {MEM_GUARD_FRACTION:.0%} of working set "
                         f"({lim * MEM_GUARD_FRACTION:.0f} MB, {MEM_GUARD_SAMPLES} samples in a row); "
                         f"soft recycle at {SOFT_RECYCLE_FRACTION:.0%} ({lim * SOFT_RECYCLE_FRACTION:.0f} MB) "
                         f"or Chromium +{CHROMIUM_GROWTH_MB} MB.")


class DiagnosticsMonitor(threading.Thread):
    """Samples memory / processes, logs them, and kills Chromium only under sustained real memory pressure."""

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
        self._over = 0
        self._last_logged = 0.0
        self._last_used = 0.0

    def reset(self):
        """Called at the start of every browser session."""
        self.tripped = None
        self._over = 0
        self._seen_chromium = False
        self._gone_logged = False

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
        ws = s["ws"] or 0.0
        self.peak_used = max(self.peak_used, ws)
        self.peak_chromium = max(self.peak_chromium, s["chrom_rss"])

        if s["chrom_n"]:
            self._seen_chromium = True
        elif self._seen_chromium and not self._gone_logged and self.phase not in QUIET_PHASES:
            self._gone_logged = True
            job_log(self.job_dir, "!!! ALL CHROMIUM PROCESSES ARE GONE (browser exited or was killed). "
                                  f"Last memory state: {format_snapshot(s)}", "ERROR")

        if elapsed < 60 or now - self._last_logged >= 15 or ws - self._last_used >= 100:
            self._last_logged, self._last_used = now, ws
            job_log(self.job_dir, f"MON +{int(elapsed)}s [{self.phase}] {format_snapshot(s)}")

        limit = s["limit"]
        if self.phase in QUIET_PHASES:
            self._over = 0
        elif limit and ws > limit * MEM_GUARD_FRACTION:
            self._over += 1
        else:
            self._over = 0

        if limit and self._over >= MEM_GUARD_SAMPLES and not self.tripped:
            self.tripped = (
                f"Memory guard stopped Chromium: the container working set stayed above "
                f"{MEM_GUARD_FRACTION:.0%} of the limit ({ws:.0f} of {limit:.0f} MB) for "
                f"{MEM_GUARD_SAMPLES} samples. Chromium used {s['chrom_rss']:.0f} MB."
            )
            job_log(self.job_dir, f"!!! {self.tripped}", "ERROR")
            for pid, role, status, rss in sorted(s["tree"], key=lambda t: -t[3])[:8]:
                job_log(self.job_dir, f"    pid={pid} {role} status={status} rss={rss:.0f}MB", "ERROR")
            kill_browser_processes(self.job_dir)


def attach_diagnostics(job_dir: Path, browser, page):
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
        if msg.type in ("error", "warning") and counters["console"] < 100:
            counters["console"] += 1
            job_log(job_dir, f"[console.{msg.type}] {msg.text[:300]}", "WARNING")

    @guarded
    def on_requestfailed(r):
        if counters["net"] < 30:
            counters["net"] += 1
            job_log(job_dir, f"[requestfailed] {r.method} {r.url[:150]} -> {r.failure}", "WARNING")

    browser.on("disconnected", guarded(lambda *_: job_log(
        job_dir, "BROWSER DISCONNECTED (closed or crashed)", "WARNING")))
    page.on("crash", guarded(lambda *_: job_log(
        job_dir, "!!! PAGE CRASHED: the renderer process died (usually out of memory)", "ERROR")))
    page.on("console", on_console)
    page.on("pageerror", guarded(lambda err: job_log(job_dir, f"[pageerror] {str(err)[:300]}", "WARNING")))
    page.on("requestfailed", on_requestfailed)


def take_shot(page, job_dir: Path, name: str):
    try:
        page.screenshot(path=str(job_dir / f"{name}.jpg"), type="jpeg", quality=50, timeout=15000)
        job_log(job_dir, f"Screenshot saved: {name}.jpg")
    except Exception as exc:
        job_log(job_dir, f"Screenshot {name} failed: {str(exc)[:150]}", "WARNING")


def _dispose(handle):
    try:
        handle.dispose()
    except Exception:
        pass


# =====================================================================
#  Playwright helpers
# =====================================================================
def launch_browser(pw, job_dir: Path, dl_dir: Path):
    args = [
        "--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu", "--mute-audio",
        "--renderer-process-limit=1", "--no-zygote",
        "--disable-features=IsolateOrigins,site-per-process", "--disable-site-isolation-trials",
        "--disable-extensions", "--disable-background-networking", "--disable-sync",
        "--disable-component-update", "--disable-breakpad", "--metrics-recording-only", "--no-first-run",
    ]
    limit = cgroup_memory()["limit_mb"]
    if limit:
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


def first_visible(page, selector, timeout=DEFAULT_TIMEOUT):
    """Return the first VISIBLE match of a selector (the hidden 'Active Alarm' tab may contain look-alikes)."""
    loc = page.locator(selector)
    end = time.time() + timeout
    while True:
        try:
            n = loc.count()
            for i in range(n):
                cand = loc.nth(i)
                if cand.is_visible():
                    return cand
        except PWError:
            pass
        if time.time() > end:
            raise RuntimeError(f"No visible element found for: {selector}")
        page.wait_for_timeout(300)


def safe_click(loc, timeout=10000):
    try:
        loc.scroll_into_view_if_needed(timeout=timeout)
    except PWError:
        pass
    try:
        loc.click(timeout=timeout)
    except PWError:
        loc.evaluate("el => el.click()")


def click_first(page, selectors, description, total_timeout=DEFAULT_TIMEOUT):
    per = max(5, int(total_timeout / max(1, len(selectors))))
    last_exc = None
    for sel in selectors:
        try:
            safe_click(first_visible(page, sel, per))
            return
        except (PWError, RuntimeError) as exc:
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


def pw_open_alarm_page(page, job_dir: Path, creds: dict):
    """Open the alarm dashboard (a fresh load = clean filters). Logs in again if the session was lost."""
    page.goto(ALARM_URL, wait_until="load", timeout=60000)
    if "/auth/login" in page.url:
        job_log(job_dir, "Session lost (redirected to login): logging in again.", "WARNING")
        pw_login(page, creds["username"], creds["password"])
        page.goto(ALARM_URL, wait_until="load", timeout=60000)
        if "/auth/login" in page.url:
            raise RuntimeError("Redirected back to login page; session not authenticated.")
    page.wait_for_selector("body", state="attached")
    page.wait_for_timeout(2000)


CLOSED_TAB_XPATHS = [
    "xpath=//*[(self::button or self::a or self::li or self::div or self::span or @role='tab')]"
    "[normalize-space()='CLOSED ALARM' or normalize-space()='Closed Alarm']",
    f"xpath=//*[contains({UPPER}, 'CLOSED ALARM') and (self::button or self::a or @role='tab' or self::li)]",
]


def _closed_tab_visible(page) -> bool:
    for sel in CLOSED_TAB_XPATHS:
        try:
            loc = page.locator(sel)
            for i in range(loc.count()):
                if loc.nth(i).is_visible():
                    return True
        except PWError:
            pass
    return False


def pw_scroll_to_alarm_section(page):
    """Scroll down the dashboard (past Daily Analytics) until the ACTIVE / CLOSED ALARM tabs are on screen."""
    section = page.locator(
        "xpath=//*[self::h1 or self::h2 or self::h3 or self::h4 or self::h5 or self::div or self::span]"
        "[normalize-space()='Alarm' or normalize-space()='ALARM']"
    ).first
    try:
        section.wait_for(state="attached", timeout=5000)
        section.evaluate("el => el.scrollIntoView({block: 'start'})")
    except PWError:
        pass
    for _ in range(25):  # keep scrolling until the tab appears
        if _closed_tab_visible(page):
            return
        page.mouse.wheel(0, 700)
        page.wait_for_timeout(500)
    page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
    page.wait_for_timeout(800)


def wait_loading_done(page, timeout=10000):
    try:
        page.wait_for_selector(
            ".spinner, .loading, .loader, .spinner-border, [class*='loading'], [class*='spinner']",
            state="hidden", timeout=timeout,
        )
    except PWError:
        pass


def pw_click_closed_alarm_tab(page, job_dir: Path, verbose=True):
    click_first(page, CLOSED_TAB_XPATHS, "CLOSED ALARM tab")
    # wait until the Closed Alarm table (header has Site ID + Severity) is really on screen
    end = time.time() + 60
    headers = []
    while time.time() < end:
        data = page.evaluate(READ_TABLE_JS)
        if data and data["heads"]:
            headers = data["heads"]
            if data["rows"]:
                break
        page.wait_for_timeout(500)
    if not headers:
        raise RuntimeError("The CLOSED ALARM table did not appear after clicking the tab.")
    if verbose:
        job_log(job_dir, f"Closed Alarm table detected, columns: {headers}")
    wait_loading_done(page, 20000)
    page.wait_for_timeout(3000)   # the pagination footer may render later than the table


# =====================================================================
#  Filter controls: Start Date / End Date / Apply / Items per page
# =====================================================================
JS_SET_VALUE = """(el, v) => {
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
  setter.call(el, v);
  el.dispatchEvent(new Event('input', {bubbles: true}));
  el.dispatchEvent(new Event('change', {bubbles: true}));
  el.dispatchEvent(new Event('blur', {bubbles: true}));
}"""

# The page also has a "Daily Analytics" section with its own Start/End Date boxes and tables.
# Everything below is anchored to the CLOSED ALARM table (header contains Site ID + Severity),
# never to "the first visible thing with that label".
_FIND_TABLE = r"""
const findTable = () => {
  const tables = Array.from(document.querySelectorAll('table')).filter(t => t.offsetParent !== null);
  const head = t => ((t.tHead ? t.tHead.innerText : '') || '').toLowerCase();
  return tables.find(t => /site\s*id/.test(head(t)) && /severity/.test(head(t)))
      || tables.find(t => /site\s*id/.test(head(t)))
      || null;
};
"""

MARK_CONTROL_JS = "(kind) => {" + _FIND_TABLE + r"""
  document.querySelectorAll('[data-pw-target]').forEach(e => e.removeAttribute('data-pw-target'));
  const table = findTable();
  if (!table) return 'no-closed-alarm-table';
  const vis = el => el.offsetParent !== null || el.getClientRects().length > 0;
  const own = el => Array.from(el.childNodes).filter(n => n.nodeType === 3)
                         .map(n => n.textContent).join(' ').replace(/[*:]/g, '').replace(/\s+/g, ' ').trim().toLowerCase();
  const before = el => !table.contains(el) && (el.compareDocumentPosition(table) & Node.DOCUMENT_POSITION_FOLLOWING);
  const after = el => !table.contains(el) && (table.compareDocumentPosition(el) & Node.DOCUMENT_POSITION_FOLLOWING);
  const all = Array.from(document.querySelectorAll('body *'));
  const tableTop = table.getBoundingClientRect().top;
  let target = null;
  if (kind === 'start' || kind === 'end') {
    const text = kind === 'start' ? 'start date' : 'end date';
    // the label that sits closest ABOVE the Closed Alarm table
    const labels = all.filter(el => vis(el) && before(el) && own(el) === text);
    const label = labels[labels.length - 1];
    if (!label) return 'no-label';
    const inputs = Array.from(document.querySelectorAll('input'))
      .filter(i => vis(i) && (label.contains(i) || (label.compareDocumentPosition(i) & Node.DOCUMENT_POSITION_FOLLOWING)));
    target = inputs[0];
  } else if (kind === 'apply') {
    const btns = Array.from(document.querySelectorAll('button, [role="button"]'))
      .filter(el => vis(el) && before(el) && (el.innerText || '').replace(/\s+/g, ' ').trim().toUpperCase() === 'APPLY');
    target = btns[btns.length - 1];
  } else if (kind === 'items') {
    const labels = all.filter(el => vis(el) && after(el) && own(el) === 'items');
    const label = labels[0];
    if (!label) return 'no-label';
    const cands = Array.from(document.querySelectorAll('select, [role="combobox"], mat-select, [class*="select"]'))
      .filter(el => vis(el) && (label.compareDocumentPosition(el) & Node.DOCUMENT_POSITION_FOLLOWING));
    target = cands[0];
  }
  if (!target) return 'no-target';
  target.setAttribute('data-pw-target', kind);
  const d = Math.round(Math.abs(target.getBoundingClientRect().top - tableTop));
  return 'ok:' + d;
}"""


def mark_control(page, kind: str, timeout=DEFAULT_TIMEOUT):
    """Locate a filter control that belongs to the CLOSED ALARM table (not Daily Analytics)."""
    end = time.time() + timeout
    res = ""
    while time.time() < end:
        res = page.evaluate(MARK_CONTROL_JS, kind)
        if str(res).startswith("ok"):
            return page.locator(f"[data-pw-target='{kind}']").first
        page.wait_for_timeout(400)
    raise RuntimeError(f"Could not find the '{kind}' control next to the Closed Alarm table (result: {res}).")


def pw_set_date(page, job_dir: Path, label: str, value: date, verbose=True):
    """Type a date into the Closed Alarm 'Start Date' / 'End Date' box and verify what the page holds."""
    text = value.strftime(DATE_INPUT_FORMAT)
    want = re.sub(r"\D", "", text)
    kind = "start" if label.lower().startswith("start") else "end"
    inp = mark_control(page, kind)
    inp.wait_for(state="visible", timeout=DEFAULT_TIMEOUT * 1000)
    inp.scroll_into_view_if_needed()
    last = ""
    for mode in ("typed", "digits", "js"):
        try:
            inp = mark_control(page, kind, 10)
            if mode == "js":
                inp.evaluate(JS_SET_VALUE, text)
            else:
                inp.click()
                inp.press("Control+A")
                inp.press("Backspace")
                inp.press_sequentially(text if mode == "typed" else want, delay=50)
            inp.press("Tab")
            page.wait_for_timeout(400)
            last = mark_control(page, kind, 10).input_value()
        except (PWError, RuntimeError) as exc:
            job_log(job_dir, f"[{label}] mode={mode} error: {str(exc)[:120]}", "WARNING")
            continue
        if re.sub(r"\D", "", last) == want:
            if verbose:
                job_log(job_dir, f"[{label}] (Closed Alarm filter) set to '{last}' (mode={mode})")
            return
        job_log(job_dir, f"[{label}] mode={mode} left '{last}' instead of '{text}'", "WARNING")
    raise RuntimeError(f"Could not set {label} to {text} (the box shows '{last}').")


def pw_click_apply(page):
    safe_click(mark_control(page, "apply"))


READ_TABLE_JS = "() => {" + _FIND_TABLE + r"""
  const clean = s => (s || '').replace(/\s+/g, ' ').trim();
  const table = findTable();
  if (!table) return null;
  const heads = Array.from(table.querySelectorAll('thead th')).map(th => clean(th.innerText));
  const rows = Array.from(table.querySelectorAll('tbody tr')).map(tr =>
      Array.from(tr.querySelectorAll('td')).map(td => clean(td.innerText)));
  return {heads, rows};
}"""

COUNT_ROWS_JS = "() => {" + _FIND_TABLE + r"""
  const table = findTable();
  if (!table) return -1;
  return Array.from(table.querySelectorAll('tbody tr'))
      .filter(tr => tr.querySelectorAll('td').length >= 3 && (tr.innerText || '').trim()).length;
}"""

# Tiny signature of the table (row count + first row + last row). Computed INSIDE the browser,
# so a 500-row table is never copied to Python while we wait for a page change.
SIG_JS = "() => {" + _FIND_TABLE + r"""
  const t = findTable();
  if (!t) return '';
  const tr = t.querySelectorAll('tbody tr');
  if (!tr.length) return '';
  const c = e => (e.innerText || '').replace(/\s+/g, ' ').trim();
  return tr.length + '|' + c(tr[0]) + '|' + c(tr[tr.length - 1]);
}"""

TOTAL_PAGES_JS = "() => {" + _FIND_TABLE + r"""
  const table = findTable();
  if (!table) return 0;
  let max = 0;
  document.querySelectorAll('button, a, li, [role="button"]').forEach(el => {
    if (table.contains(el) || !(table.compareDocumentPosition(el) & Node.DOCUMENT_POSITION_FOLLOWING)) return;
    const t = (el.innerText || '').trim();
    if (/^\d{1,5}$/.test(t)) max = Math.max(max, parseInt(t, 10));
  });
  return max;
}"""

FIND_NEXT_JS = "() => {" + _FIND_TABLE + r"""
  const table = findTable();
  if (!table) return null;
  const after = el => !table.contains(el) && (table.compareDocumentPosition(el) & Node.DOCUMENT_POSITION_FOLLOWING);
  const vis = el => el.offsetParent !== null || el.getClientRects().length > 0;
  const cands = Array.from(document.querySelectorAll('button, a, [role="button"]')).filter(el => after(el) && vis(el));
  if (!cands.length) return null;
  const txt = el => (el.innerText || '').trim();
  const cls = el => (el.className && el.className.toString) ? el.className.toString() : '';
  const meta = el => ((el.getAttribute('aria-label') || '') + ' ' + (el.getAttribute('title') || '') + ' ' + cls(el)).toLowerCase();
  const byMeta = cands.filter(el => /next|chevron_right|chevron-right|arrow_forward/.test(meta(el)) && !/last|prev|first/.test(meta(el)));
  if (byMeta.length) return byMeta[byMeta.length - 1];
  const byText = cands.filter(el => /^(chevron_right|navigate_next|keyboard_arrow_right|arrow_forward_ios|arrow_right|›|>|»)$/i.test(txt(el)));
  if (byText.length) return byText[byText.length - 1];
  const noDigit = cands.filter(el => !/\d/.test(txt(el)));
  return noDigit.length ? noDigit[noDigit.length - 1] : null;
}"""

# The page-number button that gets us closest to `target` without passing it:
# the highest visible number n with  current < n <= target.
FIND_PAGE_BTN_JS = "([target, cur]) => {" + _FIND_TABLE + r"""
  const table = findTable();
  if (!table) return null;
  const after = el => !table.contains(el) && (table.compareDocumentPosition(el) & Node.DOCUMENT_POSITION_FOLLOWING);
  const vis = el => el.offsetParent !== null || el.getClientRects().length > 0;
  let best = null, bestN = -1;
  document.querySelectorAll('button, a, li, [role="button"]').forEach(el => {
    if (!after(el) || !vis(el)) return;
    const t = (el.innerText || '').trim();
    if (!/^\d{1,5}$/.test(t)) return;
    const n = parseInt(t, 10);
    if (n > cur && n <= target && n >= bestN) { best = el; bestN = n; }   // >= : prefer the inner <button> over its <li>
  });
  return best;
}"""

IS_DISABLED_JS = r"""
el => {
  const cls = n => ((n && n.className && n.className.toString) ? n.className.toString() : '').toLowerCase();
  const p = el.parentElement;
  return !!(el.disabled || el.hasAttribute('disabled') || el.getAttribute('aria-disabled') === 'true'
            || /disabled/.test(cls(el)) || (p && /disabled/.test(cls(p))));
}
"""

_ICON_TAIL_RE = re.compile(r"\s+(info|info_outline|help|help_outline|arrow_upward|arrow_downward|unfold_more)$", re.I)


def read_table(page):
    """Return (headers, rows) of the visible table. Only used for the initial checks (it copies the whole table)."""
    data = page.evaluate(READ_TABLE_JS)
    if not data:
        return [], []
    rows = [r for r in data["rows"] if len(r) >= 3 and any(r)]   # skip 'no data' placeholder rows
    width = max([len(data["heads"])] + [len(r) for r in rows])
    seen, headers = {}, []
    for i in range(width):
        h = _ICON_TAIL_RE.sub("", data["heads"][i]).strip() if i < len(data["heads"]) else ""
        h = h or f"col_{i + 1}"
        n = seen.get(h, 0)
        seen[h] = n + 1
        headers.append(h if n == 0 else f"{h}_{n + 1}")
    return headers, rows


def count_rows(page) -> int:
    """Cheap row counter (does not copy every cell back to Python, so it is fine for 500-row tables)."""
    try:
        return max(0, int(page.evaluate(COUNT_ROWS_JS)))
    except (PWError, TypeError, ValueError):
        return 0


def get_sig(page) -> str:
    try:
        return page.evaluate(SIG_JS) or ""
    except PWError:
        return ""


def wait_sig_change(page, old, timeout=40, poll_ms=700):
    """Light-weight page-change detection: compares a tiny signature computed inside the browser."""
    end = time.time() + timeout
    while time.time() < end:
        s = get_sig(page)
        if s and s != old:
            page.wait_for_timeout(500)
            if get_sig(page) == s:
                return True
            continue
        page.wait_for_timeout(poll_ms)
    return False


# ---------------------------------------------------------------------
#  "Items" dropdown (10 / 50 / 100 / 250 / 500) -- FOOTER DISCOVERY EDITION
#  (unchanged: this part already works)
# ---------------------------------------------------------------------
SIZE_ORDER = ["10", "50", "100", "250", "500"]
SITE_INFO = {"framework": "generic", "virtualized": False}

SELECT_LIKE_CSS = (
    "select, [role='combobox'], [role='listbox'], [aria-haspopup='listbox'], mat-select, "
    ".v-select, .v-field, .MuiSelect-select, .MuiTablePagination-select, .ant-select, .q-select, "
    "[class*='select' i], [class*='dropdown' i], [class*='combo' i]"
)

FRAMEWORK_SELECTORS = {
    "vuetify_app": ".v-application, .v-app",
    "vuetify_table": ".v-data-table",
    "vuetify2_footer": ".v-data-footer",
    "vuetify3_footer": ".v-data-table-footer",
    "vuetify_select": ".v-select",
    "vuetify_virtual": ".v-data-table-virtual, .v-virtual-scroll",
    "mui_table": ".MuiTable-root",
    "mui_pagination": ".MuiTablePagination-root",
    "mui_datagrid": ".MuiDataGrid-root",
    "mui_virtual": ".MuiDataGrid-virtualScroller",
    "mat_paginator": "mat-paginator, .mat-mdc-paginator",
    "ant_table": ".ant-table",
    "ant_pagination": ".ant-pagination",
    "quasar": ".q-table",
    "ag_grid": ".ag-root",
    "virtual_strict": ".vue-recycle-scroller, .vue-virtual-scroller, [class*='virtual-scroll' i], "
                      "[class*='react-window' i], [class*='virtuoso' i]",
    "virtual_loose": "[class*='virtual' i]",
    "plain_table": "table",
    "tfoot": "tfoot",
}

# (name, css) -> tried in this order with Playwright locators. Framework-specific first, then generic.
ITEMS_SELECTORS = [
    ("vuetify3 items-per-page", ".v-data-table-footer__items-per-page .v-select, "
                                ".v-data-table-footer__items-per-page .v-field"),
    ("vuetify2 footer select", ".v-data-footer__select .v-select"),
    ("mui table pagination", ".MuiTablePagination-select, .MuiTablePagination-root [role='combobox']"),
    ("angular-material paginator", "mat-paginator mat-select, .mat-mdc-paginator-page-size-select mat-select, "
                                   ".mat-paginator-page-size-select mat-select"),
    ("antd pagination options", ".ant-pagination-options .ant-select"),
    ("quasar table bottom", ".q-table__bottom .q-select"),
    ("aria-label per page", "[aria-label*='per page' i], [aria-label*='items' i][role='combobox'], "
                            "[aria-label*='rows' i][role='combobox']"),
]

FRAMEWORK_JS = """(sels) => {
  const out = {};
  for (const [k, s] of Object.entries(sels)) {
    try { out[k] = document.querySelectorAll(s).length; } catch (e) { out[k] = 'err'; }
  }
  return out;
}"""

SCROLL_FOOTER_JS = "() => {" + _FIND_TABLE + r"""
  const out = {containers: [], win: null, noTable: false};
  const table = findTable();
  const scrollable = e => {
    const s = getComputedStyle(e);
    return /(auto|scroll|overlay)/.test(s.overflowY) && e.scrollHeight > e.clientHeight + 2;
  };
  const describe = e => e.tagName.toLowerCase() + (e.id ? '#' + e.id : '') +
      ((e.className && e.className.toString && e.className.toString().trim())
        ? '.' + e.className.toString().trim().split(/\s+/).slice(0, 3).join('.') : '');
  if (table) {
    const rows = table.querySelectorAll('tbody tr');
    if (rows.length) rows[rows.length - 1].scrollIntoView({block: 'end', behavior: 'instant'});
    let p = table.parentElement;
    while (p && p !== document.body && p !== document.documentElement) {
      if (scrollable(p)) {
        const before = Math.round(p.scrollTop);
        p.scrollTop = p.scrollHeight;
        out.containers.push({el: describe(p), before: before, after: Math.round(p.scrollTop),
                             scrollH: p.scrollHeight, clientH: p.clientHeight});
      }
      p = p.parentElement;
    }
  } else {
    out.noTable = true;
  }
  const se = document.scrollingElement || document.documentElement;
  const b = Math.round(se.scrollTop);
  se.scrollTop = se.scrollHeight;
  window.scrollTo(0, document.body.scrollHeight);
  out.win = {before: b, after: Math.round(se.scrollTop), scrollH: se.scrollHeight, innerH: window.innerHeight};
  return out;
}"""

FOOTER_DUMP_JS = "() => {" + _FIND_TABLE + r"""
  const table = findTable();
  if (!table) return {error: 'no Closed Alarm table found'};
  const cls = e => (e.className && e.className.toString) ? e.className.toString().trim() : '';
  const out = {chain: [], tfoot: table.tFoot ? table.tFoot.outerHTML.slice(0, 6000) : '', following: []};
  let child = table, parent = table.parentElement, level = 0;
  while (parent && level < 7) {
    out.chain.push(level + ': ' + parent.tagName.toLowerCase() + (parent.id ? '#' + parent.id : '') + ' .' + cls(parent).slice(0, 120));
    let sib = child.nextElementSibling;
    while (sib) {
      out.following.push({level: level, tag: sib.tagName.toLowerCase(), cls: cls(sib).slice(0, 120),
                          html: sib.outerHTML.slice(0, 4000)});
      sib = sib.nextElementSibling;
    }
    child = parent; parent = parent.parentElement; level++;
  }
  return out;
}"""

ENUM_JS = "([css]) => {" + _FIND_TABLE + r"""
  const table = findTable();
  const vis = el => (el.offsetParent !== null || el.getClientRects().length > 0) && el.getBoundingClientRect().width > 0;
  const tr = table ? table.getBoundingClientRect() : null;
  const tableBottom = tr ? tr.bottom + window.scrollY : 0;
  const base = css + ", button, [role='button'], a[role='button'], input, textarea";
  const out = [];
  document.querySelectorAll(base).forEach(el => {
    if (!vis(el)) return;
    const r = el.getBoundingClientRect();
    const absTop = r.top + window.scrollY;
    const inTfoot = !!el.closest('tfoot');
    const inBody = table && table.contains(el) && !inTfoot;
    const afterTable = table ? ((table.compareDocumentPosition(el) & Node.DOCUMENT_POSITION_FOLLOWING) && !table.contains(el)) : true;
    if (inBody) return;
    if (!(inTfoot || afterTable || absTop >= tableBottom - 150)) return;
    out.push({
      tag: el.tagName.toLowerCase(), role: el.getAttribute('role') || '', type: el.getAttribute('type') || '',
      cls: ((el.className && el.className.toString) ? el.className.toString() : '').slice(0, 70),
      text: (el.innerText || el.value || '').replace(/\s+/g, ' ').trim().slice(0, 40),
      aria: el.getAttribute('aria-label') || '', title: el.getAttribute('title') || '',
      selectLike: el.matches(css), x: Math.round(r.left), y: Math.round(r.top),
      w: Math.round(r.width), h: Math.round(r.height), absTop: Math.round(absTop),
      cx: r.left + r.width / 2, cy: r.top + r.height / 2, left: r.left, top: r.top
    });
  });
  return out.slice(0, 90);
}"""

PAGINATION_JS = "([css]) => {" + _FIND_TABLE + r"""
  const table = findTable();
  const vis = el => (el.offsetParent !== null || el.getClientRects().length > 0) && el.getBoundingClientRect().width > 0;
  const cls = e => (e.className && e.className.toString) ? e.className.toString().trim() : '';
  const tableTop = table ? table.getBoundingClientRect().top + window.scrollY : 0;
  const below = el => (el.getBoundingClientRect().top + window.scrollY) >= tableTop - 5 && !(table && table.contains(el) && !el.closest('tfoot'));
  const named = Array.from(document.querySelectorAll(
    "nav, [role='navigation'], [class*='pagination' i], [class*='paginator' i], .v-pagination, .MuiPagination-root, " +
    ".MuiTablePagination-root, mat-paginator, .mat-mdc-paginator, .ant-pagination, .v-data-footer, " +
    ".v-data-table-footer, .q-table__bottom, tfoot")).filter(el => vis(el) && below(el));
  const isDigit = el => /^\d{1,5}$/.test((el.innerText || '').trim());
  const digits = Array.from(document.querySelectorAll('button, a, li, [role="button"]'))
      .filter(el => vis(el) && below(el) && isDigit(el));
  const common = list => {
    let a = list[0].parentElement;
    while (a && !list.every(x => a.contains(x))) a = a.parentElement;
    return a;
  };
  let root = null, how = '';
  if (named.length) { root = named[named.length - 1]; how = 'named container'; }
  else if (digits.length >= 2) { root = common(digits); how = 'common parent of ' + digits.length + ' page-number buttons'; }
  else if (digits.length === 1) { root = digits[0].parentElement; how = 'parent of the only page-number button'; }
  if (!root) return {root: null, digits: digits.length, named: named.length};
  root.scrollIntoView({block: 'center', behavior: 'instant'});
  const chain = [];
  let n = root;
  for (let i = 0; i < 6 && n && n !== document.body; i++) {
    chain.push(i + ': ' + n.tagName.toLowerCase() + (n.id ? '#' + n.id : '') + ' .' + cls(n).slice(0, 100));
    n = n.parentElement;
  }
  // look for a select-like control in the pagination root and then in its ancestors
  let control = null, level = -1;
  n = root;
  for (let i = 0; i < 6 && n && !control; i++) {
    const cands = Array.from(n.querySelectorAll(css)).filter(c => {
      if (!vis(c)) return false;
      const r = c.getBoundingClientRect();
      if (r.width < 30 || r.height < 16) return false;
      if (isDigit(c) && c.tagName !== 'SELECT') return false;
      return true;
    });
    cands.sort((a, b) => {
      const ra = a.getBoundingClientRect(), rb = b.getBoundingClientRect();
      return ra.width * ra.height - rb.width * rb.height;
    });
    if (cands.length) { control = cands[0]; level = i; }
    n = n.parentElement;
  }
  const rr = root.getBoundingClientRect();
  const res = {root: {tag: root.tagName.toLowerCase(), cls: cls(root).slice(0, 150), how: how,
                      html: root.outerHTML.slice(0, 2500), left: rr.left, top: rr.top, w: rr.width, h: rr.height},
               chain: chain, digits: digits.length, named: named.length, control: null, level: level};
  if (control) {
    const r = control.getBoundingClientRect();
    res.control = {tag: control.tagName.toLowerCase(), cls: cls(control).slice(0, 120),
                   left: r.left, top: r.top, w: r.width, h: r.height, text: (control.innerText || '').slice(0, 30)};
  }
  return res;
}"""

LABEL_JS = "([css]) => {" + _FIND_TABLE + r"""
  const table = findTable();
  const tableTop = table ? table.getBoundingClientRect().top + window.scrollY : 0;
  const vis = el => (el.offsetParent !== null || el.getClientRects().length > 0) && el.getBoundingClientRect().width > 0;
  const own = el => Array.from(el.childNodes).filter(n => n.nodeType === 3)
                         .map(n => n.textContent).join(' ').replace(/[*:]/g, '').replace(/\s+/g, ' ').trim().toLowerCase();
  const re = /^(items|items per page|item per page|rows per page|per page|show items|show rows|tampilkan|baris per halaman|item per halaman|jumlah baris|jumlah data)$/;
  let labels = Array.from(document.querySelectorAll('body *')).filter(el => vis(el) && re.test(own(el)));
  labels = labels.filter(el => (el.getBoundingClientRect().top + window.scrollY) >= tableTop - 5);
  if (!labels.length) return null;
  const label = labels[labels.length - 1];
  label.scrollIntoView({block: 'center', behavior: 'instant'});
  const lr = label.getBoundingClientRect();
  const lcy = (lr.top + lr.bottom) / 2;
  const cls = e => (e.className && e.className.toString) ? e.className.toString().trim() : '';
  let control = null, node = label, level = -1;
  for (let i = 0; i < 5 && !control; i++) {
    node = node.parentElement;
    if (!node) break;
    const cands = Array.from(node.querySelectorAll(css)).filter(c => {
      if (!vis(c) || c.contains(label) || label.contains(c)) return false;
      const r = c.getBoundingClientRect();
      const cy = (r.top + r.bottom) / 2;
      return r.width >= 30 && r.height >= 16 && r.right > lr.left && Math.abs(cy - lcy) < 45;
    });
    cands.sort((a, b) => {
      const ra = a.getBoundingClientRect(), rb = b.getBoundingClientRect();
      return ra.width * ra.height - rb.width * rb.height;
    });
    if (cands.length) { control = cands[0]; level = i; }
  }
  const res = {label: {text: own(label), tag: label.tagName.toLowerCase(), cls: cls(label).slice(0, 100),
                       left: lr.left, right: lr.right, top: lr.top, bottom: lr.bottom},
               control: null, level: level,
               html: (label.parentElement && label.parentElement.parentElement
                      ? label.parentElement.parentElement : label.parentElement || label).outerHTML.slice(0, 1500)};
  if (control) {
    const r = control.getBoundingClientRect();
    res.control = {tag: control.tagName.toLowerCase(), cls: cls(control).slice(0, 120),
                   left: r.left, top: r.top, w: r.width, h: r.height};
  }
  return res;
}"""

NATIVE_SELECT_JS = r"""() => {
  document.querySelectorAll('[data-pw-items]').forEach(e => e.removeAttribute('data-pw-items'));
  const sizes = ['10', '50', '100', '250', '500'];
  const all = Array.from(document.querySelectorAll('select'));
  const info = {total: all.length, matches: []};
  all.forEach((s, i) => {
    const vals = Array.from(s.options).map(o => (o.value || o.text || '').trim());
    const hit = sizes.filter(z => vals.includes(z) || Array.from(s.options).some(o => (o.text || '').trim() === z));
    if (hit.length >= 2) info.matches.push({i: i, vals: vals.slice(0, 12), value: s.value});
  });
  if (info.matches.length) {
    all[info.matches[info.matches.length - 1].i].setAttribute('data-pw-items', '1');
    all[info.matches[info.matches.length - 1].i].scrollIntoView({block: 'center', behavior: 'instant'});
  }
  return info;
}"""

TEXT_HITS_JS = "() => {" + _FIND_TABLE + r"""
  const table = findTable();
  const re = /items|per page|rows|tampilkan|halaman|show/i;
  const own = el => Array.from(el.childNodes).filter(n => n.nodeType === 3).map(n => n.textContent).join(' ').replace(/\s+/g, ' ').trim();
  const out = [];
  document.querySelectorAll('body *').forEach(el => {
    const t = own(el);
    if (!t || t.length > 40 || !re.test(t)) return;
    const r = el.getBoundingClientRect();
    out.push({tag: el.tagName.toLowerCase(), text: t, inTable: !!(table && table.contains(el)),
              inTfoot: !!el.closest('tfoot'),
              visible: (el.offsetParent !== null || el.getClientRects().length > 0) && r.width > 0,
              absTop: Math.round(r.top + window.scrollY), absLeft: Math.round(r.left)});
  });
  return {hits: out.slice(0, 30), docH: document.documentElement.scrollHeight, bodyTail: (document.body.innerText || '').slice(-500)};
}"""

PAGE_RANGE_JS = r"""() => {
  const els = document.querySelectorAll("tfoot, [class*='pagination' i], [class*='footer' i], mat-paginator, " +
                                        ".MuiTablePagination-root, .q-table__bottom, .ant-pagination");
  for (const el of els) {
    const t = (el.innerText || '');
    const m = t.match(/(\d[\d.,]*)\s*[-\u2013\u2014]\s*(\d[\d.,]*)\s*(?:of|dari|\/)\s*(\d[\d.,]*)/i);
    if (m) return [m[1], m[2], m[3]];
  }
  return null;
}"""

FIND_OPTIONS_JS = r"""([ax, ay]) => {
  const sizes = ['10', '50', '100', '250', '500'];
  const found = [];
  document.querySelectorAll('body *').forEach(el => {
    if (el.closest('table')) return;
    const r = el.getBoundingClientRect();
    if (r.width <= 0 || r.height <= 0) return;
    const s = getComputedStyle(el);
    if (s.visibility === 'hidden' || s.display === 'none') return;
    const t = (el.innerText || '').replace(/\s+/g, ' ').trim();
    if (!sizes.includes(t)) return;
    if (Math.abs(r.left - ax) > 500 || Math.abs((r.top + r.bottom) / 2 - ay) > 800) return;
    found.push({el, t, x: r.left + r.width / 2, y: r.top + r.height / 2, w: r.width, h: r.height});
  });
  const keep = found.filter(o => !found.some(p => p !== o && o.el.contains(p.el)));   // deepest only
  return keep.map(o => ({t: o.t, x: o.x, y: o.y, w: o.w, h: o.h}));
}"""


def _num(x) -> int:
    return int(re.sub(r"\D", "", str(x)) or 0)


def detect_site(page,
