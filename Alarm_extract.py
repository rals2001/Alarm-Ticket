"""
PMT Alarm – Closed Alarm downloader (3-step workflow, Playwright edition)

STEP 1  Retrieve Excel  -> login in Chromium, set Start/End Date, set 100 items per page,
                           then walk through every page of the CLOSED ALARM table and collect the rows
                           (no Download button, no export API call)
STEP 2  Filter Excel    -> stream the collected rows, keep the chosen Year-Month, write XLSX
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
from datetime import date, datetime, timedelta, timezone
from itertools import chain, islice
from pathlib import Path

import streamlit as st
import ijson
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
DEFAULT_TIMEOUT = 40          # seconds for normal waits
JOBS_ROOT = Path(tempfile.gettempdir()) / "pmt_alarm_jobs"
STALE_AFTER = 300             # a "running" step with no heartbeat for 5 min = interrupted
JOB_MAX_AGE_HOURS = 12
POLL_SECONDS = 3
MEM_GUARD_FRACTION = 0.85     # kill Chromium when the container reaches 85 % of its memory limit
MEM_LIMIT_MB_OVERRIDE = float(os.environ.get("MEM_LIMIT_MB", 0)) or None
MONITOR_INTERVAL = 2          # seconds between diagnostic samples
BLOCK_IMAGES = True

SITE_UTC_OFFSET_HOURS = 7     # the site works in WIB (UTC+7)
DATE_INPUT_FORMAT = "%d-%m-%Y"        # how the Start/End Date boxes display dates (02-10-2026)
PAGE_SIZE_CHOICES = ("100", "50", "20")  # tried in this order in the "Items" dropdown
MAX_PAGES = 3000              # safety cap for the pagination loop

# ---- optional: the site's own listing API (only READ from the browser traffic, never replayed) ----
API_HOST = "pmt-api.komdigi.go.id"
API_DATE_FIELD = None         # force the date column used by Step 2, e.g. "Alarm Start Time" or "api.created_at"
COLUMN_LABELS = {}            # e.g. {"api.created_at": "Created At"} to rename Excel headers
MONTHS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
# =========================================================

WIB = timezone(timedelta(hours=SITE_UTC_OFFSET_HOURS))

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


def source_path(job_dir: Path) -> Path:
    return job_dir / "source.json"


def source_ready(job_dir: Path) -> bool:
    p = source_path(job_dir)
    return p.exists() and p.stat().st_size > 0


def read_source_meta(job_dir: Path) -> dict:
    try:
        return json.loads((job_dir / "source_meta.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


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
    return {"used_mb": _mb(used), "limit_mb": MEM_LIMIT_MB_OVERRIDE or _mb(limit),
            "peak_mb": _mb(peak), "oom_kills": oom}


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
    if used is None and tree:
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
    """Samples memory / processes, logs them, and kills Chromium before the container hits its memory limit."""

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
        elif self._seen_chromium and not self._gone_logged and self.phase not in ("finished", "closing"):
            self._gone_logged = True
            job_log(self.job_dir, "!!! ALL CHROMIUM PROCESSES ARE GONE (browser exited or was killed). "
                                  f"Last memory state: {format_snapshot(s)}", "ERROR")

        if elapsed < 60 or now - self._last_logged >= 15 or used - self._last_used >= 100:
            self._last_logged, self._last_used = now, used
            job_log(self.job_dir, f"MON +{int(elapsed)}s [{self.phase}] {format_snapshot(s)}")

        limit = s["limit"]
        if limit and used > limit * MEM_GUARD_FRACTION and not self.tripped:
            self.tripped = (
                f"Memory guard stopped Chromium: the container reached {used:.0f} MB of its {limit:.0f} MB limit "
                f"({used / limit:.0%}). Chromium used {s['chrom_rss']:.0f} MB."
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

    browser.on("disconnected", guarded(lambda *_: job_log(
        job_dir, "BROWSER DISCONNECTED (closed or crashed)", "WARNING")))
    page.on("crash", guarded(lambda *_: job_log(
        job_dir, "!!! PAGE CRASHED: the renderer process died (usually out of memory)", "ERROR")))
    page.on("console", on_console)
    page.on("pageerror", guarded(lambda err: job_log(job_dir, f"[pageerror] {str(err)[:300]}", "WARNING")))
    page.on("requestfailed", guarded(lambda r: job_log(
        job_dir, f"[requestfailed] {r.method} {r.url[:150]} -> {r.failure}", "WARNING")))


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


def wait_loading_done(page, timeout=10000):
    try:
        page.wait_for_selector(
            ".spinner, .loading, .loader, .spinner-border, [class*='loading'], [class*='spinner']",
            state="hidden", timeout=timeout,
        )
    except PWError:
        pass


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
    wait_loading_done(page)
    page.wait_for_timeout(1000)


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


def pw_set_date(page, job_dir: Path, label: str, value: date):
    """Type a date into the 'Start Date' / 'End Date' box and verify what the page actually holds."""
    text = value.strftime(DATE_INPUT_FORMAT)
    want = re.sub(r"\D", "", text)
    lab = first_visible(page, f"xpath=//*[normalize-space(text())='{label}']")
    inp = lab.locator("xpath=(.//input | following::input)[1]")
    inp.wait_for(state="visible", timeout=DEFAULT_TIMEOUT * 1000)
    inp.scroll_into_view_if_needed()
    last = ""
    for mode in ("typed", "digits", "js"):
        try:
            if mode == "js":
                inp.evaluate(JS_SET_VALUE, text)
            else:
                inp.click()
                inp.press("Control+A")
                inp.press("Backspace")
                inp.press_sequentially(text if mode == "typed" else want, delay=50)
            inp.press("Tab")
            page.wait_for_timeout(400)
            last = inp.input_value()
        except PWError as exc:
            job_log(job_dir, f"[{label}] mode={mode} error: {str(exc)[:120]}", "WARNING")
            continue
        if re.sub(r"\D", "", last) == want:
            job_log(job_dir, f"[{label}] set to '{last}' (mode={mode})")
            return
        job_log(job_dir, f"[{label}] mode={mode} left '{last}' instead of '{text}'", "WARNING")
    raise RuntimeError(f"Could not set {label} to {text} (the box shows '{last}').")


def pw_click_apply(page):
    click_first(
        page,
        [f"xpath=//button[{UPPER}='APPLY']", f"xpath=//*[@role='button' and {UPPER}='APPLY']"],
        "Apply button",
    )


READ_TABLE_JS = r"""
() => {
  const clean = s => (s || '').replace(/\s+/g, ' ').trim();
  const table = Array.from(document.querySelectorAll('table')).find(t => t.offsetParent !== null);
  if (!table) return null;
  const heads = Array.from(table.querySelectorAll('thead th')).map(th => clean(th.innerText));
  const rows = Array.from(table.querySelectorAll('tbody tr')).map(tr =>
      Array.from(tr.querySelectorAll('td')).map(td => clean(td.innerText)));
  return {heads, rows};
}
"""

TOTAL_PAGES_JS = r"""
() => {
  const table = Array.from(document.querySelectorAll('table')).find(t => t.offsetParent !== null);
  if (!table) return 0;
  let max = 0;
  document.querySelectorAll('button, a, li, [role="button"]').forEach(el => {
    if (table.contains(el) || !(table.compareDocumentPosition(el) & Node.DOCUMENT_POSITION_FOLLOWING)) return;
    const t = (el.innerText || '').trim();
    if (/^\d{1,5}$/.test(t)) max = Math.max(max, parseInt(t, 10));
  });
  return max;
}
"""

FIND_NEXT_JS = r"""
() => {
  const table = Array.from(document.querySelectorAll('table')).find(t => t.offsetParent !== null);
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
}
"""

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
    """Return (headers, rows) of the visible table. rows = list of list[str]."""
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


def table_sig(rows):
    if not rows:
        return ""
    return f"{len(rows)}|{'¦'.join(rows[0])}|{'¦'.join(rows[-1])}"


def wait_table_change(page, old_sig, timeout=30):
    """Wait until the table content differs from old_sig and has stopped changing."""
    end = time.time() + timeout
    while time.time() < end:
        _, rows = read_table(page)
        if rows:
            s = table_sig(rows)
            if s != old_sig:
                page.wait_for_timeout(500)
                _, rows2 = read_table(page)
                if rows2 and table_sig(rows2) == s:
                    return True
                continue
        page.wait_for_timeout(400)
    return False


def pw_set_page_size(page, job_dir: Path):
    """Pick 100 (or 50 / 20) in the 'Items' dropdown. Returns the number of rows now shown."""
    _, rows0 = read_table(page)
    n0 = len(rows0)
    if n0 and n0 < 10:
        return n0  # everything already fits on one page
    old = table_sig(rows0)
    label = first_visible(page, "xpath=//*[normalize-space(text())='Items']", 20)
    for size in PAGE_SIZE_CHOICES:
        try:
            trigger = label.locator(
                "xpath=(following::*[self::select or @role='combobox' or self::mat-select "
                "or contains(@class,'select')])[1]")
            trigger.wait_for(state="visible", timeout=8000)
            trigger.scroll_into_view_if_needed()
            if trigger.evaluate("el => el.tagName") == "SELECT":
                try:
                    trigger.select_option(label=size)
                except PWError:
                    trigger.select_option(value=size)
            else:
                safe_click(trigger)
                opt = first_visible(
                    page,
                    f"xpath=//*[(@role='option' or self::mat-option or self::li or self::option) "
                    f"and normalize-space()='{size}']", 8)
                safe_click(opt)
            wait_table_change(page, old, 20)
            _, rows = read_table(page)
            if len(rows) > n0:
                job_log(job_dir, f"Page size set to {size}: {len(rows)} rows on the page now.")
                return len(rows)
            job_log(job_dir, f"Page size {size} did not change the table ({len(rows)} rows).", "WARNING")
        except (PWError, RuntimeError) as exc:
            job_log(job_dir, f"Page size {size} failed: {str(exc)[:150]}", "WARNING")
        try:
            page.keyboard.press("Escape")
        except PWError:
            pass
    job_log(job_dir, "Could not change the page size; continuing with the default (slower).", "WARNING")
    return n0


# =====================================================================
#  Optional: read the site's own listing API traffic (for real dates)
# =====================================================================
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


class ListingCapture:
    """
    Passively remembers the JSON the site's own table requests receive (small: one page each).
    Nothing is replayed and nothing is stored on disk. It is used to add real date fields to the
    scraped rows, because the table shows 'Invalid date' in the Alarm Start Time column.
    """

    def __init__(self, job_dir: Path):
        self.job_dir = job_dir
        self.items = []
        self.n = 0

    def attach(self, page):
        page.on("response", self._guard(self._on_response))

    @staticmethod
    def _guard(fn):
        def wrapper(*args):
            try:
                fn(*args)
            except Exception:
                pass
        return wrapper

    def _on_response(self, resp):
        req = resp.request
        if API_HOST not in resp.url or req.resource_type not in ("xhr", "fetch") or resp.status != 200:
            return
        if "json" not in resp.headers.get("content-type", ""):
            return
        if int(resp.headers.get("content-length", "0") or 0) > 8_000_000:
            return
        data = json.loads(resp.body())
        path, recs = find_records_path(data)
        if not recs:
            return
        self.n += 1
        self.items.append({"t": time.time(), "url": resp.url.split("?")[0], "records": recs})
        del self.items[:-8]
        if self.n <= 12:
            flat_first = flatten(recs[0])
            job_log(self.job_dir, f"[listing #{self.n}] {resp.url.split('?')[0][-70:]} path={'.'.join(path) or '(root)'} "
                                  f"records={len(recs)} fields={list(flat_first.keys())[:40]}")


API_EXTRA_KEY_RE = re.compile(r"date|time|_at$|^at$|creat|updat|start|clos|end$|resolv", re.I)


def match_api(listing: ListingCapture, since: float, rows):
    """Find the captured API page that corresponds to the rows currently shown. Returns list of flat dicts or None."""
    if not rows:
        return None

    def overlap(cells, flat):
        vals = {str(v).strip() for v in flat.values() if v is not None}
        return sum(1 for c in cells if len(c) > 2 and c in vals)

    for item in reversed(listing.items):
        if item["t"] < since - 0.5 or len(item["records"]) != len(rows):
            continue
        flats = [flatten(r) for r in item["records"]]
        if overlap(rows[0], flats[0]) >= 2 and overlap(rows[-1], flats[-1]) >= 2:
            return flats
    return None


# =====================================================================
#  Page-by-page scraping
# =====================================================================
def find_next(page):
    handle = page.evaluate_handle(FIND_NEXT_JS)
    el = handle.as_element()
    if el is None:
        return None, True
    return el, bool(el.evaluate(IS_DISABLED_JS))


def scrape_all_pages(page, job_dir: Path, rep: StepReporter, listing: ListingCapture, monitor, browser, since: float):
    all_rows, page_no, api_pages = [], 0, 0
    headers, rows = read_table(page)
    if not rows:
        raise RuntimeError("The table is empty for this date range (no closed alarms found).")
    total_pages = max(1, page.evaluate(TOTAL_PAGES_JS))

    while True:
        page_no += 1
        flats = match_api(listing, since, rows)
        if flats:
            api_pages += 1
        for i, cells in enumerate(rows):
            rec = dict(zip(headers, cells))
            if flats:
                for k, v in flats[i].items():
                    if API_EXTRA_KEY_RE.search(k) and not isinstance(v, (dict, list)):
                        rec[f"api.{k}"] = v
            all_rows.append(rec)

        total_pages = max(total_pages, page_no, page.evaluate(TOTAL_PAGES_JS))
        rep.running(f"Scraping page {page_no} of ~{total_pages}: {len(all_rows):,} rows collected...",
                    0.40 + 0.55 * min(page_no / total_pages, 1.0))

        if monitor is not None and monitor.tripped:
            raise RuntimeError(monitor.tripped)
        if not browser.is_connected() or page.is_closed():
            raise RuntimeError("Chromium exited while scraping the table.")
        if page_no >= MAX_PAGES:
            job_log(job_dir, f"Reached the safety cap of {MAX_PAGES} pages.", "WARNING")
            break

        nxt, disabled = find_next(page)
        if nxt is None:
            job_log(job_dir, "No 'next page' button found: assuming this is the only/last page.", "WARNING")
            break
        if disabled:
            job_log(job_dir, f"'Next' is disabled on page {page_no}: that was the last page.")
            break

        old = table_sig(rows)
        since = time.time()
        changed = False
        for attempt in (1, 2):
            safe_click(nxt)
            if wait_table_change(page, old, 25):
                changed = True
                break
            job_log(job_dir, f"Page {page_no}: table did not change after clicking next (attempt {attempt}).", "WARNING")
            nxt, disabled = find_next(page)
            if nxt is None or disabled:
                break
        if not changed:
            job_log(job_dir, "Table stopped changing: assuming the last page was reached.")
            break
        headers, rows = read_table(page)
        if not rows:
            break

    return all_rows, page_no, api_pages


# =====================================================================
#  STEP 1 worker – set filters, then walk through every page
# =====================================================================
def retrieve_worker(job_dir: Path, username: str, password: str, start_d: date, end_d: date):
    rep = StepReporter(job_dir, 1)
    dl_dir = job_dir / "dl"
    shutil.rmtree(dl_dir, ignore_errors=True)
    dl_dir.mkdir(parents=True, exist_ok=True)
    for pattern in ("shot_*.png", "source_meta.json", "error.png", "source.part"):
        for old in job_dir.glob(pattern):
            old.unlink(missing_ok=True)

    monitor = None
    try:
        log_system_info(job_dir)
        job_log(job_dir, f"Requested range: {start_d:%d-%m-%Y} -> {end_d:%d-%m-%Y}")
        rep.running("Starting browser...", 0.05)
        monitor = DiagnosticsMonitor(job_dir, dl_dir, rep)
        monitor.start()

        with sync_playwright() as pw:
            browser = launch_browser(pw, job_dir, dl_dir)
            page = None
            try:
                context = browser.new_context(accept_downloads=False, viewport={"width": 1600, "height": 1000})
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
                listing = ListingCapture(job_dir)
                listing.attach(page)

                monitor.phase = "login"
                rep.running("Logging in...", 0.10)
                pw_login(page, username, password)

                monitor.phase = "alarm-page"
                rep.running("Opening alarm page...", 0.18)
                pw_open_alarm_page(page)
                pw_scroll_to_alarm_section(page)

                rep.running("Opening CLOSED ALARM tab...", 0.22)
                pw_click_closed_alarm_tab(page)

                monitor.phase = "filters"
                rep.running(f"Setting Start Date {start_d:%d-%m-%Y} and End Date {end_d:%d-%m-%Y}...", 0.26)
                pw_set_date(page, job_dir, "Start Date", start_d)
                pw_set_date(page, job_dir, "End Date", end_d)
                take_shot(page, job_dir, "shot_1_filters")

                _, rows_before = read_table(page)
                since = time.time()
                rep.running("Applying the filter...", 0.30)
                pw_click_apply(page)
                wait_table_change(page, table_sig(rows_before), 25)
                wait_loading_done(page)
                page.wait_for_timeout(800)

                rep.running("Setting items per page...", 0.34)
                pw_set_page_size(page, job_dir)
                take_shot(page, job_dir, "shot_2_after_page_size")

                monitor.phase = "scrape"
                rep.running("Reading the table page by page...", 0.40)
                all_rows, pages, api_pages = scrape_all_pages(page, job_dir, rep, listing, monitor, browser, since)
                job_log(job_dir, f"Scraped {len(all_rows):,} rows from {pages} pages "
                                 f"(real API dates attached on {api_pages} of {pages} pages).")
                if api_pages == 0:
                    job_log(job_dir, "No API date fields could be attached. If the Alarm Start Time column says "
                                     "'Invalid date', Step 2 will not be able to filter by month. "
                                     "See the [listing #n] lines above.", "WARNING")
                take_shot(page, job_dir, "shot_3_last_page")
            except Exception:
                job_log(job_dir, "State at failure: " + format_snapshot(collect_snapshot(dl_dir)), "ERROR")
                if page is not None:
                    try:
                        page.screenshot(path=str(job_dir / "error.png"), timeout=10000)
                    except Exception:
                        pass
                raise
            finally:
                monitor.phase = "closing"
                try:
                    browser.close()
                except Exception:
                    pass

        rep.running(f"Saving {len(all_rows):,} rows...", 0.97)
        part = job_dir / "source.part"
        with open(part, "w", encoding="utf-8") as fh:
            json.dump({"records": all_rows}, fh, ensure_ascii=False)
        write_json_atomic(job_dir / "source_meta.json", {
            "records_prefix": "records.item", "rows": len(all_rows), "pages": pages, "api_pages": api_pages,
            "range_start": start_d.isoformat(), "range_end": end_d.isoformat(),
        })
        os.replace(part, source_path(job_dir))
        size_mb = source_path(job_dir).stat().st_size / 1024 / 1024
        rep.done(f"Retrieved {len(all_rows):,} rows from {pages} pages "
                 f"({start_d:%d-%m-%Y} to {end_d:%d-%m-%Y}).",
                 size_mb=round(size_mb, 2), rows=len(all_rows), pages=pages)
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
        (job_dir / "source.part").unlink(missing_ok=True)


# =====================================================================
#  STEP 2 worker – streaming filter of the collected rows (low memory)
# =====================================================================
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
        if not s or s.lower().startswith("invalid"):
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


_PREFERRED_DATE_FIELDS = ["alarm_start_time", "start_time", "start_date", "started_at", "start_at", "start",
                          "created_at", "createdat", "date_created", "data_created", "created", "created_date",
                          "createddate", "datetime_created", "create_at", "created_time"]


def pick_date_field(columns, sample_rows, parser):
    if API_DATE_FIELD:
        if API_DATE_FIELD not in columns:
            raise RuntimeError(f"API_DATE_FIELD '{API_DATE_FIELD}' not found. Available columns: {columns[:60]}")
        return API_DATE_FIELD
    lower = {c: c.lower().replace(" ", "_") for c in columns}
    ordered = [c for pref in _PREFERRED_DATE_FIELDS for c in columns if lower[c].split(".")[-1] == pref]
    ordered += [c for c in columns if ("start" in lower[c] or "creat" in lower[c]) and c not in ordered]
    ordered += [c for c in columns if any(t in lower[c] for t in ("date", "time", "_at")) and c not in ordered]
    for col in ordered:
        vals = [r.get(col) for r in sample_rows if r.get(col) not in (None, "")][:200]
        if vals and sum(parser(v) is not None for v in vals) / len(vals) >= 0.8:
            return col
    raise RuntimeError(
        "No usable date column found. The site's table shows 'Invalid date' in Alarm Start Time and no real date "
        "field could be read from the site's own data (see the [listing #n] lines in Diagnostics). "
        f"Columns available: {columns[:40]}")


def stream_filter_json(src: Path, dst: Path, year: int, month: int, prefix: str, on_progress):
    parser = DateParser()
    with open(src, "rb") as fh:
        records = ijson.items(fh, prefix, use_float=True)
        head = [flatten(r) for r in islice(records, 500)]
        if not head:
            raise RuntimeError(f"No records found at path '{prefix}' in the collected data.")

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
        meta = read_source_meta(job_dir)

        for old in job_dir.glob("result_*.xlsx"):
            old.unlink(missing_ok=True)

        rep.running(f"Reading source data and filtering for {label}...", None)

        def progress(total, kept):
            rep.running(f"Filtering {label}: {total:,} records scanned, {kept:,} kept so far...", None, quiet=True)

        stats = stream_filter_json(source_path(job_dir), part, year, month,
                                   meta.get("records_prefix", "records.item"), progress)
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


def default_range():
    """Last two months: the 1st of the previous month up to today (site time)."""
    today = datetime.now(WIB).date()
    prev_month_last_day = today.replace(day=1) - timedelta(days=1)
    return prev_month_last_day.replace(day=1), today


def render_workflow(job_dir: Path, job_id: str):
    s1, s2 = get_status(job_dir, 1), get_status(job_dir, 2)
    busy = s1["state"] == "running" or s2["state"] == "running"

    was_busy = st.session_state.get("_was_busy", False)
    st.session_state["_was_busy"] = busy
    if was_busy and not busy:
        st.rerun()

    src_ok = source_ready(job_dir)

    # ------------------------------------------------------------ STEP 1
    st.subheader("Step 1 · Retrieve Excel")
    st.caption("Logs in, fills Start/End Date, shows 100 items per page and reads every page of the Closed Alarm "
               "table one by one. No Download button is used. No filtering happens here.")
    if src_ok and s1["state"] != "running":
        meta = read_source_meta(job_dir)
        size_mb = source_path(job_dir).stat().st_size / 1024 / 1024
        rng = ""
        if meta.get("range_start"):
            rs, re_ = date.fromisoformat(meta["range_start"]), date.fromisoformat(meta["range_end"])
            rng = f" · {rs:%d-%m-%Y} to {re_:%d-%m-%Y}"
        st.success(f"✅ Source data is ready on the server: {meta.get('rows', '?'):,} rows from "
                   f"{meta.get('pages', '?')} pages{rng} ({size_mb:.1f} MB). No need to retrieve again."
                   if isinstance(meta.get("rows"), int) else
                   f"✅ Source data is ready on the server ({size_mb:.1f} MB).")
    show_status(s1 if s1["state"] in ("running", "error") else {"state": "idle"})
    if s1["state"] == "error" and (job_dir / "error.png").exists():
        st.image(str(job_dir / "error.png"), caption="Browser screenshot when the error happened")

    d0, d1 = default_range()
    with st.form("step1_form"):
        username = st.text_input("Username", disabled=busy)
        password = st.text_input("Password", type="password", disabled=busy)
        cs, ce = st.columns(2)
        start_d = cs.date_input("Start date", value=d0, format="DD/MM/YYYY", disabled=busy, key="rng_start")
        end_d = ce.date_input("End date", value=d1, format="DD/MM/YYYY", disabled=busy, key="rng_end")
        force = st.checkbox("Replace the existing data (retrieve again)", disabled=busy or not src_ok)
        go1 = st.form_submit_button("Retrieve Excel", type="primary", disabled=busy, use_container_width=True)

    if go1:
        if src_ok and not force:
            st.info("The source data already exists, so retrieving was skipped. "
                    "Tick “Replace the existing data” to retrieve a new copy.")
        elif not username or not password:
            st.error("Please enter both username and password.")
        elif start_d > end_d:
            st.error("Start date must not be after end date.")
        else:
            source_path(job_dir).unlink(missing_ok=True)
            (job_dir / "step2.json").unlink(missing_ok=True)
            (job_dir / "error.png").unlink(missing_ok=True)
            for old in job_dir.glob("result_*.xlsx"):
                old.unlink(missing_ok=True)
            StepReporter(job_dir, 1).running("Queued...", 0.0)
            start_thread(retrieve_worker, job_dir, username, password, start_d, end_d)
            st.session_state["_was_busy"] = True
            st.rerun()

    st.divider()

    # ------------------------------------------------------------ STEP 2
    st.subheader("Step 2 · Filter Excel")
    st.caption("Streams the collected rows and keeps only the chosen Year-Month, then builds the Excel file. "
               "Nothing is downloaded here.")
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
            on_click="ignore",
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
    st.fragment(run_every=POLL_SECONDS if busy else None)(render_workflow)(job_dir, job_id)

    st.caption("Your credentials are used only to start Step 1 and are never written to disk or logs.")


if __name__ == "__main__":
    main()
