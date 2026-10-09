"""
PMT Alarm – Closed Alarm downloader (2 visible steps, Playwright edition)

STEP 1  Retrieve   -> login in Chromium, open CLOSED ALARM, fill Start/End Date, Apply, set Items per page (500),
                      then for every page: Download Alarm > By Page (one file per page), click next page, repeat.
                      All page files are compiled into ONE Excel file.
STEP 3  Download   -> hand the compiled XLSX to the browser (no processing)

All state lives on disk in a per-session job folder (not only in st.session_state),
so it survives websocket reconnects and page refreshes (the job id is kept in the URL).
"""
import csv
import faulthandler
import hashlib
import io
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
STALE_AFTER = 600             # a "running" step with no heartbeat for 5 min = interrupted
JOB_MAX_AGE_HOURS = 12
POLL_SECONDS = 3
MEM_GUARD_FRACTION = 0.85     # kill Chromium when the container reaches 85 % of its memory limit
MEM_LIMIT_MB_OVERRIDE = float(os.environ.get("MEM_LIMIT_MB", 0)) or None
MONITOR_INTERVAL = 2          # seconds between diagnostic samples
BLOCK_IMAGES = True

SITE_UTC_OFFSET_HOURS = 7     # the site works in WIB (UTC+7)
DATE_INPUT_FORMAT = "%d-%m-%Y"        # how the Start/End Date boxes display dates (02-10-2026)
PAGE_SIZE_CHOICES = ("500", "250", "100", "50")  # tried in this order in the "Items" dropdown (site offers 10/50/100/250/500)
MAX_PAGES = 3000              # safety cap for the pagination loop
DOWNLOAD_TIMEOUT = 240        # seconds to wait for ONE "By Page" file
PAGE_SIZE_LOAD_TIMEOUT = 60   # seconds to wait for the table to reload after choosing a bigger page size

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


def pw_click_closed_alarm_tab(page, job_dir: Path):
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
    job_log(job_dir, f"Closed Alarm table detected, columns: {headers}")
    wait_loading_done(page)
    page.wait_for_timeout(800)


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


def pw_set_date(page, job_dir: Path, label: str, value: date):
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


# Finds the option (e.g. "500") inside the dropdown list that popped up next to the "Items" box.
# Works with <li>, <div>, <span>, <mat-option>, role=option ... : any visible element whose text is exactly the size,
# which is not inside the table and sits near the Items dropdown. The deepest matching element wins.
MARK_OPTION_JS = r"""(size) => {
  document.querySelectorAll('[data-pw-option]').forEach(e => e.removeAttribute('data-pw-option'));
  const trig = document.querySelector('[data-pw-target="items"]');
  if (!trig) return 'no-trigger';
  const tr = trig.getBoundingClientRect();
  const visible = el => {
    const r = el.getBoundingClientRect();
    const s = getComputedStyle(el);
    return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none';
  };
  const cands = Array.from(document.querySelectorAll('body *')).filter(el => {
    if (el === trig || trig.contains(el) || el.closest('table')) return false;
    if (!visible(el)) return false;
    const t = (el.innerText || el.textContent || '').replace(/\s+/g, ' ').trim();
    if (t !== size) return false;
    const r = el.getBoundingClientRect();
    return Math.abs(r.left - tr.left) < 250 && Math.abs((r.top + r.bottom) / 2 - (tr.top + tr.bottom) / 2) < 600;
  });
  if (!cands.length) return 'no-option';
  const deepest = cands.filter(el => !cands.some(o => o !== el && el.contains(o)));
  const pick = deepest[0] || cands[0];
  pick.setAttribute('data-pw-option', '1');
  return 'ok';
}"""


def _click_page_size_option(page, size: str, timeout=8):
    """Wait for the popup list and click the option whose text is exactly `size`."""
    end = time.time() + timeout
    res = ""
    while time.time() < end:
        res = page.evaluate(MARK_OPTION_JS, size)
        if res == "ok":
            safe_click(page.locator("[data-pw-option='1']").first)
            return
        page.wait_for_timeout(300)
    raise RuntimeError(f"option '{size}' not found in the Items dropdown (result: {res})")


def pw_set_page_size(page, job_dir: Path):
    """Pick 500 (or 250 / 100 / 50 as fallback) in the 'Items' dropdown. Returns the number of rows now shown."""
    _, rows0 = read_table(page)
    n0 = len(rows0)
    if n0 and n0 < 10:
        return n0  # everything already fits on one page
    old = table_sig(rows0)
    for size in PAGE_SIZE_CHOICES:
        try:
            trigger = mark_control(page, "items", 20)
            trigger.wait_for(state="visible", timeout=8000)
            trigger.scroll_into_view_if_needed()
            if trigger.evaluate("el => el.tagName") == "SELECT":
                try:
                    trigger.select_option(label=size)
                except PWError:
                    trigger.select_option(value=size)
            else:
                safe_click(trigger)                     # open the dropdown list (10 / 50 / 100 / 250 / 500)
                page.wait_for_timeout(500)
                _click_page_size_option(page, size)     # click the "500" entry
            job_log(job_dir, f"Clicked '{size}' in the Items dropdown, waiting for the table to reload...")
            wait_table_change(page, old, PAGE_SIZE_LOAD_TIMEOUT)
            wait_loading_done(page)
            page.wait_for_timeout(800)
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
#  Download Alarm > By Page  (one file per page)
# =====================================================================
BY_PAGE_XPATH = (f"xpath=//*[(self::button or self::a or self::li or self::span or self::div "
                 f"or @role='menuitem') and {UPPER}='BY PAGE']")


def find_next(page):
    handle = page.evaluate_handle(FIND_NEXT_JS)
    el = handle.as_element()
    if el is None:
        return None, True
    return el, bool(el.evaluate(IS_DISABLED_JS))


def file_sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def pw_download_by_page(page, job_dir: Path, dest_base: Path) -> Path:
    """Open 'Download Alarm', click 'By Page', and save the file the browser downloads."""
    last_exc = None
    for attempt in (1, 2, 3):
        try:
            try:
                page.keyboard.press("Escape")  # make sure no old menu is open
            except PWError:
                pass
            with page.expect_download(timeout=DOWNLOAD_TIMEOUT * 1000) as info:
                safe_click(first_visible(page, f"xpath=//button[contains({UPPER}, 'DOWNLOAD ALARM')]", 20))
                page.wait_for_timeout(600)
                safe_click(first_visible(page, BY_PAGE_XPATH, 15))
            dl = info.value
            failure = dl.failure()
            if failure:
                raise RuntimeError(f"browser download failed: {failure}")
            ext = Path(dl.suggested_filename or "").suffix.lower() or ".xlsx"
            dest = Path(f"{dest_base}{ext}")
            dl.save_as(str(dest))
            if not dest.exists() or dest.stat().st_size == 0:
                raise RuntimeError("the downloaded file is empty")
            return dest
        except (PWError, RuntimeError) as exc:
            last_exc = exc
            job_log(job_dir, f"By Page download attempt {attempt} failed: {str(exc)[:200]}", "WARNING")
            page.wait_for_timeout(2000)
    raise RuntimeError(f"Download Alarm > By Page failed 3 times: {str(last_exc)[:200]}")


def download_all_pages(page, job_dir: Path, pages_dir: Path, rep: StepReporter, monitor, browser):
    files, seen_hashes, page_no = [], {}, 0
    _, rows = read_table(page)
    if not rows:
        raise RuntimeError("The table is empty for this date range (no closed alarms found).")
    total_pages = max(1, page.evaluate(TOTAL_PAGES_JS))

    while True:
        page_no += 1
        total_pages = max(total_pages, page_no, page.evaluate(TOTAL_PAGES_JS))
        rep.running(f"Page {page_no} of ~{total_pages}: downloading 'By Page' ({len(rows)} rows on screen)...",
                    0.40 + 0.50 * min((page_no - 1) / total_pages, 1.0))
        dest = pw_download_by_page(page, job_dir, pages_dir / f"page_{page_no:04d}")
        digest = file_sha256(dest)
        if digest in seen_hashes:
            job_log(job_dir, f"Page {page_no}: the file is identical to page {seen_hashes[digest]}; skipped.", "WARNING")
            dest.unlink(missing_ok=True)
        else:
            seen_hashes[digest] = page_no
            files.append(dest)
            job_log(job_dir, f"Page {page_no}: saved {dest.name} ({dest.stat().st_size / 1024:.0f} KB)")

        if monitor is not None and monitor.tripped:
            raise RuntimeError(monitor.tripped)
        if not browser.is_connected() or page.is_closed():
            raise RuntimeError("Chromium exited while downloading pages.")
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
        changed = False
        for attempt in (1, 2):
            safe_click(nxt)
            if wait_table_change(page, old, 40):
                changed = True
                break
            job_log(job_dir, f"Page {page_no}: table did not change after clicking next (attempt {attempt}).", "WARNING")
            nxt, disabled = find_next(page)
            if nxt is None or disabled:
                break
        if not changed:
            job_log(job_dir, "Table stopped changing: assuming the last page was reached.")
            break
        _, rows = read_table(page)
        if not rows:
            break
    return files, page_no


# =====================================================================
#  Compile all page files into one workbook
# =====================================================================
def _iter_xlsx(path: Path):
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        for row in wb.worksheets[0].iter_rows(values_only=True):
            yield list(row)
    finally:
        wb.close()


def _iter_csv(path: Path):
    raw = path.read_bytes()
    text = None
    for enc in ("utf-8-sig", "utf-16", "cp1252", "latin-1"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeError:
            continue
    try:
        dialect = csv.Sniffer().sniff(text[:4096], delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel
    yield from csv.reader(io.StringIO(text), dialect)


def iter_file_rows(path: Path):
    ext = path.suffix.lower()
    if ext == ".xls":
        raise RuntimeError("The site returned an old .xls file, which this app cannot read. "
                           "Install xlrd or ask me to add .xls support.")
    if ext in (".csv", ".txt"):
        return _iter_csv(path)
    return _iter_xlsx(path)


def split_header(rows_iter):
    """First row with at least 3 filled cells is the header. Returns (header_list, rest_iterator)."""
    for row in rows_iter:
        vals = ["" if c is None else str(c).strip() for c in row]
        if sum(1 for v in vals if v) >= 3:
            return vals, rows_iter
    return None, rows_iter


def unique_names(header):
    seen, out = {}, []
    for i, h in enumerate(header):
        h = h or f"col_{i + 1}"
        n = seen.get(h, 0)
        seen[h] = n + 1
        out.append(h if n == 0 else f"{h}_{n + 1}")
    return out


def compile_files(files, out_path: Path, job_dir: Path, rep: StepReporter) -> int:
    """Merge every page file into one XLSX (columns matched by header name). Returns the row count."""
    master, seen, headers = [], set(), []
    for f in files:                                   # pass 1: collect the columns
        gen = iter_file_rows(f)
        header, _ = split_header(gen)
        if hasattr(gen, "close"):
            gen.close()
        if header is None:
            raise RuntimeError(f"Could not find a header row in {f.name}.")
        names = unique_names(header)
        headers.append(names)
        for n in names:
            if n not in seen:
                seen.add(n)
                master.append(n)

    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Closed Alarm")
    ws.freeze_panes = "A2"
    for i, label in enumerate(master, start=1):
        ws.column_dimensions[get_column_letter(i)].width = min(max(len(label) + 4, 14), 40)
    ws.append(master)

    total = 0
    for idx, f in enumerate(files):                   # pass 2: write the rows
        names = headers[idx]
        pos = {n: i for i, n in enumerate(names)}
        order = [pos.get(m) for m in master]
        gen = iter_file_rows(f)
        _, rest = split_header(gen)
        count = 0
        for row in rest:
            if not any(c not in (None, "") for c in row):
                continue
            out = []
            for p in order:
                v = row[p] if p is not None and p < len(row) else None
                out.append(ILLEGAL_CHARACTERS_RE.sub("", v) if isinstance(v, str) else v)
            ws.append(out)
            count += 1
        if hasattr(gen, "close"):
            gen.close()
        total += count
        job_log(job_dir, f"Compiled {f.name}: {count:,} rows")
        rep.running(f"Compiling files into one Excel: {idx + 1}/{len(files)} done, {total:,} rows so far...",
                    0.90 + 0.08 * (idx + 1) / len(files), quiet=True)
    wb.save(out_path)
    return total


# =====================================================================
#  STEP 1 worker – filters, per-page downloads, compile
# =====================================================================
def retrieve_worker(job_dir: Path, username: str, password: str, start_d: date, end_d: date):
    rep = StepReporter(job_dir, 1)
    dl_dir = job_dir / "dl"
    pages_dir = job_dir / "pages"
    for d in (dl_dir, pages_dir):
        shutil.rmtree(d, ignore_errors=True)
        d.mkdir(parents=True, exist_ok=True)
    for pattern in ("shot_*.png", "error.png", "result_*.xlsx", "result.part"):
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
                context = browser.new_context(accept_downloads=True, viewport={"width": 1600, "height": 1000})
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

                monitor.phase = "login"
                rep.running("Logging in...", 0.10)
                pw_login(page, username, password)

                monitor.phase = "alarm-page"
                rep.running("Opening alarm page...", 0.16)
                pw_open_alarm_page(page)
                pw_scroll_to_alarm_section(page)

                rep.running("Opening CLOSED ALARM tab...", 0.20)
                pw_click_closed_alarm_tab(page, job_dir)
                take_shot(page, job_dir, "shot_0_closed_alarm_tab")

                monitor.phase = "filters"
                rep.running(f"Setting Start Date {start_d:%d-%m-%Y} and End Date {end_d:%d-%m-%Y}...", 0.24)
                pw_set_date(page, job_dir, "Start Date", start_d)
                pw_set_date(page, job_dir, "End Date", end_d)
                take_shot(page, job_dir, "shot_1_filters")

                _, rows_before = read_table(page)
                rep.running("Applying the filter...", 0.28)
                pw_click_apply(page)
                wait_table_change(page, table_sig(rows_before), 25)
                wait_loading_done(page)
                page.wait_for_timeout(800)

                rep.running("Setting items per page to 500...", 0.33)
                pw_set_page_size(page, job_dir)
                take_shot(page, job_dir, "shot_2_after_page_size")

                monitor.phase = "download-pages"
                rep.running("Downloading page by page...", 0.40)
                files, pages = download_all_pages(page, job_dir, pages_dir, rep, monitor, browser)
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

        if not files:
            raise RuntimeError("No page files were downloaded.")
        monitor.phase = "compile"
        rep.running(f"Compiling {len(files)} downloaded files into one Excel...", 0.90)
        result_name = f"closed_alarm_{start_d:%Y%m%d}_{end_d:%Y%m%d}.xlsx"
        part = job_dir / "result.part"
        total_rows = compile_files(files, part, job_dir, rep)
        os.replace(part, job_dir / result_name)
        size_mb = (job_dir / result_name).stat().st_size / 1024 / 1024
        rep.done(f"Compiled {total_rows:,} rows from {len(files)} page files "
                 f"({start_d:%d-%m-%Y} to {end_d:%d-%m-%Y}, {size_mb:.1f} MB).",
                 result_file=result_name, rows=total_rows, pages=len(files),
                 range_start=start_d.isoformat(), range_end=end_d.isoformat())
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
        shutil.rmtree(pages_dir, ignore_errors=True)
        (job_dir / "result.part").unlink(missing_ok=True)


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
    s1 = get_status(job_dir, 1)
    busy = s1["state"] == "running"

    was_busy = st.session_state.get("_was_busy", False)
    st.session_state["_was_busy"] = busy
    if was_busy and not busy:
        st.rerun()

    result_path = job_dir / s1["result_file"] if s1.get("state") == "done" and s1.get("result_file") else None
    result_ok = result_path is not None and result_path.exists()

    # ------------------------------------------------------------ STEP 1
    st.subheader("Step 1 · Retrieve closed alarms")
    st.caption("Logs in, opens CLOSED ALARM, fills Start/End Date, sets 500 items per page, then clicks "
               "Download Alarm › By Page on every page and compiles all files into one Excel.")
    show_status(s1)
    if s1["state"] == "error" and (job_dir / "error.png").exists():
        st.image(str(job_dir / "error.png"), caption="Browser screenshot when the error happened")

    d0, d1 = default_range()
    with st.form("step1_form"):
        username = st.text_input("Username", disabled=busy)
        password = st.text_input("Password", type="password", disabled=busy)
        cs, ce = st.columns(2)
        start_d = cs.date_input("Start date", value=d0, format="DD/MM/YYYY", disabled=busy, key="rng_start")
        end_d = ce.date_input("End date", value=d1, format="DD/MM/YYYY", disabled=busy, key="rng_end")
        force = st.checkbox("Replace the existing result (retrieve again)", disabled=busy or not result_ok)
        go1 = st.form_submit_button("Retrieve closed alarms", type="primary", disabled=busy, use_container_width=True)

    if go1:
        if result_ok and not force:
            st.info("A compiled result already exists, so retrieving was skipped. "
                    "Tick “Replace the existing result” to retrieve a new copy.")
        elif not username or not password:
            st.error("Please enter both username and password.")
        elif start_d > end_d:
            st.error("Start date must not be after end date.")
        else:
            (job_dir / "error.png").unlink(missing_ok=True)
            for old in job_dir.glob("closed_alarm_*.xlsx"):
                old.unlink(missing_ok=True)
            StepReporter(job_dir, 1).running("Queued...", 0.0)
            start_thread(retrieve_worker, job_dir, username, password, start_d, end_d)
            st.session_state["_was_busy"] = True
            st.rerun()

    st.divider()

    # ------------------------------------------------------------ STEP 3
    st.subheader("Step 3 · Download Result")
    st.caption("Just hands you the compiled file from Step 1. No processing happens here.")
    if result_ok:
        rs, re_ = s1.get("range_start"), s1.get("range_end")
        st.download_button(
            "Download Result",
            data=result_path.read_bytes(),
            file_name=result_path.name,
            mime=XLSX_MIME,
            type="primary",
            use_container_width=True,
            on_click="ignore",
        )
        st.caption(f"{s1.get('rows', '?'):,} rows · {s1.get('pages', '?')} pages · "
                   f"{result_path.stat().st_size / 1024 / 1024:.2f} MB"
                   + (f" · {date.fromisoformat(rs):%d-%m-%Y} to {date.fromisoformat(re_):%d-%m-%Y}" if rs and re_ else "")
                   if isinstance(s1.get("rows"), int) else f"{result_path.stat().st_size / 1024 / 1024:.2f} MB")
    else:
        st.button("Download Result", disabled=True, use_container_width=True, key="dl_disabled")
        st.caption("🔒 Complete Step 1 first.")

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

    busy = get_status(job_dir, 1)["state"] == "running"
    st.fragment(run_every=POLL_SECONDS if busy else None)(render_workflow)(job_dir, job_id)

    st.caption("Your credentials are used only to start Step 1 and are never written to disk or logs.")


if __name__ == "__main__":
    main()
