"""
PMT Alarm – Closed Alarm downloader (3-step workflow)

STEP 1  Retrieve Excel  -> login + download the raw XLSX only (background thread)
STEP 2  Filter Excel    -> stream the raw XLSX, keep the chosen Year-Month (background thread)
STEP 3  Download Result -> hand the processed XLSX to the browser (no processing)

All state lives on disk in a per-session job folder (not only in st.session_state),
so it survives websocket reconnects and page refreshes (the job id is kept in the URL).
"""
import json
import logging
import os
import re
import shutil
import tempfile
import threading
import time
import traceback
import uuid
from datetime import date, datetime
from pathlib import Path

import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel
from selenium import webdriver
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    StaleElementReferenceException,
    TimeoutException,
)
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# ===================== CONFIGURATION =====================
LOGIN_URL = "https://pmt-alarm.komdigi.go.id/auth/login"
ALARM_URL = "https://pmt-alarm.komdigi.go.id/dashboard/alarm"
DOWNLOAD_TIMEOUT = 900        # seconds to wait for the big export (15 min)
DEFAULT_TIMEOUT = 40
DATE_COLUMN_NAME = "Data Created"
JOBS_ROOT = Path(tempfile.gettempdir()) / "pmt_alarm_jobs"
STALE_AFTER = 300             # a "running" step with no heartbeat for 5 min = interrupted
JOB_MAX_AGE_HOURS = 12
POLL_SECONDS = 3
MONTHS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
# =========================================================

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger("pmt_alarm")


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

    def running(self, message, progress=None):
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
    return job_dir / "source.xlsx"


def source_ready(job_dir: Path) -> bool:
    p = source_path(job_dir)
    return p.exists() and p.stat().st_size > 0


# =====================================================================
#  Selenium helpers (same logic as before)
# =====================================================================
def build_driver(download_dir: Path) -> webdriver.Chrome:
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--window-size=1600,900")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    # keep Chromium's memory footprint small
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-background-networking")
    options.add_argument("--disable-sync")
    options.add_argument("--mute-audio")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-popup-blocking")
    options.add_experimental_option("excludeSwitches", ["enable-logging"])

    chromium = shutil.which("chromium") or shutil.which("chromium-browser")
    if chromium:
        options.binary_location = chromium

    options.add_experimental_option(
        "prefs",
        {
            "download.default_directory": str(download_dir),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": False,  # avoids scanning a 90 MB file
            "profile.default_content_setting_values.automatic_downloads": 1,
            "profile.managed_default_content_settings.images": 2,  # no images
        },
    )

    chromedriver = shutil.which("chromedriver")
    if chromedriver:  # Streamlit Cloud (packages.txt)
        service = Service(chromedriver)
    else:  # local PC with regular Google Chrome
        from webdriver_manager.chrome import ChromeDriverManager

        service = Service(ChromeDriverManager().install())

    driver = webdriver.Chrome(service=service, options=options)
    try:
        driver.execute_cdp_cmd(
            "Page.setDownloadBehavior",
            {"behavior": "allow", "downloadPath": str(download_dir)},
        )
    except Exception as exc:
        log.debug("setDownloadBehavior not applied: %s", exc)
    return driver


def safe_click(driver, locator, timeout=DEFAULT_TIMEOUT, description="element"):
    wait = WebDriverWait(driver, timeout)
    last_exc = None
    for attempt in range(1, 4):
        try:
            element = wait.until(EC.element_to_be_clickable(locator))
            driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});", element
            )
            time.sleep(0.5)
            try:
                element.click()
            except (ElementClickInterceptedException, StaleElementReferenceException):
                element = driver.find_element(*locator)
                driver.execute_script("arguments[0].click();", element)
            return element
        except (TimeoutException, StaleElementReferenceException) as exc:
            last_exc = exc
            log.warning("Attempt %d to click %s failed: %s", attempt, description, type(exc).__name__)
            time.sleep(1)
    raise TimeoutException(f"Could not click {description}: {last_exc}")


def first_clickable(driver, locators, timeout=DEFAULT_TIMEOUT, description="element"):
    per_locator_timeout = max(5, timeout // max(1, len(locators)))
    last_exc = None
    for loc in locators:
        try:
            return safe_click(driver, loc, timeout=per_locator_timeout, description=description)
        except TimeoutException as exc:
            last_exc = exc
    raise TimeoutException(f"None of the locators worked for {description}: {last_exc}")


def login(driver, username, password):
    driver.get(LOGIN_URL)
    wait = WebDriverWait(driver, DEFAULT_TIMEOUT)
    user_field = wait.until(
        EC.visibility_of_element_located(
            (By.XPATH, "//input[@placeholder='Username' or @name='username' or @id='username']")
        )
    )
    user_field.clear()
    user_field.send_keys(username)
    pass_field = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@type='password']")))
    pass_field.clear()
    pass_field.send_keys(password)

    first_clickable(
        driver,
        [
            (By.XPATH, "//button[normalize-space()='LOGIN' or normalize-space()='Login']"),
            (By.XPATH, "//button[contains(translate(., 'login', 'LOGIN'), 'LOGIN')]"),
            (By.XPATH, "//input[@type='submit']"),
            (By.CSS_SELECTOR, "button[type='submit']"),
        ],
        description="LOGIN button",
    )
    try:
        WebDriverWait(driver, 60).until(lambda d: "/auth/login" not in d.current_url)
    except TimeoutException:
        raise RuntimeError(
            "Login failed: still on the login page. Check username/password, "
            "or the site may require CAPTCHA/OTP or block this server's IP."
        )


def open_alarm_page(driver):
    driver.get(ALARM_URL)
    WebDriverWait(driver, 60).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )
    if "/auth/login" in driver.current_url:
        raise RuntimeError("Redirected back to login page; session not authenticated.")
    WebDriverWait(driver, DEFAULT_TIMEOUT).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    time.sleep(2)


def scroll_to_alarm_section(driver):
    wait = WebDriverWait(driver, DEFAULT_TIMEOUT)
    try:
        section = wait.until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "//*[self::h1 or self::h2 or self::h3 or self::h4 or self::h5 or self::div or self::span]"
                    "[normalize-space()='Alarm' or normalize-space()='ALARM']",
                )
            )
        )
        driver.execute_script("arguments[0].scrollIntoView({block: 'start'});", section)
    except TimeoutException:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(1)


def click_closed_alarm_tab(driver):
    first_clickable(
        driver,
        [
            (
                By.XPATH,
                "//*[(self::button or self::a or self::li or self::div or self::span or @role='tab')]"
                "[normalize-space()='CLOSED ALARM' or normalize-space()='Closed Alarm']",
            ),
            (
                By.XPATH,
                "//*[contains(translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', "
                "'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), 'CLOSED ALARM')"
                " and (self::button or self::a or @role='tab' or self::li)]",
            ),
        ],
        description="CLOSED ALARM tab",
    )
    wait = WebDriverWait(driver, 60)
    wait.until(EC.presence_of_element_located((By.XPATH, "//table")))
    try:
        wait.until(EC.presence_of_element_located((By.XPATH, "//table//tbody//tr")))
    except TimeoutException:
        pass
    try:
        WebDriverWait(driver, 10).until(
            EC.invisibility_of_element_located(
                (
                    By.CSS_SELECTOR,
                    ".spinner, .loading, .loader, .spinner-border, [class*='loading'], [class*='spinner']",
                )
            )
        )
    except TimeoutException:
        pass
    time.sleep(1)


def click_download_and_all_page(driver):
    upper = "translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ')"
    first_clickable(
        driver,
        [
            (By.XPATH, f"//button[contains({upper}, 'DOWNLOAD ALARM')]"),
            (By.XPATH, f"//*[contains({upper}, 'DOWNLOAD ALARM') and (self::a or self::button or @role='button')]"),
            (By.CSS_SELECTOR, "button.btn-success"),
        ],
        description="Download Alarm button",
    )
    first_clickable(
        driver,
        [
            (
                By.XPATH,
                "//*[(self::a or self::button or self::li or self::span or self::div "
                "or @role='menuitem') and (normalize-space()='All Page' or normalize-space()='All page')]",
            ),
            (
                By.XPATH,
                f"//*[contains({upper}, 'ALL PAGE') and (self::a or self::button or self::li or @role='menuitem')]",
            ),
        ],
        description="All Page option",
    )


def snapshot_files(folder: Path):
    return {p.name: p.stat().st_mtime for p in folder.iterdir() if p.is_file()}


def wait_for_download_complete(folder: Path, before: dict, on_tick, timeout=DOWNLOAD_TIMEOUT) -> str:
    """Wait for a new, size-stable .xlsx. Calls on_tick(elapsed_s, size_mb) every ~3 s."""
    start = time.time()
    last_tick = 0.0
    last_size, stable = {}, 0

    while time.time() - start < timeout:
        files = [p for p in folder.iterdir() if p.is_file()]
        in_progress = [p for p in files if p.suffix.lower() in (".crdownload", ".tmp")]
        new_files = [
            p for p in files
            if p.suffix.lower() in (".xlsx", ".xls")
            and not p.name.startswith("~$")
            and (p.name not in before or p.stat().st_mtime > before[p.name])
        ]

        if time.time() - last_tick >= 3:
            last_tick = time.time()
            current = in_progress + new_files
            size_mb = max((p.stat().st_size for p in current), default=0) / 1024 / 1024
            on_tick(time.time() - start, size_mb)

        if new_files and not in_progress:
            newest = max(new_files, key=lambda p: p.stat().st_mtime)
            size = newest.stat().st_size
            stable = stable + 1 if (last_size.get(newest.name) == size and size > 0) else 0
            last_size[newest.name] = size
            if stable >= 3:
                return newest.name
        else:
            stable = 0
        time.sleep(1)

    raise TimeoutException(f"Download did not complete within {timeout} seconds.")


# =====================================================================
#  STEP 1 worker – retrieve (download only, no processing)
# =====================================================================
def retrieve_worker(job_dir: Path, username: str, password: str):
    rep = StepReporter(job_dir, 1)
    dl_dir = job_dir / "dl"
    shutil.rmtree(dl_dir, ignore_errors=True)
    dl_dir.mkdir(parents=True, exist_ok=True)
    driver = None
    try:
        rep.running("Starting browser...", 0.05)
        driver = build_driver(dl_dir)

        rep.running("Logging in...", 0.15)
        login(driver, username, password)

        rep.running("Opening alarm page...", 0.25)
        open_alarm_page(driver)
        scroll_to_alarm_section(driver)

        rep.running("Opening CLOSED ALARM tab...", 0.30)
        click_closed_alarm_tab(driver)

        before = snapshot_files(dl_dir)
        rep.running("Requesting full export (All Page)...", 0.35)
        click_download_and_all_page(driver)

        def tick(elapsed, size_mb):
            mins, secs = divmod(int(elapsed), 60)
            rep.running(
                f"Waiting for the server to prepare/send the file... "
                f"{mins}m {secs:02d}s elapsed, {size_mb:.1f} MB received so far",
                0.40,
            )

        filename = wait_for_download_complete(dl_dir, before, tick)

        rep.running("Download finished, closing browser...", 0.95)
        driver.quit()
        driver = None

        os.replace(dl_dir / filename, source_path(job_dir))
        size_mb = source_path(job_dir).stat().st_size / 1024 / 1024
        rep.done(
            f"Download complete: {size_mb:.1f} MB saved on the server.",
            size_mb=round(size_mb, 1),
            original_name=filename,
        )
    except Exception as exc:
        job_log(job_dir, traceback.format_exc(), "ERROR")
        if driver is not None:
            try:
                (job_dir / "error.png").write_bytes(driver.get_screenshot_as_png())
            except Exception:
                pass
        rep.fail(str(exc) or type(exc).__name__)
    finally:
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass
        shutil.rmtree(dl_dir, ignore_errors=True)


# =====================================================================
#  STEP 2 worker – streaming filter (low memory)
# =====================================================================
_DATE_FORMATS = [
    "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M:%S.%f", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
    "%d/%m/%Y %H:%M:%S", "%d-%m-%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d-%m-%Y %H:%M", "%d/%m/%Y", "%d-%m-%Y",
]


class DateParser:
    """Fast row-by-row date parsing; remembers the format that worked."""

    def __init__(self):
        self.fmt = None

    def __call__(self, value):
        if value is None:
            return None
        if isinstance(value, datetime):
            return value
        if isinstance(value, date):
            return datetime(value.year, value.month, value.day)
        if isinstance(value, (int, float)):
            try:
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
        for fmt in _DATE_FORMATS:
            try:
                parsed = datetime.strptime(s, fmt)
                self.fmt = fmt
                return parsed
            except ValueError:
                continue
        try:  # rare formats
            import pandas as pd

            iso = len(s) >= 5 and s[:4].isdigit() and s[4] in "-/"
            ts = pd.to_datetime(s, errors="coerce", dayfirst=not iso)
            return None if pd.isna(ts) else ts.to_pydatetime()
        except Exception:
            return None


def _clean(v):
    return ILLEGAL_CHARACTERS_RE.sub("", v) if isinstance(v, str) else v


def stream_filter(src: Path, dst: Path, year: int, month: int, on_progress):
    """
    Read `src` row by row (openpyxl read-only) and write matching rows straight to `dst`
    (openpyxl write-only). Memory stays small no matter how big the file is.
    Returns dict(total, kept, invalid, date_col).
    """
    parser = DateParser()
    wb = load_workbook(src, read_only=True, data_only=True)
    try:
        ws = wb.worksheets[0]
        try:
            est_total = ws.max_row  # from the sheet header; may be None / approximate
        except Exception:
            est_total = None

        rows = ws.iter_rows(values_only=True)
        header = next(rows, None)
        if not header or all(h is None for h in header):
            raise RuntimeError("The downloaded Excel file is empty or has no header row.")
        ncols = max(i for i, h in enumerate(header) if h is not None) + 1
        header = list(header[:ncols])

        normalized = {str(h).strip().lower(): i for i, h in enumerate(header) if h is not None}
        date_idx = None
        for cand in (DATE_COLUMN_NAME.lower(), "date created", "data created"):
            if cand in normalized:
                date_idx = normalized[cand]
                break
        if date_idx is None:
            date_idx = ncols - 1  # fall back to the LAST column
        date_col = str(header[date_idx])

        wb_out = Workbook(write_only=True)
        ws_out = wb_out.create_sheet("Closed Alarm")
        ws_out.freeze_panes = "A2"
        for i, h in enumerate(header, start=1):
            ws_out.column_dimensions[get_column_letter(i)].width = min(max(len(str(h or "")) + 4, 14), 40)
        ws_out.append([_clean(h) for h in header])

        total = kept = invalid = 0
        last_report = time.time()
        for row in rows:
            if not any(v is not None for v in row):
                continue  # blank row
            total += 1
            if len(row) < ncols:
                row = tuple(row) + (None,) * (ncols - len(row))
            dt = parser(row[date_idx])
            if dt is None:
                invalid += 1
            elif dt.year == year and dt.month == month:
                out = [_clean(v) for v in row[:ncols]]
                out[date_idx] = dt
                ws_out.append(out)
                kept += 1
            if total % 2000 == 0 and time.time() - last_report >= 2:
                last_report = time.time()
                on_progress(total, est_total, kept)

        wb_out.save(dst)
        return {"total": total, "kept": kept, "invalid": invalid, "date_col": date_col}
    finally:
        wb.close()


def filter_worker(job_dir: Path, year: int, month: int):
    rep = StepReporter(job_dir, 2)
    label = f"{MONTHS[month - 1]} {year}"
    part = job_dir / "result.part"
    try:
        src = source_path(job_dir)
        if not source_ready(job_dir):
            raise RuntimeError("Source Excel not found. Run Step 1 first.")

        for old in job_dir.glob("result_*.xlsx"):
            old.unlink(missing_ok=True)

        rep.running(f"Reading source file and filtering for {label}...", 0.02)

        def progress(total, est, kept):
            frac = min(total / est, 0.98) if est and est > 0 else None
            rep.running(f"Filtering {label}: {total:,} rows scanned, {kept:,} kept so far...", frac)

        stats = stream_filter(src, part, year, month, progress)

        result_name = f"result_{year}-{month:02d}.xlsx"
        os.replace(part, job_dir / result_name)
        rep.done(
            f"Filtering complete for {label}: {stats['kept']:,} of {stats['total']:,} rows kept.",
            result_file=result_name,
            label=label,
            year=year,
            month=month,
            **stats,
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
    st.caption("Logs in and downloads the full Closed Alarm export. No filtering happens here.")
    if src_ok and s1["state"] != "running":
        size_mb = source_path(job_dir).stat().st_size / 1024 / 1024
        st.success(f"✅ Source file is ready on the server ({size_mb:.1f} MB). No need to download again.")
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
            st.info("The source file already exists, so the download was skipped. "
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
    st.caption("Streams the downloaded file and keeps only the chosen Year-Month. Nothing is downloaded here.")
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
    with st.expander("Activity log"):
        log_file = job_dir / "job.log"
        lines = log_file.read_text(encoding="utf-8").splitlines()[-60:] if log_file.exists() else []
        st.code("\n".join(lines) or "(empty)", language=None)
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
