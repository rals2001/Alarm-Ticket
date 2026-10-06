import io
import logging
import shutil
import tempfile
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from selenium import webdriver
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    StaleElementReferenceException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# ===================== CONFIGURATION =====================
LOGIN_URL = "https://pmt-alarm.komdigi.go.id/auth/login"
ALARM_URL = "https://pmt-alarm.komdigi.go.id/dashboard/alarm"
DOWNLOAD_TIMEOUT = 600
DEFAULT_TIMEOUT = 40
DATE_COLUMN_NAME = "Data Created"
MONTHS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]
# =========================================================

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger("pmt_alarm")


class PipelineError(Exception):
    """Error that can carry a screenshot of the browser for debugging."""

    def __init__(self, message, screenshot=None):
        super().__init__(message)
        self.screenshot = screenshot


# ---------------------------------------------------------------- browser
def build_driver(download_dir: Path) -> webdriver.Chrome:
    options = Options()
    options.add_argument("--headless=new")  # always headless on a server
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
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
            "safebrowsing.enabled": True,
            "profile.default_content_setting_values.automatic_downloads": 1,
        },
    )

    chromedriver = shutil.which("chromedriver")
    if chromedriver:  # Streamlit Cloud / Linux with chromium-driver installed
        service = Service(chromedriver)
    else:  # local machine: download a matching driver automatically
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


# ---------------------------------------------------------------- site steps
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
            "Login failed: still on the login page. Check username/password "
            "or any CAPTCHA/OTP requirement."
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
            (
                By.XPATH,
                f"//*[contains({upper}, 'DOWNLOAD ALARM') and (self::a or self::button or @role='button')]",
            ),
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
                f"//*[contains({upper}, 'ALL PAGE') "
                "and (self::a or self::button or self::li or @role='menuitem')]",
            ),
        ],
        description="All Page option",
    )


def snapshot_files(folder: Path):
    return {p.name: p.stat().st_mtime for p in folder.iterdir() if p.is_file()}


def wait_for_download_complete(folder: Path, before: dict, timeout: int = DOWNLOAD_TIMEOUT) -> str:
    end_time = time.time() + timeout
    last_size, stable_count = {}, 0

    while time.time() < end_time:
        files = [p for p in folder.iterdir() if p.is_file()]
        in_progress = [p for p in files if p.suffix.lower() in (".crdownload", ".tmp")]
        new_files = [
            p
            for p in files
            if p.suffix.lower() in (".xlsx", ".xls")
            and not p.name.startswith("~$")
            and (p.name not in before or p.stat().st_mtime > before[p.name])
        ]

        if new_files and not in_progress:
            newest = max(new_files, key=lambda p: p.stat().st_mtime)
            size = newest.stat().st_size
            if last_size.get(newest.name) == size and size > 0:
                stable_count += 1
            else:
                stable_count = 0
            last_size[newest.name] = size
            if stable_count >= 3:
                return newest.name
        else:
            stable_count = 0
        time.sleep(1)

    raise TimeoutException(f"Download did not complete within {timeout} seconds.")


# ---------------------------------------------------------------- filtering
def parse_datetime_column(series: pd.Series) -> pd.Series:
    if pd.api.types.is_datetime64_any_dtype(series):
        return series
    non_null = series.dropna().astype(str).str.strip()
    dayfirst = True
    if not non_null.empty:
        sample = non_null.iloc[0]
        if len(sample) >= 5 and sample[:4].isdigit() and sample[4] in "-/":
            dayfirst = False
    return pd.to_datetime(series, errors="coerce", dayfirst=dayfirst)


def filter_by_month(raw_path: Path, year: int, month: int):
    """Return (xlsx_bytes, total_rows, kept_rows, invalid_dates, date_col)."""
    df = pd.read_excel(raw_path, engine="openpyxl")
    total_rows = len(df)

    normalized = {str(c).strip().lower(): c for c in df.columns}
    date_col = None
    for candidate in (DATE_COLUMN_NAME.lower(), "date created", "data created"):
        if candidate in normalized:
            date_col = normalized[candidate]
            break
    if date_col is None:
        date_col = df.columns[-1]

    dates = parse_datetime_column(df[date_col])
    invalid = int(dates.isna().sum())

    mask = (dates.dt.year == year) & (dates.dt.month == month)
    filtered = df.loc[mask].copy()
    filtered[date_col] = dates.loc[mask]

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl", datetime_format="yyyy-mm-dd hh:mm:ss") as writer:
        filtered.to_excel(writer, index=False, sheet_name="Closed Alarm")
        ws = writer.sheets["Closed Alarm"]
        for idx, col in enumerate(filtered.columns, start=1):
            max_len = max([len(str(col))] + [len(str(v)) for v in filtered[col].head(200).tolist()])
            ws.column_dimensions[ws.cell(row=1, column=idx).column_letter].width = min(max_len + 2, 50)
        ws.freeze_panes = "A2"

    return buffer.getvalue(), total_rows, len(filtered), invalid, str(date_col)


# ---------------------------------------------------------------- pipeline
def run_pipeline(username, password, year, month, report):
    download_dir = Path(tempfile.mkdtemp(prefix="pmt_alarm_"))
    driver = None
    try:
        report("Starting browser...")
        driver = build_driver(download_dir)

        report("Logging in...")
        login(driver, username, password)

        report("Opening alarm page...")
        open_alarm_page(driver)
        scroll_to_alarm_section(driver)

        report("Opening CLOSED ALARM tab...")
        click_closed_alarm_tab(driver)

        before = snapshot_files(download_dir)
        report("Requesting full download (All Page)...")
        click_download_and_all_page(driver)

        report("Waiting for the file to finish downloading (can take a few minutes)...")
        filename = wait_for_download_complete(download_dir, before)
        raw_path = download_dir / filename
        raw_bytes = raw_path.read_bytes()

        report(f"Filtering rows for {MONTHS[month - 1]} {year}...")
        filtered_bytes, total, kept, invalid, date_col = filter_by_month(raw_path, year, month)

        return {
            "raw_bytes": raw_bytes,
            "raw_name": filename,
            "filtered_bytes": filtered_bytes,
            "filtered_name": f"PMT_Alarm_Closed_{year}-{month:02d}.xlsx",
            "total": total,
            "kept": kept,
            "invalid": invalid,
            "date_col": date_col,
            "label": f"{MONTHS[month - 1]} {year}",
        }
    except Exception as exc:
        shot = None
        if driver is not None:
            try:
                shot = driver.get_screenshot_as_png()
            except Exception:
                pass
        raise PipelineError(str(exc) or type(exc).__name__, shot) from exc
    finally:
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass
        shutil.rmtree(download_dir, ignore_errors=True)


# ---------------------------------------------------------------- UI
def default_period():
    try:
        from zoneinfo import ZoneInfo

        now = datetime.now(ZoneInfo("Asia/Jakarta"))
    except Exception:
        now = datetime.now()
    return now.year, now.month


st.set_page_config(page_title="PMT Alarm Downloader", page_icon="🚨", layout="centered")
st.title("🚨 PMT Alarm – Closed Alarm Downloader")
st.caption(
    "Enter your PMT Alarm account and choose a period. The app logs in, downloads all "
    "Closed Alarm data, and filters it to the month you pick."
)

cur_year, cur_month = default_period()
years = list(range(cur_year, 2019, -1))

with st.form("download_form"):
    st.subheader("1. Account")
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")

    st.subheader("2. Period (Year-Month)")
    c1, c2 = st.columns(2)
    year = c1.selectbox("Year", years, index=0)
    month_name = c2.selectbox("Month", MONTHS, index=cur_month - 1)

    submitted = st.form_submit_button("Download & Filter", type="primary", use_container_width=True)

if submitted:
    st.session_state.pop("result", None)
    if not username or not password:
        st.error("Please fill in both username and password.")
    else:
        month = MONTHS.index(month_name) + 1
        with st.status("Working...", expanded=True) as status:
            try:
                result = run_pipeline(
                    username, password, int(year), month, report=lambda msg: st.write(f"⏳ {msg}")
                )
                st.session_state["result"] = result
                status.update(label="Done!", state="complete", expanded=False)
            except PipelineError as err:
                status.update(label="Failed", state="error", expanded=True)
                st.error(f"{err}")
                if err.screenshot:
                    st.image(err.screenshot, caption="Browser screenshot at the moment of failure")
            except Exception as err:
                status.update(label="Failed", state="error", expanded=True)
                st.error(f"Unexpected error: {err}")

result = st.session_state.get("result")
if result:
    st.success(
        f"{result['kept']:,} rows in {result['label']} "
        f"(out of {result['total']:,} total; date column: '{result['date_col']}')."
    )
    if result["invalid"]:
        st.warning(f"{result['invalid']:,} rows had an empty/unreadable date and were excluded.")
    if result["kept"] == 0:
        st.info("No rows matched that month. Try another period, or download the full raw file below.")

    d1, d2 = st.columns(2)
    d1.download_button(
        f"⬇️ Filtered ({result['label']})",
        data=result["filtered_bytes"],
        file_name=result["filtered_name"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    d2.download_button(
        "⬇️ Raw full export",
        data=result["raw_bytes"],
        file_name=result["raw_name"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.divider()
st.caption("Your credentials are used only for this run and are not stored or logged.")
