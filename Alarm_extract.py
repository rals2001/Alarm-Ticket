# =====================================================================================
#  PATCH  –  new per-page cycle
#     500 rows -> Start Date via REAL calendar -> (End Date untouched) -> Apply
#     -> verify 500 again -> go to page N -> verify 500 again -> Download By Page
#
#  HOW TO APPLY
#   A) DELETE:   JS_SET_VALUE, pw_set_date(), the old ensure_dates(), the old apply_filters(),
#                the old quick_set_size(), the old goto_page(), the old fetch_page()
#   B) CHANGE CONFIG:  remove REAPPLY_EVERY_PAGE, add the constants in section 0
#   C) PASTE everything below (sections 0-7) in their place (after pw_click_apply / goto helpers
#      are fine: all helpers they use already exist in your file)
#   D) retrieve_worker(): initial state dict becomes  (see section 8)
# =====================================================================================

# ---------------------------------------------------------------- 0) config -------
TARGET_PAGE_SIZE = 500        # the ONLY page size used (never derived from the number of rows on screen)
SIZE_VERIFY_WAIT = 10         # seconds to wait for the table to reach the wanted size before re-selecting it
SIZE_RESTORE_ATTEMPTS = 4     # how many times a reset page size is restored before the cycle fails
STABLE_SECONDS = 2.0          # table signature must stay unchanged this long after Apply
CAL_MAX_MONTH_CLICKS = 60     # safety cap when navigating months in the calendar


# ---------------------------------------------------------------- 1) page-size checks
def page_size_ok(page, size: int) -> bool:
    """
    True only when the table really shows the wanted size:
      * footer range available  -> rows on screen == min(size, rows remaining from the first row of this page)
      * no footer range         -> more than 10 rows are visible
    Never returns True while only 10 rows are visible and more records exist.
    """
    n = effective_rows(page)
    if n <= 0:
        return False
    r = read_range(page)
    if r:
        expected = min(size, r[2] - r[0] + 1)
        return n >= expected
    return n > DEFAULT_PAGE_ROWS


def log_state(page, job_dir: Path, size: int, tag: str):
    """One line with everything that matters: size, rows, footer range, current page."""
    rows = effective_rows(page)
    rng = read_range(page)
    job_log(job_dir, f"STATE [{tag}] page_size_wanted={size} visible_rows={rows} footer_range={rng} "
                     f"current_page={current_page(page, size)} size_ok={page_size_ok(page, size)}")


def ensure_page_size(page, job_dir: Path, state: dict, rep, stage: str, prog: float = 0.0) -> int:
    """
    Make sure the table shows TARGET_PAGE_SIZE rows. Waits for a slow reload first; if the size was reset
    (10 rows) it is selected again, up to SIZE_RESTORE_ATTEMPTS times. Raises if it cannot be restored:
    the cycle must never continue with 10 rows.
    """
    size = state["page_size"]
    for attempt in range(1, SIZE_RESTORE_ATTEMPTS + 1):
        end = time.time() + SIZE_VERIFY_WAIT
        while time.time() < end and not page_size_ok(page, size):
            page.wait_for_timeout(500)
        if page_size_ok(page, size):
            log_state(page, job_dir, size, f"{stage} -> OK")
            return effective_rows(page)

        log_state(page, job_dir, size, f"{stage} -> NOT {size}, restoring (attempt {attempt})")
        if not state.get("size_discovered"):
            # very first time in this job: full discovery of the Items control (screenshots + DOM dump)
            pw_set_page_size(page, job_dir, rep=rep, strict=True, tag="2a_", base_progress=prog,
                             sizes=(str(size),))
            state["size_discovered"] = True
        else:
            quick_set_size(page, job_dir, size, rep=rep, base_progress=prog)
        wait_loading_done(page, 15000)
        page.wait_for_timeout(500)

    log_state(page, job_dir, size, f"{stage} -> FAILED")
    raise RuntimeError(f"[{stage}] the page size could not be set to {size} "
                       f"(table shows {effective_rows(page)} rows). Not continuing with a small page.")


# quick_set_size: ONLY the early-return line changes (everything else stays as it is)
#   OLD:   n0 = effective_rows(page)
#          if n0 > DEFAULT_PAGE_ROWS:
#              return n0
#   NEW:   n0 = effective_rows(page)
#          if page_size_ok(page, size):
#              return n0
# (the old check "more than 10 rows" accepted a half-loaded table of 50 or 100 rows as "done")


# ---------------------------------------------------------------- 2) calendar widget
_CAL_COMMON = r"""
const CAL_CSS = ".v-date-picker, .v-picker, .v-date-picker-month, .flatpickr-calendar, .MuiDateCalendar-root, " +
  ".MuiPickersPopper-root, .MuiPickerStaticWrapper-root, .ant-picker-dropdown, .ant-picker-panel, " +
  ".datepicker, .datepicker-dropdown, .react-datepicker, .react-datepicker-popper, .q-date, .mat-calendar, " +
  ".mat-datepicker-content, .dp__menu, .vc-container, .daterangepicker, [class*='datepicker' i], " +
  "[class*='date-picker' i], [class*='calendar' i], [role='dialog']";
const MON = {jan:0,feb:1,mar:2,apr:3,mei:4,may:4,jun:5,jul:6,agu:7,agt:7,ags:7,aug:7,sep:8,okt:9,oct:9,nov:10,des:11,dec:11};
const MON_RE = /\b(jan|feb|mar|apr|mei|may|jun|jul|agu|agt|ags|aug|sep|okt|oct|nov|des|dec)[a-z]*\.?,?\s*(20\d\d)\b/i;
const vis = el => (el.offsetParent !== null || el.getClientRects().length > 0) && el.getBoundingClientRect().width > 0;
const dayCount = e => Array.from(e.querySelectorAll('button, td, [role="gridcell"], span, div'))
    .filter(x => x.children.length <= 1 && /^\d{1,2}$/.test((x.innerText || '').trim())).length;
const findCal = () => {
  const c = Array.from(document.querySelectorAll(CAL_CSS)).filter(e => {
    if (!vis(e)) return false;
    const r = e.getBoundingClientRect();
    return r.width > 150 && r.width < 700 && r.height > 150 && r.height < 700 && dayCount(e) >= 28;
  });
  c.sort((a, b) => b.getBoundingClientRect().width * b.getBoundingClientRect().height -
                   a.getBoundingClientRect().width * a.getBoundingClientRect().height);
  return c[0] || null;      // the largest matching container = includes the header and the prev/next buttons
};
const clear = () => document.querySelectorAll('[data-pw-cal]').forEach(e => e.removeAttribute('data-pw-cal'));
"""

CAL_STATE_JS = "() => {" + _CAL_COMMON + r"""
  const c = findCal();
  if (!c) return null;
  let month = -1, year = 0;
  // flatpickr style: <select> with months + <input> with the year
  const sel = Array.from(c.querySelectorAll('select')).find(s =>
      Array.from(s.options).some(o => MON[(o.text || '').trim().slice(0, 3).toLowerCase()] !== undefined));
  const yin = Array.from(c.querySelectorAll('input[type="number"], input.cur-year')).find(i => /^20\d\d$/.test(i.value));
  if (sel && yin) {
    month = MON[(sel.options[sel.selectedIndex].text || '').trim().slice(0, 3).toLowerCase()];
    year = parseInt(yin.value, 10);
  } else {
    const m = (c.innerText || '').match(MON_RE);
    if (m) { month = MON[m[1].toLowerCase().slice(0, 3)]; year = parseInt(m[2], 10); }
  }
  return {month: month === undefined ? -1 : month, year: year, head: (c.innerText || '').slice(0, 80)};
}"""

CAL_MARK_NAV_JS = "(dir) => {" + _CAL_COMMON + r"""
  clear();
  const c = findCal();
  if (!c) return false;
  const meta = e => ((e.getAttribute('aria-label') || '') + ' ' + (e.getAttribute('title') || '') + ' ' +
                     ((e.className && e.className.toString) ? e.className.toString() : '') + ' ' +
                     (e.innerText || '')).toLowerCase();
  const re = dir === 'prev' ? /prev|chevron[-_]?left|arrow[-_]?left|navigate_before|keyboard_arrow_left|‹|«|^\s*<\s*$/
                            : /next|chevron[-_]?right|arrow[-_]?right|navigate_next|keyboard_arrow_right|›|»|^\s*>\s*$/;
  const bad = e => e.disabled || e.getAttribute('aria-disabled') === 'true' || /disabled/.test(meta(e));
  const cands = Array.from(c.querySelectorAll('button, [role="button"], span, i, a, th, div'))
      .filter(e => vis(e) && re.test(meta(e)) && !/^\d{1,2}$/.test((e.innerText || '').trim()) && !bad(e));
  const leaf = cands.filter(e => !cands.some(o => o !== e && e.contains(o)));   // deepest only
  const btn = leaf.find(e => e.closest('button, [role="button"]')) || leaf[0];
  if (!btn) return false;
  (btn.closest('button, [role="button"]') || btn).setAttribute('data-pw-cal', 'nav');
  return true;
}"""

CAL_MARK_DAY_JS = "(day) => {" + _CAL_COMMON + r"""
  clear();
  const c = findCal();
  if (!c) return 'no-calendar';
  const OUT = /(^|[-_ ])(adjacent|outside|outside-month|othermonth|other-month|prevmonthday|nextmonthday|prevmonth|nextmonth|old|new|disabled|muted)([-_ ]|$)/i;
  const cls = e => ((e.className && e.className.toString) ? e.className.toString() : '');
  const excluded = e => {
    let n = e;
    for (let i = 0; i < 4 && n && n !== c; i++, n = n.parentElement) {
      if (OUT.test(cls(n)) || n.disabled || n.getAttribute('aria-disabled') === 'true') return true;
      if (/ant-picker-cell/.test(cls(n)) && !/in-view/.test(cls(n))) return true;
    }
    return false;
  };
  const all = Array.from(c.querySelectorAll('button, td, [role="gridcell"], span, div, a'))
      .filter(e => vis(e) && (e.innerText || '').trim() === String(day) && !excluded(e));
  const leaf = all.filter(e => !all.some(o => o !== e && e.contains(o)));
  const pick = leaf.find(e => e.closest('button')) || leaf[0];
  if (!pick) return 'no-day-cell (' + all.length + ' candidates)';
  pick.setAttribute('data-pw-cal', 'day');
  return 'ok';
}"""

CAL_MARK_OK_JS = "() => {" + _CAL_COMMON + r"""
  clear();
  const c = findCal();
  if (!c) return false;
  const b = Array.from(c.querySelectorAll('button, [role="button"]')).filter(e => vis(e) &&
      /^(ok|set|select|done|choose|pilih|simpan|confirm)$/i.test((e.innerText || '').trim()));
  if (!b.length) return false;
  b[b.length - 1].setAttribute('data-pw-cal', 'ok');
  return true;
}"""

CAL_HTML_JS = "() => {" + _CAL_COMMON + r"""
  const c = findCal();
  return c ? c.outerHTML.slice(0, 3500) : '(no calendar container found)';
}"""

CAL_OPEN_JS = "() => {" + _CAL_COMMON + "  return !!findCal();\n}"


def _click_real(page, name: str):
    """Genuine mouse click on the element marked data-pw-cal=<name> (no JS click, no value injection)."""
    loc = page.locator(f"[data-pw-cal='{name}']").first
    try:
        loc.scroll_into_view_if_needed(timeout=3000)
        loc.click(timeout=6000)
    except PWError:
        box = loc.bounding_box()
        if not box:
            raise RuntimeError(f"calendar element '{name}' has no bounding box")
        page.mouse.click(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)


def _calendar_open(page) -> bool:
    try:
        return bool(page.evaluate(CAL_OPEN_JS))
    except PWError:
        return False


def _dump_calendar(page, job_dir: Path, why: str):
    try:
        html = page.evaluate(CAL_HTML_JS)
    except PWError:
        html = "(could not read)"
    job_log(job_dir, f"CALENDAR DUMP ({why}): {html[:1500]}", "ERROR")
    append_dom(job_dir, "cal_", f"CALENDAR DUMP ({why}):\n{html}")
    take_shot(page, job_dir, "shot_calendar_problem")


def open_calendar(page, job_dir: Path):
    """Click the Start Date field like a user; if no calendar opens, click the calendar icon next to it."""
    inp = mark_control(page, "start", 15)
    inp.scroll_into_view_if_needed()
    inp.click()                                           # real click on the field
    for _ in range(10):
        if _calendar_open(page):
            return
        page.wait_for_timeout(300)
    # fallback: the icon/button inside the same field wrapper
    for xp in ("xpath=ancestor::*[self::div or self::label][1]//*[self::button or self::i or self::svg or "
               "contains(@class,'icon')]",
               "xpath=ancestor::*[self::div or self::label][2]//*[self::button or self::i or self::svg or "
               "contains(@class,'icon')]"):
        try:
            icon = mark_control(page, "start", 5).locator(xp).first
            if icon.count() and icon.is_visible():
                icon.click(timeout=4000)
                for _ in range(10):
                    if _calendar_open(page):
                        job_log(job_dir, "[Start Date] calendar opened via the calendar icon.")
                        return
                    page.wait_for_timeout(300)
        except PWError:
            continue
    _dump_calendar(page, job_dir, "calendar did not open")
    raise RuntimeError("The Start Date calendar did not open (see CALENDAR DUMP in job.log and "
                       "shot_calendar_problem.jpg).")


def pw_pick_start_date(page, job_dir: Path, value: date):
    """
    Select the Start Date through the REAL calendar: open it, navigate months, click the day, (confirm),
    verify the field. No fill / typing / keyboard entry / JS value injection.
    """
    text = value.strftime(DATE_INPUT_FORMAT)
    ok_values = {re.sub(r"\D", "", text), value.strftime("%Y%m%d")}
    field = mark_control(page, "start", 15)
    job_log(job_dir, f"[Start Date] field shows '{field.input_value()}' before; selecting {text} in the calendar")

    open_calendar(page, job_dir)

    # ---- navigate to the wanted month/year
    target = value.year * 12 + (value.month - 1)
    for step in range(CAL_MAX_MONTH_CLICKS + 1):
        st_ = page.evaluate(CAL_STATE_JS)
        if not st_ or st_["month"] < 0 or not st_["year"]:
            _dump_calendar(page, job_dir, "cannot read month/year")
            raise RuntimeError(f"Cannot read the month/year shown by the calendar (header: "
                               f"{st_ and st_.get('head')!r}).")
        cur = st_["year"] * 12 + st_["month"]
        if cur == target:
            job_log(job_dir, f"[Start Date] calendar shows month {st_['month'] + 1}/{st_['year']} after "
                             f"{step} click(s).")
            break
        if step == CAL_MAX_MONTH_CLICKS:
            raise RuntimeError("Calendar navigation did not reach the wanted month.")
        direction = "prev" if cur > target else "next"
        if not page.evaluate(CAL_MARK_NAV_JS, direction):
            _dump_calendar(page, job_dir, f"no '{direction}' button")
            raise RuntimeError(f"Calendar has no usable '{direction}' month button.")
        _click_real(page, "nav")
        page.wait_for_timeout(250)
        st2 = page.evaluate(CAL_STATE_JS)
        if st2 and (st2["month"], st2["year"]) == (st_["month"], st_["year"]):
            page.wait_for_timeout(500)                      # animation may still be running

    # ---- click the day
    res = page.evaluate(CAL_MARK_DAY_JS, value.day)
    if res != "ok":
        _dump_calendar(page, job_dir, f"day {value.day}: {res}")
        raise RuntimeError(f"Could not find day {value.day} in the calendar ({res}).")
    _click_real(page, "day")
    page.wait_for_timeout(500)

    # ---- some pickers need an OK / Set button
    if _calendar_open(page) and page.evaluate(CAL_MARK_OK_JS):
        _click_real(page, "ok")
        page.wait_for_timeout(500)

    # ---- verify what the field really holds
    got = ""
    for _ in range(10):
        got = mark_control(page, "start", 10).input_value()
        if re.sub(r"\D", "", got) in ok_values:
            break
        page.wait_for_timeout(300)
    else:
        _dump_calendar(page, job_dir, "wrong value after click")
        raise RuntimeError(f"After clicking {text} in the calendar the field shows '{got}'.")

    if _calendar_open(page):                                 # popup still open: close it (value is kept)
        page.keyboard.press("Escape")
        page.wait_for_timeout(400)
        got = mark_control(page, "start", 10).input_value()
        if re.sub(r"\D", "", got) not in ok_values:
            raise RuntimeError(f"Closing the calendar changed the Start Date to '{got}'.")
    job_log(job_dir, f"[Start Date] SELECTED via calendar: '{got}' (wanted {text})")


def read_end_date(page) -> str:
    """READ-ONLY look at the End Date box (no click, no typing) for the log."""
    try:
        return mark_control(page, "end", 5).input_value()
    except (PWError, RuntimeError):
        return "?"


def ensure_dates(page, job_dir: Path, start_d: date, end_d: date):
    """
    Start Date: always selected through the calendar widget.
    End Date  : NEVER touched. The site keeps its own default (today). `end_d` is accepted only so that the
                callers do not change; it is not used here.
    """
    pw_pick_start_date(page, job_dir, start_d)
    job_log(job_dir, f"[End Date] left untouched, site value: '{read_end_date(page)}'")


# ---------------------------------------------------------------- 3) Apply ---------
def wait_table_stable(page, timeout=30, still=STABLE_SECONDS) -> bool:
    """True once the table signature stayed identical for `still` seconds."""
    end = time.time() + timeout
    last, since = None, time.time()
    while time.time() < end:
        s = get_sig(page)
        if s != last:
            last, since = s, time.time()
        elif time.time() - since >= still:
            return True
        page.wait_for_timeout(500)
    return False


def apply_filters(page, job_dir: Path, expect_change: bool, size: int = TARGET_PAGE_SIZE):
    """Click Apply, wait for the reload, the loading indicators and a stable table. Logs size before/after."""
    log_state(page, job_dir, size, "BEFORE Apply")
    old = get_sig(page)
    pw_click_apply(page)
    changed = wait_sig_change(page, old, APPLY_LOAD_TIMEOUT if expect_change else APPLY_QUICK_TIMEOUT)
    wait_loading_done(page, 30000)
    stable = wait_table_stable(page)
    page.wait_for_timeout(500)
    job_log(job_dir, f"Apply done: table_changed={changed} stable={stable} rows={effective_rows(page)}")
    log_state(page, job_dir, size, "AFTER Apply (before size check)")


# ---------------------------------------------------------------- 4) quick_set_size: see note in section 1

# ---------------------------------------------------------------- 5) navigation ----
def goto_page(page, job_dir: Path, target: int, state: dict, rep, monitor, browser):
    """Go to page `target`, always re-verifying that the page size is still TARGET_PAGE_SIZE."""
    size = state["page_size"]
    resets = 0
    for step in range(1, MAX_NAV_STEPS + 1):
        cur = current_page(page, size)

        if cur == target:
            if page_size_ok(page, size):
                log_state(page, job_dir, size, f"NAV page {target} reached after {step - 1} click(s)")
                return
            resets += 1
            job_log(job_dir, f"NAV: on page {target} but the size is NOT {size} (reset #{resets}); restoring.",
                    "WARNING")
            if resets > 4:
                raise RuntimeError(f"The page size keeps resetting while navigating to page {target}.")
            ensure_page_size(page, job_dir, state, rep, f"nav-restore p{target}")
            continue                                       # a restore may jump back to page 1 -> loop re-navigates

        try:
            res = page.evaluate(MARK_PAGE_JS, target)
        except PWError as exc:
            raise RuntimeError(f"Could not read the page buttons: {str(exc)[:120]}")
        mode = res.get("mode")
        old = get_sig(page)
        if mode in ("exact", "jump"):
            safe_click(page.locator("[data-pw-page]").first)
        else:
            outcome = click_next(page)
            if outcome != "clicked":
                raise RuntimeError(f"Cannot reach page {target}: no page button visible and 'Next' is "
                                   f"'{outcome}' (current page {cur}, buttons {res.get('digits')}).")
        changed = wait_sig_change(page, old, 40)
        wait_loading_done(page, 15000)
        page.wait_for_timeout(500)
        job_log(job_dir, f"NAV step {step}: target={target} was_on={cur} click={mode if mode != 'none' else 'next'} "
                         f"table_changed={changed} rows_now={effective_rows(page)} buttons={res.get('digits')} "
                         f"footer_range={read_range(page)}")
        if monitor.tripped:
            raise RuntimeError(monitor.tripped)
        if not browser.is_connected() or page.is_closed():
            raise RuntimeError("Chromium exited while navigating.")

        if size_was_reset(page, size, target, state):
            resets += 1
            job_log(job_dir, f"NAV: page size RESET to {effective_rows(page)} rows after the click "
                             f"(reset #{resets}); selecting {size} again.", "WARNING")
            if resets > 4:
                raise RuntimeError(f"The page size keeps resetting to 10 while navigating to page {target}.")
            ensure_page_size(page, job_dir, state, rep, f"nav-reset p{target}")
            continue

        if mode == "exact" and changed and current_page(page, size) == 0:
            job_log(job_dir, f"NAV: clicked page button {target}; current page unreadable, assuming success.")
            return
    raise RuntimeError(f"Could not reach page {target} within {MAX_NAV_STEPS} clicks.")


# ---------------------------------------------------------------- 6) the page cycle --
def fetch_page(page, job_dir: Path, dl_dir: Path, pages_dir: Path, rep: StepReporter, monitor, browser,
               state: dict, n: int, start_d: date, end_d: date, first_cycle: bool) -> bool:
    """
    ONE cycle:
      1 size 500 -> 2 Start Date via calendar -> 3 End Date untouched -> 4 Apply -> 5 verify 500
      -> 6 go to page n, verify 500 -> 7 Download Alarm > By Page.
    Returns True when page n was the LAST page.
    """
    size = state["page_size"]
    total = max(state["total_pages"], n)
    prog = 0.46 + 0.44 * min((n - 1) / max(total, 1), 1.0)
    rep.running(f"Page {n} of ~{total}: 500 rows, start date, Apply, go to page, download...", prog, quiet=True)
    job_log(job_dir, f"P{n} CYCLE start | {format_snapshot(collect_snapshot(dl_dir))}")

    # ---- STEP 1: items per page FIRST
    ensure_page_size(page, job_dir, state, rep, f"P{n} step1 before dates", prog)

    # ---- STEP 2 + 3: Start Date through the calendar, End Date untouched
    ensure_dates(page, job_dir, start_d, end_d)

    # ---- STEP 4: Apply
    apply_filters(page, job_dir, expect_change=first_cycle, size=size)

    # ---- STEP 5: the site may reset the size after Apply
    ensure_page_size(page, job_dir, state, rep, f"P{n} step5 after Apply", prog)

    # ---- STEP 6: navigate (after Apply the table is on page 1)
    if n > 1:
        goto_page(page, job_dir, n, state, rep, monitor, browser)
    ensure_page_size(page, job_dir, state, rep, f"P{n} step6 after navigation", prog)

    # ---- final gate before the download
    rows = effective_rows(page)
    rng = read_range(page)
    cur = current_page(page, size)
    job_log(job_dir, f"P{n} READY visible_rows={rows} footer_range={rng} current_page={cur} page_size={size}")
    if rng and rng[0] != (n - 1) * size + 1:
        raise RuntimeError(f"Wrong page on screen: footer range starts at row {rng[0]} but page {n} must start "
                           f"at row {(n - 1) * size + 1}.")
    if cur not in (0, n):
        raise RuntimeError(f"Wrong page on screen: current page is {cur}, expected {n}.")
    if not page_size_ok(page, size):
        raise RuntimeError(f"Refusing to download page {n}: page size is not {size} ({rows} rows visible).")

    # ---- STEP 7: download
    job_log(job_dir, f"P{n} BEFORE download | {format_snapshot(collect_snapshot(dl_dir))}")
    dest = pw_download_by_page(page, job_dir, pages_dir / f"page_{n:04d}")      # raises if missing / empty
    job_log(job_dir, f"P{n} DOWNLOAD OK: {dest.name} {dest.stat().st_size / 1024:.0f} KB "
                     f"(exists={dest.exists()}, non-empty={dest.stat().st_size > 0})")
    job_log(job_dir, f"P{n} AFTER download  | {format_snapshot(collect_snapshot(dl_dir))}")

    digest = file_sha256(dest)
    if digest in state["hashes"]:
        dest.unlink(missing_ok=True)
        raise RuntimeError(f"The file downloaded for page {n} is identical to page {state['hashes'][digest]} "
                           f"(the table was not on page {n}).")
    state["hashes"][digest] = n
    state["files"].append(dest)
    state["resume_page"] = n + 1

    if n <= PROBE_FIRST_PAGES or n % PROBE_EVERY == 0:
        n_file = probe_xlsx_rows(dest)
        n_scr = count_rows(page)
        job_log(job_dir, f"P{n} VERIFY file_rows={n_file} screen_rows={n_scr}",
                "WARNING" if n_file not in (-1, n_scr) else "INFO")

    # ---- last page?  (unchanged logic)
    last = False
    rng = read_range(page)
    try:
        if rng and size:
            state["total_pages"] = max(state["total_pages"], n, math.ceil(rng[2] / size))
        else:
            state["total_pages"] = max(state["total_pages"], n, int(page.evaluate(TOTAL_PAGES_JS) or 0))
    except PWError:
        pass
    if rng and rng[1] >= rng[2]:
        last = True
        job_log(job_dir, f"Page {n}: footer range {rng[0]}-{rng[1]} of {rng[2]} -> last page.")
    else:
        nxt, disabled = find_next(page)
        if nxt is None:
            job_log(job_dir, "No 'next page' button found: assuming this is the only/last page.", "WARNING")
            last = True
        else:
            _dispose(nxt)
            if disabled:
                job_log(job_dir, f"'Next' is disabled on page {n}: last page.")
                last = True
    return last


# ---------------------------------------------------------------- 7) nothing else changes
#   download_loop(), run_session(), pw_download_by_page(), compile_files() keep working as they are.


# ---------------------------------------------------------------- 8) retrieve_worker state
#   OLD: state = {"files": [], "hashes": {}, "resume_page": 1, "total_pages": 1, "session_no": 0, "page_size": None}
#   NEW:
#   state = {"files": [], "hashes": {}, "resume_page": 1, "total_pages": 1, "session_no": 0,
#            "page_size": TARGET_PAGE_SIZE, "size_discovered": False}
#
#   In run_session(), after a browser RESTART the new browser instance must run the full discovery again only if
#   you want the screenshots; otherwise set state["size_discovered"] = True after session 1 (it is already
#   True after the first successful ensure_page_size, so nothing to do).
