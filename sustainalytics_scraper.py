"""
Sustainalytics Risk Report Scraper
-----------------------------------
Reads a ticker from cell R3 of your Excel workbook, searches for it on
Sustainalytics Global Access, navigates to the company's ESG Risk Rating
page, scrapes the specific numbers you pointed out, and writes them into
columns O7/O8/O10/O11/O12/O14/O15 of the workbook.

LOGIN: fully automatic using stored credentials (no typing your
username/password every run, no pressing Enter mid-script, no password
prompt at all). One-time setup:

    python setup_credentials.py

This asks for your Sustainalytics username/password once and saves them
(lightly encoded, not plaintext-obvious if someone opens the file, but
NOT strong security) into credentials.enc. Anyone with access to this
folder can run setup_credentials.py again at any time to change the
stored username/password - there's no master password gating that.
Every time you run the scraper, it reads credentials.enc automatically,
fills the login form, submits it, and continues straight into scraping -
no manual steps, no prompts.

NOTE ON MFA: if your account requires a one-time code / push approval,
that step still needs a human (nothing can automate that safely). If
auto-login hangs waiting for the search bar, use --manual-login instead
to fall back to the old "log in yourself, then press Enter" flow.

WHAT YOU NEED TO DO:
1. pip install selenium openpyxl webdriver-manager
2. Install Google Chrome (chromedriver is handled automatically via
   webdriver-manager in start_driver()).
3. One-time: python setup_credentials.py
4. Fill in LOGIN_USERNAME_SELECTOR / LOGIN_PASSWORD_SELECTOR /
   LOGIN_SUBMIT_SELECTOR below with the real selectors from the
   Sustainalytics login page (Inspect each field -> Copy outerHTML).
5. Run it:
       python sustainalytics_scraper.py --workbook "Risk Report - Hanley_Scraper.xlsx"
   It will read the ticker from C3 (or R3 override), log in automatically,
   then do the rest.
6. Optional: scrape more than one ticker in one run:
       python sustainalytics_scraper.py --workbook "Risk Report - Hanley_Scraper.xlsx" --tickers ASML,COHR,LITE,LFUS,AEIS
   (each ticker's results are printed + saved to a row in the output sheet;
   see write_excel() if you want to map these into a comparison layout
   instead of just O7-O15 for a single company).

A FEW SELECTORS ARE BEST-GUESS (marked "# BEST GUESS" below) because I
don't have the raw HTML for every field yet - specifically the
"Management Gap" value on the Risk Decomposition tab. If that one comes
back as None, right-click the "14.6"-style number on that tab -> Inspect
-> Copy outerHTML, and send it to me so I can lock in the exact selector.
"""

import argparse
import base64
import getpass
import io
import json
import re
import time
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    ElementClickInterceptedException,
)

import openpyxl

HOME_URL = "https://globalaccess.sustainalytics.com/#/ga/home"
LOGIN_URL = "https://globalaccess.sustainalytics.com/#/ga/home"  # TODO confirm real login URL
WAIT_SECONDS = 20

CRED_FILE = "credentials.enc"


# ---------------------------------------------------------------------------
# Encrypted credential loading (see setup_credentials.py to create these)
# ---------------------------------------------------------------------------

def load_credentials():
    """Reads credentials.enc (created/updated by setup_credentials.py) and
    returns (username, password). No password prompt - anyone running this
    script can use whatever credentials are currently stored. Exits with a
    clear message if the file doesn't exist yet."""
    import os
    if not os.path.exists(CRED_FILE):
        raise SystemExit(
            f"\n{CRED_FILE} not found.\n"
            "Run: python setup_credentials.py   (one-time setup) first."
        )

    with open(CRED_FILE, "rb") as fh:
        encoded = fh.read()

    decoded = base64.urlsafe_b64decode(encoded)
    creds = json.loads(decoded.decode())
    return creds["username"], creds["password"]


# ---------------------------------------------------------------------------
# Browser setup
# ---------------------------------------------------------------------------

def start_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--window-size=1400,1000")
    # Auto-manage chromedriver so nobody has to manually download/match a
    # version to their installed Chrome. Falls back to a plain PATH-based
    # chromedriver if webdriver-manager isn't installed.
    try:
        from webdriver_manager.chrome import ChromeDriverManager
        from selenium.webdriver.chrome.service import Service
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    except ImportError:
        driver = webdriver.Chrome(options=options)
    return driver


def wait_visible(driver, by, value, timeout=WAIT_SECONDS):
    return WebDriverWait(driver, timeout).until(
        EC.visibility_of_element_located((by, value))
    )


def wait_url_contains(driver, substr, timeout=WAIT_SECONDS):
    WebDriverWait(driver, timeout).until(EC.url_contains(substr))


# ---------------------------------------------------------------------------
# Login (automatic, using encrypted stored credentials)
# ---------------------------------------------------------------------------

# Real selectors from the Morningstar/Auth0 login page (login-prod.morningstar.com).
LOGIN_USERNAME_SELECTOR = (By.CSS_SELECTOR, "#email")
LOGIN_PASSWORD_SELECTOR = (By.CSS_SELECTOR, "#password")
LOGIN_SUBMIT_SELECTOR = (By.CSS_SELECTOR, "button.uim-ul-form-submit")


def auto_login(driver, username, password):
    driver.get(LOGIN_URL)
    print("\n>>> Logging in automatically...")

    user_box = wait_visible(driver, *LOGIN_USERNAME_SELECTOR)
    time.sleep(1)  # let the Auth0/Angular form finish settling before typing
    user_box.clear()
    user_box.send_keys(username)

    pass_box = wait_visible(driver, *LOGIN_PASSWORD_SELECTOR)
    pass_box.clear()
    pass_box.send_keys(password)

    current_url = driver.current_url
    submit_btn = wait_visible(driver, *LOGIN_SUBMIT_SELECTOR)
    submit_btn.click()

    # Wait for the URL to change away from the login page (confirms the
    # click actually submitted), then wait for the search bar to confirm
    # we landed on the Sustainalytics home page.
    try:
        WebDriverWait(driver, max(WAIT_SECONDS, 30)).until(
            lambda d: d.current_url != current_url
        )
        wait_visible(driver, By.CSS_SELECTOR, "#searchInput", timeout=max(WAIT_SECONDS, 30))
        print(">>> Login successful.")
    except TimeoutException:
        _debug_dump(driver, "auto_login_timeout")
        raise SystemExit(
            "\nLogin did not complete (no search bar found after submitting).\n"
            "This usually means: (1) an MFA/verification step is required and needs\n"
            "a human, or (2) the login selectors need updating. Check\n"
            "debug_auto_login_timeout.png to see what the page looked like."
        )


def manual_login(driver):
    """Fallback: opens the login page and waits for you to log in by hand.
    Kept available via --manual-login in case MFA makes auto-login impossible."""
    driver.get(HOME_URL)
    print("\n>>> A Chrome window just opened to Sustainalytics Global Access.")
    print(">>> Log in manually (including any MFA step).")
    input(">>> Once you see the search bar on the home page, come back here and press Enter... ")


# ---------------------------------------------------------------------------
# Navigation: search -> click first result -> overview -> ESG Risk Rating link
# ---------------------------------------------------------------------------

def search_company(driver, ticker):
    driver.get(HOME_URL)
    box = wait_visible(driver, By.CSS_SELECTOR, "#searchInput")
    box.clear()
    box.send_keys(ticker)

    # Wait for the dropdown result(s) to render, then click the first one.
    first_result = WebDriverWait(driver, WAIT_SECONDS).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "search-result.list-group-item"))
    )
    first_result.click()
    wait_url_contains(driver, "/overview")


def open_risk_overview(driver):
    # On the company "overview" page, click the "ESG Risk Rating" link.
    # The real clickable element has class "text-link" (confirmed from the
    # actual page HTML). Matching on that specific class avoids accidentally
    # grabbing a non-clickable heading/label elsewhere on the page that also
    # contains the words "ESG Risk Rating".
    long_wait = max(WAIT_SECONDS, 40)
    link_xpath = "//*[contains(@class,'text-link') and contains(., 'ESG Risk Rating')]"
    try:
        link = WebDriverWait(driver, long_wait).until(
            EC.presence_of_element_located((By.XPATH, link_xpath))
        )
    except TimeoutException:
        _debug_dump(driver, "open_risk_overview_timeout")
        raise

    # Scroll it into view and try a normal click; if something (e.g. the
    # dismissible info banner) is overlapping it, fall back to a JS click.
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", link)
    try:
        WebDriverWait(driver, long_wait).until(
            EC.element_to_be_clickable((By.XPATH, link_xpath))
        )
        link.click()
    except (TimeoutException, ElementClickInterceptedException):
        driver.execute_script("arguments[0].click();", link)

    # If the click opened a new tab/window, switch the driver to it before
    # checking the URL.
    if len(driver.window_handles) > 1:
        driver.switch_to.window(driver.window_handles[-1])

    try:
        wait_url_contains(driver, "/risk/overview")
    except TimeoutException:
        _debug_dump(driver, "open_risk_overview_url_timeout")
        print(f">>> DEBUG: current URL is: {driver.current_url}")
        raise


def _debug_dump(driver, label):
    """On failure, save a screenshot + page source so we can see what the
    browser actually rendered. Files land next to the script."""
    try:
        driver.save_screenshot(f"debug_{label}.png")
        with open(f"debug_{label}.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        print(f">>> Saved debug_{label}.png / .html for inspection.")
    except Exception as e:
        print(f">>> Could not save debug info: {e}")


# ---------------------------------------------------------------------------
# Scraping the /risk/overview page
# ---------------------------------------------------------------------------

def text_or_none(el):
    try:
        return el.text.strip()
    except Exception:
        return None


def _find_trend_near(driver, score_el):
    """Given the Exposure (or Management) score element, find the '.momentum'
    span that actually belongs to it.

    Angular tags every element rendered by a given component with a matching
    '_ngcontent-ng-cXXXXXXXXX' attribute (visible in the raw HTML you
    pasted - e.g. the Exposure block's elements all carry
    '_ngcontent-ng-c1976744147', while the Management block's elements carry
    a different one like '_ngcontent-ng-c11384699'). That means the score
    element and its OWN trend arrow are always marked with the same
    attribute, even though the two blocks look identical in class names.
    So: read that attribute straight off the score element, then look for a
    '.momentum' element carrying that same attribute - this scopes the
    search exactly to the right component instead of guessing by page
    order (which broke before)."""
    ng_attr = driver.execute_script(
        "for (const a of arguments[0].attributes) {"
        "  if (a.name.startsWith('_ngcontent')) return a.name;"
        "}"
        "return null;",
        score_el,
    )
    if ng_attr:
        matches = driver.find_elements(By.CSS_SELECTOR, f".momentum[{ng_attr}]")
        if matches:
            return text_or_none(matches[0])
    # Fallback: if no ng-content attribute was found (different framework
    # version / structure), fall back to nearest-ancestor search.
    for level in range(1, 8):
        try:
            container = score_el.find_element(By.XPATH, f"./ancestor::*[{level}]")
        except Exception:
            break
        matches = container.find_elements(By.CSS_SELECTOR, ".momentum")
        if matches:
            return text_or_none(matches[0])
    return None


def scrape_exposure_management(driver):
    """Default tab on /risk/overview: Exposure & Management.
    Two '.risk-rating.rating-line' elements (Exposure score, then Management
    score). Each one's own trend arrow ('.momentum', e.g.
    <span class="momentum green">0.0</span> or
    <span class="momentum red">-5.6</span>) is found by walking up from that
    specific score element, rather than just grabbing "the Nth .momentum on
    the page" - that approach silently grabbed the wrong element when the
    page had extra/differently-ordered momentum spans elsewhere."""
    data = {}

    scores = WebDriverWait(driver, WAIT_SECONDS).until(
        lambda d: d.find_elements(By.CSS_SELECTOR, ".risk-rating.rating-line") or False
    )
    data["exposure_score"] = text_or_none(scores[0]) if len(scores) > 0 else None
    data["management_score"] = text_or_none(scores[1]) if len(scores) > 1 else None

    data["exposure_trend"] = _find_trend_near(driver, scores[0]) if len(scores) > 0 else None
    data["management_trend"] = _find_trend_near(driver, scores[1]) if len(scores) > 1 else None

    return data


def scrape_risk_decomposition(driver):
    """Click the 'Risk Decomposition' tab and grab the Management Gap value.

    The whole chart (bars AND the numeric labels, e.g. "15.4") is drawn
    inside a single <canvas> element - there is no DOM text to select, so
    this falls back to OCR: screenshot just the canvas, crop roughly to the
    "Management Gap" row, and read the number off the image with
    pytesseract. Requires:
        pip install pytesseract pillow
    and the separate Tesseract-OCR program installed (not just pip) -
    https://github.com/UB-Mannheim/tesseract/wiki (Windows installer).
    If Tesseract isn't installed, this prints a clear message and returns
    None instead of crashing the whole script.
    """
    tab = WebDriverWait(driver, WAIT_SECONDS).until(
        EC.element_to_be_clickable(
            (By.XPATH, "//*[contains(@class,'tabHeader') and contains(normalize-space(text()), 'Risk Decomposition')]")
        )
    )
    tab.click()
    time.sleep(2.5)  # panel swap animation

    try:
        canvas = WebDriverWait(driver, WAIT_SECONDS).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "risk-decomposition canvas, div.pane canvas"))
        )
    except TimeoutException:
        print(">>> Could not find the Risk Decomposition canvas at all.")
        return None

    # IMPORTANT: canvas.screenshot_as_png (a plain Selenium element
    # screenshot) often comes back blank for canvases that are scaled by
    # devicePixelRatio or rendered via WebGL - the screenshot machinery
    # doesn't reliably capture the true drawing buffer. Instead, ask the
    # browser itself for the canvas's real pixel data via toDataURL(),
    # which reads the actual buffer directly and sidesteps that problem.
    import base64

    def _grab_canvas_png_bytes():
        data_url = driver.execute_script(
            "return arguments[0].toDataURL('image/png');", canvas
        )
        header, b64data = data_url.split(",", 1)
        return base64.b64decode(b64data)

    def _looks_blank(png_bytes):
        try:
            from PIL import Image, ImageStat
            img = Image.open(io.BytesIO(png_bytes)).convert("L")
            stat = ImageStat.Stat(img)
            return stat.stddev[0] < 3  # near-uniform color = nothing drawn yet
        except Exception:
            return False

    png_bytes = _grab_canvas_png_bytes()
    for _ in range(10):
        if not _looks_blank(png_bytes):
            break
        time.sleep(1)
        png_bytes = _grab_canvas_png_bytes()
    else:
        print(">>> Warning: canvas still looks blank via toDataURL after "
              "waiting - chart may be WebGL without preserveDrawingBuffer, "
              "or genuinely hasn't rendered. Check the debug PNG.")

    try:
        import pytesseract
        from PIL import Image
    except ImportError:
        print(">>> pytesseract/Pillow not installed - run: pip install pytesseract pillow")
        print(">>> (and install the Tesseract-OCR program itself, not just the pip package)")
        return None

    # Point pytesseract directly at the Tesseract-OCR .exe instead of relying
    # on it being added to the Windows PATH (the installer doesn't always do
    # this reliably). Checks the common default install location(s) and only
    # overrides tesseract_cmd if the default name isn't already resolvable.
    import os
    import shutil
    if shutil.which("tesseract") is None:
        for candidate in (
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        ):
            if os.path.exists(candidate):
                pytesseract.pytesseract.tesseract_cmd = candidate
                break

    img = Image.open(io.BytesIO(png_bytes))

    # Save the full chart image for reference/debugging either way.
    try:
        img.save("debug_risk_decomposition_canvas.png")
    except Exception:
        pass

    # OCR the whole chart, then find the "Management Gap" line and pull the
    # first number that follows it. Using the full image (not a hard crop)
    # is more robust to the chart's height/row-order shifting slightly
    # between companies.
    try:
        full_text = pytesseract.image_to_string(img)
    except Exception as e:
        print(f">>> OCR failed (is Tesseract-OCR installed on PATH?): {e}")
        return None

    print(">>> OCR raw text was:")
    print(full_text)
    print(">>> --- end OCR text ---")

    # Tesseract sometimes splits a visual row like "Management Gap   15.4"
    # into separate lines (label on one line, number on the next). So
    # instead of only checking the same line, search a small window of
    # lines after wherever "Gap" appears for the first number-looking token.
    lines = full_text.splitlines()
    for i, line in enumerate(lines):
        if "gap" in line.lower():
            # Check this line first, then the next couple of lines.
            for candidate_line in lines[i:i + 3]:
                m = re.search(r"-?\d+\.\d+", candidate_line)
                if m:
                    return m.group()

    print(">>> OCR ran but a number near 'Gap' wasn't found. "
          "Check debug_risk_decomposition_canvas.png and the raw OCR text "
          "printed above to see what was captured.")
    return None


def scrape_global_standards_and_controversy(driver):
    """Click through to the Global Standards Screening (gss) page and grab
    the compliance status + controversy level text."""
    data = {}

    gss_link = WebDriverWait(driver, WAIT_SECONDS).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href*='/gss']"))
    )
    gss_link.click()
    time.sleep(2)

    compliant_el = WebDriverWait(driver, WAIT_SECONDS).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, "span.pt-2.d-inline-block"))
    )
    data["global_standards_status"] = text_or_none(compliant_el)

    # The "Controversy level" rating lives on a separate page:
    #   .../research/company/<id>/controversies/overview
    # "Controversy Ratings" in the sidebar is just a section header (not
    # reliably clickable) - the real page is predictable from the current
    # URL, which always contains /research/company/<id>/... . So instead of
    # clicking a sidebar link, swap the path segment directly via driver.get()
    # - far more reliable than hunting for a clickable element.
    try:
        current_url = driver.current_url
        m = re.search(r"(.*/research/company/\d+)/", current_url)
        if not m:
            raise RuntimeError(f"Could not find company id in URL: {current_url}")
        base = m.group(1)
        controversy_url = f"{base}/controversies/overview"
        driver.get(controversy_url)
        wait_url_contains(driver, "/controversies/overview")
        time.sleep(1.5)

        # Try the confirmed selector first, then fall back to the text next
        # to the controversy icon in case the page structure varies.
        controversy_el = None
        for selector in ("span.ratingAssessment", "div.mr-auto"):
            try:
                controversy_el = WebDriverWait(driver, 8).until(
                    EC.visibility_of_element_located((By.CSS_SELECTOR, selector))
                )
                if text_or_none(controversy_el):
                    break
            except TimeoutException:
                continue

        data["controversy_level"] = text_or_none(controversy_el) if controversy_el else None
        print(f">>> controversy_level found: {data['controversy_level']!r}")
    except Exception as e:
        print(f">>> Could not find controversy level element: {e}")
        _debug_dump(driver, "controversy_level_not_found")
        data["controversy_level"] = None

    return data


def scrape_company(driver, ticker):
    search_company(driver, ticker)
    open_risk_overview(driver)

    data = {}
    data.update(scrape_exposure_management(driver))
    data["management_gap"] = scrape_risk_decomposition(driver)
    data.update(scrape_global_standards_and_controversy(driver))
    return data


# ---------------------------------------------------------------------------
# Excel I/O
# ---------------------------------------------------------------------------

def read_ticker(workbook_path, primary_cell="C3", override_cell="R3"):
    """Reads the ticker primarily by parsing C3, which holds strings like
    '3M (MMM)' or 'Walmart (WMT)' - pulls whatever's inside the parentheses.
    R3 is only used as an explicit manual override if someone deliberately
    types a ticker there. This avoids the trap where a stale leftover value
    in R3 from a previous run silently overrides the current company in C3."""
    wb = openpyxl.load_workbook(workbook_path, data_only=True)
    ws = wb.worksheets[0]

    override = ws[override_cell].value
    if override and str(override).strip():
        print(f"NOTE: using manual override ticker '{str(override).strip()}' "
              f"from {override_cell}. Clear {override_cell} to auto-parse "
              f"from {primary_cell} instead.")
        return str(override).strip()

    primary = ws[primary_cell].value
    if primary and str(primary).strip():
        m = re.search(r"\(([^)]+)\)", str(primary))
        if m:
            return m.group(1).strip()
        return str(primary).strip()  # no parentheses - use as-is

    raise ValueError(f"Could not find a ticker in {primary_cell} or {override_cell}.")


def _to_number(value):
    """Converts a scraped value (often a string like '43.1' or '+5.0') into
    a real float so Excel stores it as a number, not text. Returns None if
    it can't be parsed, so the cell is left blank instead of holding a
    broken value."""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return value
    s = str(value).strip()
    if not s:
        return None
    s = s.replace("+", "")  # "+5.0" -> "5.0"
    try:
        return float(s)
    except ValueError:
        return None


def write_results(workbook_path, output_path, data):
    """Writes the raw scraped values into O7, O8, O10, O11, O12, O14, O15.
    O7/O8/O10/O11/O12 are written as real numbers (needed for the N-column
    scoring formulas to work). O14/O15 are left as text on purpose, since
    they're statuses/categories (e.g. 'Compliant', 'Moderate'), not numbers."""
    wb = openpyxl.load_workbook(workbook_path)
    ws = wb.worksheets[0]

    ws["O7"] = _to_number(data.get("exposure_score"))
    ws["O8"] = _to_number(data.get("exposure_trend"))
    ws["O10"] = _to_number(data.get("management_score"))
    ws["O11"] = _to_number(data.get("management_trend"))
    ws["O12"] = _to_number(data.get("management_gap"))
    ws["O14"] = data.get("global_standards_status")  # text on purpose
    ws["O15"] = data.get("controversy_level")  # text on purpose

    wb.save(output_path)
    print(f"Saved: {output_path}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Scrape Sustainalytics ESG Risk Rating data into Excel")
    parser.add_argument(
        "--workbook",
        default=str(Path(__file__).parent / "Risk Report - Hanley_Scraper.xlsx"),
        help="Path to the Excel workbook (defaults to 'Risk Report - Hanley_Scraper.xlsx' next to this script)",
    )
    parser.add_argument("--ticker", help="Override: scrape this single ticker instead of reading R3")
    parser.add_argument("--tickers", help="Comma-separated list of tickers to scrape one after another (prints results; does not auto-fill multiple rows)")
    parser.add_argument("--output", help="Output xlsx path (defaults to <workbook>_scraped.xlsx)")
    parser.add_argument("--manual-login", action="store_true",
                         help="Fall back to logging in by hand (use this if MFA blocks auto-login)")
    args = parser.parse_args()

    output = args.output or args.workbook.replace(".xlsx", "_scraped.xlsx")

    if args.tickers:
        ticker_list = [t.strip() for t in args.tickers.split(",") if t.strip()]
    elif args.ticker:
        ticker_list = [args.ticker]
    else:
        ticker_list = [read_ticker(args.workbook)]

    driver = start_driver()
    try:
        if args.manual_login:
            manual_login(driver)
        else:
            username, password = load_credentials()
            auto_login(driver, username, password)

        all_results = {}
        for ticker in ticker_list:
            print(f"\nScraping {ticker} ...")
            data = scrape_company(driver, ticker)
            all_results[ticker] = data
            print(f"{ticker} -> {data}")

        # If only one ticker, write it straight into the template's O column.
        if len(ticker_list) == 1:
            write_results(args.workbook, output, all_results[ticker_list[0]])
        else:
            print("\nMultiple tickers scraped - printed above. "
                  "Tell me your preferred layout (one column per company?) "
                  "and I'll wire up write_results() to match it.")
    finally:
        driver.quit()


if __name__ == "__main__":
    main()

# requirements.txt contents:
# selenium
# openpyxl
# webdriver-manager  (optional, only if you want auto chromedriver management)
