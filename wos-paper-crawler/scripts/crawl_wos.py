"""
WoS Crawler - AER 2025 test. Edge + Selenium.
Key fixes: (1) gradual scroll for 50 articles/page
           (2) type journal + click 1st autocomplete suggestion
           (3) save actual journal name from article page
"""
import csv, time, random, re, os, sys
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

DRIVER_PATH = r"C:\Users\24882\Desktop\edgedriver_win32\msedgedriver.exe"
WOS_URL = "https://webofscience.clarivate.cn/wos/alldb/basic-search"
OUT = r"D:\Opencode\Scholar-city\WoS_QJE.csv"

STOP_TIME = datetime(2026, 6, 14, 22, 0, 0)

JOURNAL = "Quarterly Journal of Economics"

START_YEAR = 1995

driver = None

# ── browser ──
def init():
    global driver
    opts = Options()
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    driver = webdriver.Edge(service=Service(DRIVER_PATH), options=opts)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

def cookie_accept():
    try:
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "onetrust-accept-btn-handler")))
        driver.find_element(By.ID, "onetrust-accept-btn-handler").click()
        time.sleep(1)
    except: pass

def popups_close():
    for sel in ["[aria-label='Close']", ".close", ".btn-close", "button[mattooltip='Close']"]:
        for e in driver.find_elements(By.CSS_SELECTOR, sel):
            try: driver.execute_script("arguments[0].click();", e)
            except: pass

# ── search (type journal, click 1st suggestion) ──
def search_journal(name):
    driver.get(WOS_URL)
    time.sleep(4)
    cookie_accept()
    popups_close()
    time.sleep(1)

    # Select "Publication Titles" field
    print("  selecting Publication Titles field...")
    try:
        btn = WebDriverWait(driver, 12).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "#snSearchType button")))
        driver.execute_script("arguments[0].click();", btn)
        time.sleep(2.5)
        a = ActionChains(driver)
        for _ in range(3):
            a.send_keys(Keys.ARROW_DOWN); time.sleep(0.15)
        a.send_keys(Keys.ENTER); a.perform()
        time.sleep(1.5)
        print("  field selected")
    except Exception as e:
        print(f"  [WARN] field: {e}")

    # Type journal name directly (Publication Titles field is already selected)
    print(f"  typing: {name}")
    try:
        inp = driver.find_element(By.ID, "search-option-0")
        inp.clear()
        time.sleep(0.5)
        inp.send_keys(name)
        time.sleep(1.5)
        # No autocomplete needed for Publication Titles - just search
    except Exception as e:
        print(f"  [WARN] input: {e}")
        return False

    # Click search button
    print("  clicking search...")
    try:
        search_btn = driver.find_element(By.CSS_SELECTOR, "button[data-ta='run-search']")
        driver.execute_script("arguments[0].click();", search_btn)
        time.sleep(6)
        print("  search done")
    except Exception as e:
        print(f"  [WARN] search: {e}")
        return False

    # Sort newest first
    try:
        sort_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "selectSortOption")))
        driver.execute_script("arguments[0].click();", sort_btn)
        time.sleep(1.5)
        newest = driver.find_element(By.XPATH, "//span[contains(text(),'Date: newest first')]")
        driver.execute_script("arguments[0].click();", newest)
        time.sleep(4)
        print("  sorted newest first")
    except Exception as e:
        print(f"  [WARN] sort: {e}")

    pg = driver.page_source.lower()
    if "no results" in pg:
        print("  NO RESULTS")
        return False
    return True

# ── scroll: gradual step to load all 50 ──
def scroll_load():
    """Scroll to load all 50 articles per page. Tracks height growth."""
    last_h = driver.execute_script("return document.body.scrollHeight")
    pos = 0
    stable = 0
    for i in range(60):
        pos += 400
        if pos > last_h: pos = last_h
        driver.execute_script(f"window.scrollTo(0, {pos});")
        time.sleep(0.8)
        new_h = driver.execute_script("return document.body.scrollHeight")
        if new_h > last_h:
            last_h = new_h
            stable = 0
        else:
            stable += 1
        if pos >= last_h and stable >= 5:
            break
    driver.execute_script("window.scrollTo(0, 200);")
    time.sleep(1)
    
    # Second pass: scroll down again (catch missed lazy-loaded items)
    last_h = driver.execute_script("return document.body.scrollHeight")
    pos = 0
    for i in range(50):
        pos += 400
        if pos > last_h: pos = last_h
        driver.execute_script(f"window.scrollTo(0, {pos});")
        time.sleep(0.6)
        new_h = driver.execute_script("return document.body.scrollHeight")
        if new_h > last_h:
            print(f"  (+{new_h - last_h}px more content)")
            last_h = new_h
        if pos >= last_h:
            break
    
    driver.execute_script("window.scrollTo(0, 200);")
    time.sleep(1)

def get_links():
    scroll_load()
    seen = set()
    urls = []
    for a in driver.find_elements(By.XPATH, "//a[contains(@href,'full-record')]"):
        h = a.get_attribute("href")
        if h and "/full-record/WOS:" in h:
            wid = h.split("/full-record/")[-1] if "/full-record/" in h else h.split("WOS:")[-1]
            if wid not in seen:
                urls.append(h)
                seen.add(wid)
    return urls

def go_next():
    try:
        b = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "app-page-controls button:nth-child(4)")))
        if b.is_enabled():
            driver.execute_script("arguments[0].click();", b)
            time.sleep(3)
            return True
    except: pass
    return False

# ── extract article detail ──
def extract_article():
    cookie_accept()
    popups_close()
    time.sleep(2)
    r = {"title":"","authors":[],"journal":"","pub":"","ea":False,
         "kw":"","cat":"","dt":"","cw":"0","ca":"0","cr":"0",
         "addrs":[],"corr_name":"","corr_addr":""}

    try: r["title"] = driver.find_element(By.CSS_SELECTOR, "[data-ta='FullRTa-fullRecordtitle-0']").text.strip()
    except: pass

    for e in driver.find_elements(By.CSS_SELECTOR, "[id^='SumAuthTa-DisplayName-author-en-']"):
        t = e.text.strip()
        if t: r["authors"].append(t)

    # Actual journal name from WoS page
    try:
        r["journal"] = driver.find_element(By.XPATH,
            "//h3[@data-ta='FullRTa-sourceLabel']/following-sibling::span//a").text.strip()
    except: pass

    try:
        r["pub"] = driver.find_element(By.XPATH,
            "//h3[@data-ta='FullRTa-publishedLabel']/following-sibling::span").text.strip()
    except:
        try:
            r["pub"] = driver.find_element(By.XPATH,
                "//h3[@data-ta='FullRTa-earlyAccessDateLabel']/following-sibling::span").text.strip()
            r["ea"] = True
        except: pass

    try:
        k = [e.text.strip() for e in driver.find_elements(By.CSS_SELECTOR, "a.keywordsPlusLink") if e.text.strip()]
        r["kw"] = "; ".join(k)
    except: pass

    try:
        r["cat"] = driver.find_element(By.XPATH,
            "//*[contains(text(),'Web of Science Categories')]/following-sibling::span//a").text.strip()
    except: pass

    # Document Type: stable id `#FullRTa-doctype-0` contains the value
    try:
        r["dt"] = driver.find_element(By.ID, "FullRTa-doctype-0").text.strip()
    except: pass

    # citations
    try:
        e = driver.find_element(By.ID, "FullRRPTa-wos-citation-network-refCountLink")
        v = e.text.strip().replace(",", "")
        if v.isdigit(): r["cr"] = v
    except: pass
    try:
        e = driver.find_element(By.ID, "FullRRPTa-citationsLabelPlural-ALLDB")
        v = e.text.strip().replace(",", "")
        if v.isdigit(): r["ca"] = v
    except: pass
    try:
        e = driver.find_element(By.CSS_SELECTOR,
            "[id^='FullRRPTa-wos-citation-network-times-cited-count-link-']")
        v = e.text.strip().replace(",", "")
        if v.isdigit(): r["cw"] = v
    except: pass

    for e in driver.find_elements(By.CSS_SELECTOR, "[id^='address_']"):
        t = e.text.strip()
        if t: r["addrs"].append(t)

    try: r["corr_name"] = driver.find_element(By.CSS_SELECTOR, ".author-display-name").text.strip()
    except: pass
    try: r["corr_addr"] = driver.find_element(By.CSS_SELECTOR, "[data-ta='FRAOrgTa-RepAddressFull-0']").text.strip()
    except: pass

    return r

def _pd(s):
    if not s: return ""
    s = s.upper().strip()
    m = re.search(r'(\d{4})', s)
    if not m: return s
    y = m.group(1)
    for mon, mm in {"JAN":"01","FEB":"02","MAR":"03","APR":"04","MAY":"05","JUN":"06",
                    "JUL":"07","AUG":"08","SEP":"09","OCT":"10","NOV":"11","DEC":"12"}.items():
        if mon in s: return f"{y}/{mm}"
    return f"{y}/01"

# ── save ──
COLS = ["WoS_ID","Journal_Name","Article_Title","Source_Journal","Published_Date",
        "Early_Access","Keywords","WoS_Categories","Document_Type",
        "Citations_WoS_Core","Citations_All_DB","Cited_References",
        "Authors","Author_Count","Corr_Author","Corr_Address","Addresses","Address_Count"]

def load_done():
    s = set()
    if os.path.exists(OUT):
        with open(OUT, "r", encoding="utf-8-sig") as f:
            for row in csv.DictReader(f): s.add(row.get("WoS_ID",""))
    return s

def save(row):
    hdr = not os.path.exists(OUT)
    with open(OUT, "a", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=COLS, extrasaction="ignore")
        if hdr: w.writeheader()
        w.writerow(row)
        f.flush(); os.fsync(f.fileno())

# ── main ──
print(f"=== WoS Crawl: {JOURNAL} from {START_YEAR} ===")
print(f"=== Stop at: {STOP_TIME.strftime('%Y-%m-%d %H:%M')} ===")
sys.stdout.flush()
init()
sys.stdout.flush()
print("Browser started")
sys.stdout.flush()
done = load_done()
print(f"Already saved: {len(done)}")
sys.stdout.flush()

if not search_journal(JOURNAL):
    print("SEARCH FAILED")
    sys.stdout.flush()
    driver.quit()
    sys.exit(1)

print("Search OK, starting crawl...")
sys.stdout.flush()

page = 0
pre_count = 0
stop = False

# Skip to the right page based on already-crawled count
if done:
    skip_pages = len(done) // 50  # conservative: assume 50/page
    print(f"[Skip] Jumping to page ~{skip_pages + 1} (already have {len(done)} papers)")
    for _ in range(skip_pages):
        if not go_next():
            print(f"  Could not skip past page {_ + 1}")
            break
        time.sleep(2)
    page = skip_pages

while not stop:
    page += 1
    print(f"\n--- Page {page} ---")
    sys.stdout.flush()
    
    pg_src = driver.page_source.lower()
    if "unusual traffic" in pg_src or "verify you are" in pg_src:
        print("[BLOCKED]")
        break

    print("  scrolling...", end=""); sys.stdout.flush()
    links = get_links()
    print(f" done. Links found: {len(links)}")
    sys.stdout.flush()

    # Page 0 links => browser restart and re-navigate to this page
    if not links:
        print("  [RESTART] 0 links -- restarting browser...")
        try: driver.quit()
        except: pass
        time.sleep(3)
        init()
        if not search_journal(JOURNAL):
            print("[STOP] Search failed after restart")
            break
        # Navigate back to this page
        for _ in range(page - 1):
            if not go_next(): break
            time.sleep(2)
        time.sleep(2)
        continue

    new = [u for u in links if u.split("WOS:")[-1] not in done]
    print(f"  New to crawl: {len(new)}")

    if not new:
        print("  All done")
        if not go_next(): break
        time.sleep(2)
        continue

    for i, url in enumerate(new):
        wid = url.split("/full-record/")[-1] if "/full-record/" in url else url.split("WOS:")[-1]
        if wid in done: continue

        main = driver.current_window_handle
        driver.execute_script("window.open('');")
        driver.switch_to.window(driver.window_handles[-1])
        art = {}
        try:
            driver.get(url); time.sleep(2.5)
            art = extract_article()
        except Exception as e:
            print(f"  [{i+1}] {wid} ERR {e}")
        finally:
            try: driver.close()
            except: pass
            driver.switch_to.window(main)

        if not art or not art["title"]: continue

        dt = art.get("dt","")
        # Save ALL articles; document type filtering will be done in postprocess
        # (Chinese WoS portal uses different doc type labels than English)

        r = {
            "WoS_ID": wid,
            "Journal_Name": JOURNAL,  # search term (actual name in Source_Journal)
            "Article_Title": art["title"],
            "Source_Journal": art["journal"],  # ACTUAL journal name from page
            "Published_Date": _pd(art["pub"]),
            "Early_Access": str(art["ea"]),
            "Keywords": art["kw"],
            "WoS_Categories": art["cat"],
            "Document_Type": dt,
            "Citations_WoS_Core": art["cw"],
            "Citations_All_DB": art["ca"],
            "Cited_References": art["cr"],
            "Authors": " | ".join(art["authors"]),
            "Author_Count": str(len(art["authors"])),
            "Corr_Author": art["corr_name"],
            "Corr_Address": art["corr_addr"],
            "Addresses": " || ".join(art["addrs"]),
            "Address_Count": str(len(art["addrs"])),
        }
        save(r); done.add(wid)
        cw, ca, cr = r["Citations_WoS_Core"], r["Citations_All_DB"], r["Cited_References"]
        title_safe = art['title'][:60].encode('gbk', errors='replace').decode('gbk')
        jrnl_safe = art['journal'][:30].encode('gbk', errors='replace').decode('gbk')
        print(f"  [{i+1}/{len(new)}] {title_safe}")
        print(f"       j={jrnl_safe} cit={cw}/{ca} ref={cr} {_pd(art['pub'])}")
        sys.stdout.flush()
        
        # Stop at time limit
        if datetime.now() >= STOP_TIME:
            print(f"\n[Time limit reached: {datetime.now().strftime('%H:%M')}]")
            stop = True
            break
        
        time.sleep(random.uniform(2, 4))

    if stop: break
    if not go_next():
        print("No more pages")
        break
    time.sleep(random.uniform(2, 3))

driver.quit()
print(f"\nDone. {len(done)} papers in {OUT}")
