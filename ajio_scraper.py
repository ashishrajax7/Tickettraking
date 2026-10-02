import asyncio
import os
import re
import sys
import json
import time
import logging
import requests
from playwright.async_api import async_playwright

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("AjioScraper")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
VENDORS_FILE = os.path.join(BASE_DIR, "ajio_vendors.json")
AJIO_CREDENTIALS_SHEET_URL = os.environ.get(
    "AJIO_CREDENTIALS_SHEET_URL",
    "https://script.google.com/macros/s/AKfycbzO2hojy-wjdUYZXyUgF-COHQIHJ1EKffJQBxmOfDnMBbBk29H2zvPoBXq8IoYduG3aIQ/exec"
)

_VENDOR_CACHE = {}
_VENDOR_CACHE_TIME = 0
_CACHE_TTL = 900  # 15 minutes cache

def sync_vendor_credentials_from_sheet(force: bool = False) -> dict:
    """
    Fetch live credentials directly from Google Sheet Apps Script.
    Caches in memory and updates local ajio_vendors.json for persistent offline fallback.
    """
    global _VENDOR_CACHE, _VENDOR_CACHE_TIME
    now = time.time()
    if not force and _VENDOR_CACHE and (now - _VENDOR_CACHE_TIME < _CACHE_TTL):
        return _VENDOR_CACHE

    sheet_fetched = False
    try:
        logger.info(f"Syncing Ajio vendor credentials from Google Sheet: {AJIO_CREDENTIALS_SHEET_URL}")
        resp = requests.get(AJIO_CREDENTIALS_SHEET_URL, timeout=12)
        if resp.status_code == 200:
            sheet_data = resp.json()
            if isinstance(sheet_data, dict) and len(sheet_data) > 0:
                sheet_fetched = True
                existing = {}
                if os.path.exists(VENDORS_FILE):
                    try:
                        with open(VENDORS_FILE, "r", encoding="utf-8") as f:
                            existing = json.load(f)
                    except Exception:
                        existing = {}

                merged = {**existing, **sheet_data}
                try:
                    with open(VENDORS_FILE, "w", encoding="utf-8") as f:
                        json.dump(merged, f, indent=2)
                except Exception as e:
                    logger.warning(f"Failed to write merged vendors to file: {e}")

                _VENDOR_CACHE = merged
                _VENDOR_CACHE_TIME = now
                logger.info(f"Successfully synced {len(sheet_data)} parties from Google Sheet (total {len(merged)} available).")
                return _VENDOR_CACHE
    except Exception as e:
        logger.warning(f"Could not fetch vendor credentials from Google Sheet: {e}. Falling back to local store.")

    # Fallback to local file if sheet request didn't update cache
    if not sheet_fetched and os.path.exists(VENDORS_FILE):
        try:
            with open(VENDORS_FILE, "r", encoding="utf-8") as f:
                _VENDOR_CACHE = json.load(f)
                _VENDOR_CACHE_TIME = now
                logger.info(f"Loaded {len(_VENDOR_CACHE)} parties from local cache file.")
        except Exception as e:
            logger.error(f"Failed to load local vendors file: {e}")

    return _VENDOR_CACHE

def get_vendor_credentials(party_code: str) -> dict:
    """Fetch email & password for a party code from Google Sheet (with local cache fallback)."""
    vendors = sync_vendor_credentials_from_sheet(force=False)

    code_clean = str(party_code or "").strip()
    digits = re.search(r'\b(\d{1,4})\b', code_clean)
    num_code = digits.group(1) if digits else code_clean

    keys_to_check = [num_code, f"AJ{num_code}", code_clean, code_clean.upper()]
    for k in keys_to_check:
        if k in vendors and isinstance(vendors[k], dict):
            entry = vendors[k]
            email = entry.get("email") or entry.get("username")
            pwd = entry.get("password")
            if email and pwd:
                return {"email": str(email).strip(), "password": str(pwd).strip(), "partyCode": num_code}

    # If not found in current cache, attempt forced refresh from Google Sheet
    if not any(k in vendors for k in keys_to_check):
        logger.info(f"Party {num_code} not found in cache. Forcing fresh pull from Google Sheet...")
        vendors = sync_vendor_credentials_from_sheet(force=True)
        for k in keys_to_check:
            if k in vendors and isinstance(vendors[k], dict):
                entry = vendors[k]
                email = entry.get("email") or entry.get("username")
                pwd = entry.get("password")
                if email and pwd:
                    return {"email": str(email).strip(), "password": str(pwd).strip(), "partyCode": num_code}

    return {}

LAUNCH_ARGS = [
    '--no-sandbox',
    '--disable-setuid-sandbox',
    '--disable-dev-shm-usage',
    '--disable-blink-features=AutomationControlled',
    '--disable-gpu',
    '--no-zygote',
    '--js-flags=--max-old-space-size=256'
]

async def _login_and_open_help_center(page, email: str, password: str) -> bool:
    """Logs into Reliance SSO and lands on Help Center table rapidly."""
    logger.info(f"Navigating to seller.ajio.com/vmsui/ with email {email}...")
    await page.goto("https://seller.ajio.com/vmsui/", timeout=35000, wait_until="domcontentloaded")

    # Check for login inputs
    try:
        user_inp = await page.wait_for_selector('input[type="email"], input[name="username"], #username, input[type="text"]', timeout=15000)
        if user_inp:
            logger.info("Filling credentials...")
            await user_inp.fill(email)
            pwd_inp = await page.wait_for_selector('input[type="password"], #password, input[name="password"]', timeout=6000)
            if pwd_inp:
                await pwd_inp.fill(password)
            submit_btn = await page.wait_for_selector('button[type="submit"], input[type="submit"], button:has-text("Sign In"), button:has-text("Login")', timeout=6000)
            if submit_btn:
                await submit_btn.click()
            
            # Wait until browser leaves oauth/login URL
            await page.wait_for_function(
                "() => !window.location.href.includes('login') && !window.location.href.includes('callback') && !window.location.href.includes('oauth')",
                timeout=25000
            )
            logger.info(f"Successfully authenticated. Dashboard URL: {page.url}")
    except Exception as e:
        logger.info(f"SSO step check: {e}")

    # Wait 2.5s for announcement dialog and React components to mount
    await asyncio.sleep(2.5)

    # Check for and dismiss announcement modal if present
    try:
        dialog_ok = page.locator('.MuiDialog-root button:has-text("OK"), .MuiDialog-root button[aria-label="Close"], .MuiDialog-root button:has(svg[data-testid="CloseIcon"])').first
        if await dialog_ok.count() > 0 and await dialog_ok.is_visible():
            logger.info("Dismissing announcement popup...")
            await dialog_ok.click()
            await asyncio.sleep(0.5)
    except Exception as ex_dlg:
        logger.debug(f"Announcement modal note: {ex_dlg}")

    # Remove any stray backdrops just in case
    await page.evaluate("""() => {
        document.querySelectorAll('.MuiBackdrop-root').forEach(b => b.remove());
    }""")

    # Click Help Center icon in navbar
    logger.info("Clicking Help Center button...")
    help_btn = page.locator('button[aria-label="Help Center"], svg[data-testid="HelpOutlineIcon"]').first
    await help_btn.click(force=True)

    # Fast wait for Help Center filter inputs
    try:
        await page.wait_for_selector('.rt-thead.-filters input[type="text"]', timeout=12000)
        logger.info("Help Center filter inputs ready.")
        return True
    except Exception as ex:
        logger.warning(f"Filter inputs not ready on first try: {ex}")
        try:
            help_btn = page.locator('button[aria-label="Help Center"], svg[data-testid="HelpOutlineIcon"]').first
            if await help_btn.count() > 0:
                await help_btn.click(force=True)
                await page.wait_for_selector('.rt-thead.-filters input[type="text"]', timeout=10000)
                return True
        except Exception as ex2:
            logger.warning(f"Help icon retry note: {ex2}")
        return False

async def _scrape_ticket_in_page(page, ticket_id: str, order_id: str = None) -> dict:
    """Scrapes status and comments for a single ticket on an open Help Center page."""
    clean_ticket = str(ticket_id or "").strip()
    clean_order = str(order_id or "").strip()

    # Filter by ticket ID in column 2
    filter_inputs = page.locator('.rt-thead.-filters input[type="text"]')
    count_inp = await filter_inputs.count()
    if count_inp >= 2:
        ticket_box = filter_inputs.nth(1)
        await ticket_box.scroll_into_view_if_needed()
        await ticket_box.fill(clean_ticket)
        await ticket_box.press("Enter")

    # Fast dynamic wait for row to appear
    found_target = False
    row_txt = ""
    try:
        await page.wait_for_function(
            """(ticket) => {
                const rows = document.querySelectorAll('.rt-tbody .rt-tr:not(.-padRow)');
                return Array.from(rows).some(r => (r.textContent || '').includes(ticket));
            }""",
            arg=clean_ticket,
            timeout=8000
        )
        found_target = True
        row_txt = await page.evaluate(f"""(ticket) => {{
            const rows = Array.from(document.querySelectorAll('.rt-tbody .rt-tr:not(.-padRow)'));
            const target = rows.find(r => (r.textContent || '').includes(ticket));
            return target ? (target.textContent || '').toUpperCase() : '';
        }}""", clean_ticket)
    except Exception:
        row_txt = await page.evaluate(f"""(ticket) => {{
            const rows = Array.from(document.querySelectorAll('.rt-tbody .rt-tr:not(.-padRow)'));
            const target = rows.find(r => (r.textContent || '').includes(ticket));
            return target ? (target.textContent || '').toUpperCase() : '';
        }}""", clean_ticket)
        found_target = bool(row_txt)

    status_text = "UNKNOWN"
    for s in ['REJECTED', 'APPROVED', 'OPEN', 'IN PROGRESS', 'CLOSED', 'PENDING', 'RESOLVED']:
        if s in row_txt:
            status_text = s
            break

    # Click ticket button to open details modal
    ticket_btn = page.locator('.rt-tbody .rt-tr:not(.-padRow) button').first
    btn_exists = await ticket_btn.count() > 0
    if not btn_exists or not found_target:
        return {
            "ok": False,
            "error": f"Ticket #{clean_ticket} not found in Help Center table",
            "status": status_text,
            "lastComment": ""
        }

    await ticket_btn.click(force=True)

    # Wait for detail dialog
    try:
        await page.wait_for_selector('.MuiDialog-root .MuiAccordion-root', timeout=6000)
    except Exception:
        await asyncio.sleep(0.5)

    # Expand Comments accordion inside the active detail modal dialog
    await page.evaluate("""() => {
        const dialogs = document.querySelectorAll('.MuiDialog-root, [role="dialog"]');
        const activeDlg = dialogs.length > 0 ? dialogs[dialogs.length - 1] : document.body;
        const accordions = Array.from(activeDlg.querySelectorAll('.MuiAccordion-root'));
        const commentsAcc = accordions.find(a => (a.querySelector('.MuiAccordionSummary-root')?.textContent || '').toLowerCase().includes('comment'));
        if (commentsAcc) {
            const summary = commentsAcc.querySelector('.MuiAccordionSummary-root') || commentsAcc;
            if (summary.getAttribute('aria-expanded') !== 'true') {
                summary.click();
            }
        }
    }""")

    # Dynamic wait for comment network data to arrive from Ajio API
    try:
        await page.wait_for_function(
            """() => {
                const dialogs = document.querySelectorAll('.MuiDialog-root, [role="dialog"]');
                const activeDlg = dialogs.length > 0 ? dialogs[dialogs.length - 1] : document.body;
                const txt = activeDlg.innerText || '';
                return txt.includes('Dear Seller') || txt.includes('Team AJIO') || txt.includes('SPF') || txt.includes('Attachment') || txt.length > 400;
            }""",
            timeout=4000
        )
    except Exception:
        await asyncio.sleep(0.5)

    # Scrape comments targeting activeDlg
    raw_comment = await page.evaluate("""() => {
        const dialogs = document.querySelectorAll('.MuiDialog-root, [role="dialog"]');
        const activeDlg = dialogs.length > 0 ? dialogs[dialogs.length - 1] : document.body;

        // Clean HTML br tags helper
        const cleanText = (str) => {
            return (str || '')
                .replace(/<br\\s*[\\/]?>/gi, '\\n')
                .replace(/&nbsp;/gi, ' ')
                .replace(/<[^>]+>/g, '')
                .replace(/[ \\t]+/g, ' ')
                .trim();
        };

        // Strategy 1: Match full dialog text for Dear Seller ... Team AJIO
        const fullTxt = activeDlg.innerText || '';
        const match = fullTxt.match(/Dear Seller[\\s\\S]*?(?:Team AJIO|$)/i);
        if (match) return cleanText(match[0]);

        // Strategy 2: Check accordion details
        const accordions = Array.from(activeDlg.querySelectorAll('.MuiAccordion-root'));
        const commentsAcc = accordions.find(a => (a.querySelector('.MuiAccordionSummary-root')?.textContent || '').toLowerCase().includes('comment'));
        if (commentsAcc) {
            const details = commentsAcc.querySelector('.MuiAccordionDetails-root') || commentsAcc;
            const textBlocks = Array.from(details.querySelectorAll('p, span, div, td, li'))
                .map(el => cleanText(el.innerText || el.textContent))
                .filter(t => t.length > 15 && !t.includes('Comments'));
            if (textBlocks.length > 0) {
                return textBlocks[textBlocks.length - 1];
            }
        }

        // Strategy 3: Check all typography nodes
        const allNodes = Array.from(activeDlg.querySelectorAll('h5, p, span, div, .MuiTypography-root, .MuiCardContent-root'));
        for (let i = allNodes.length - 1; i >= 0; i--) {
            const txt = (allNodes[i].textContent || '').trim();
            if (txt.includes("Dear Seller") || txt.includes("Team AJIO")) {
                const m2 = txt.match(/Dear Seller[\\s\\S]*?(?:Team AJIO|$)/i);
                if (m2) return cleanText(m2[0]);
                return cleanText(txt);
            }
        }

        return '';
    }""")

    last_comment = raw_comment or ""
    if last_comment:
        last_comment = re.sub(r'<br\s*[\/]?>', '\n', last_comment, flags=re.I)
        last_comment = re.sub(r'<[^>]*>', ' ', last_comment)
        last_comment = re.sub(r'&lt;[^&]*&gt;', ' ', last_comment)
        last_comment = re.sub(r'\s+', ' ', last_comment).strip()

    # Close ticket detail dialog so next ticket has clean DOM
    try:
        await page.keyboard.press("Escape")
        await asyncio.sleep(0.3)
        await page.evaluate("""() => {
            const dialogs = document.querySelectorAll('.MuiDialog-root, [role="dialog"]');
            const activeDlg = dialogs.length > 0 ? dialogs[dialogs.length - 1] : document.body;
            const closeSvgs = Array.from(activeDlg.querySelectorAll('svg[data-testid="CloseIcon"]'));
            if (closeSvgs.length > 0) {
                const topBtn = closeSvgs[closeSvgs.length - 1].closest('button');
                if (topBtn) topBtn.click();
            }
        }""")
        await asyncio.sleep(0.3)
    except Exception:
        pass

    return {
        "ok": True,
        "ticketId": clean_ticket,
        "orderId": clean_order,
        "status": status_text,
        "lastComment": last_comment
    }

async def scrape_single_ticket_async(party_code: str, ticket_id: str, order_id: str = None) -> dict:
    """End-to-end scraper for a single ticket with detailed timing."""
    t0 = time.time()
    logger.info(f"[SCRAPE START] Party: {party_code} | Ticket: {ticket_id} | Order: {order_id or 'N/A'}")
    
    creds = get_vendor_credentials(party_code)
    if not creds or not creds.get("email") or not creds.get("password"):
        elapsed = round(time.time() - t0, 2)
        logger.error(f"[CREDS FAILED] Credentials not found for Party {party_code} (Took {elapsed}s)")
        return {"ok": False, "error": f"Credentials not found for Party {party_code}", "elapsedSeconds": elapsed}

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True, args=LAUNCH_ARGS)
        context = await browser.new_context(
            viewport={'width': 1366, 'height': 850},
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
        )
        page = await context.new_page()

        try:
            opened = await _login_and_open_help_center(page, creds["email"], creds["password"])
            if not opened:
                await browser.close()
                elapsed = round(time.time() - t0, 2)
                logger.error(f"[LOGIN FAILED] Could not access Help Center for Party {party_code} (Took {elapsed}s)")
                return {"ok": False, "error": "Could not access Help Center", "elapsedSeconds": elapsed}

            login_time = round(time.time() - t0, 2)
            logger.info(f"[AUTH SUCCESS] Dashboard & Help Center loaded in {login_time}s")

            result = await _scrape_ticket_in_page(page, ticket_id, order_id)
            result["partyCode"] = party_code
            elapsed = round(time.time() - t0, 2)
            result["elapsedSeconds"] = elapsed
            await browser.close()

            logger.info(f"[SCRAPE DONE] Ticket #{ticket_id} -> Status: {result.get('status')} | Comment: {result.get('lastComment', '')[:60]}... | Total Time: {elapsed}s")
            return result
        except Exception as e:
            await browser.close()
            elapsed = round(time.time() - t0, 2)
            logger.error(f"[SCRAPE ERROR] Error scraping ticket {ticket_id}: {e} (Took {elapsed}s)")
            return {"ok": False, "error": str(e), "partyCode": party_code, "ticketId": ticket_id, "elapsedSeconds": elapsed}

async def scrape_batch_tickets_async(tickets: list, progress_callback=None) -> list:
    """
    Batched scraper: groups tickets by partyCode.
    For each party, logs in ONCE and processes all tickets sequentially in that session.
    """
    batch_t0 = time.time()
    total_count = len(tickets)
    logger.info(f"[BATCH START] Starting batch scrape for {total_count} tickets")

    party_map = {}
    for t in tickets:
        p = str(t.get("partyCode") or t.get("party_code") or "262").strip()
        party_map.setdefault(p, []).append(t)

    results = []
    processed_count = 0

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True, args=LAUNCH_ARGS)

        for party_code, party_tickets in party_map.items():
            creds = get_vendor_credentials(party_code)
            if not creds:
                for item in party_tickets:
                    processed_count += 1
                    err_res = {
                        "ok": False,
                        "partyCode": party_code,
                        "ticketId": item.get("ticketId"),
                        "rowNum": item.get("rowNum"),
                        "status": "UNKNOWN",
                        "lastComment": "",
                        "error": f"Credentials not found for Party {party_code}",
                        "currentIndex": processed_count,
                        "totalCount": total_count,
                        "elapsedSeconds": 0
                    }
                    results.append(err_res)
                    if progress_callback:
                        progress_callback(err_res)
                continue

            party_t0 = time.time()
            logger.info(f"[BATCH PARTY START] Processing {len(party_tickets)} tickets for Party {party_code}")

            context = await browser.new_context(
                viewport={'width': 1366, 'height': 850},
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
            )
            page = await context.new_page()

            try:
                await _login_and_open_help_center(page, creds["email"], creds["password"])
                logger.info(f"[BATCH LOGIN OK] Logged in for Party {party_code} in {round(time.time() - party_t0, 2)}s")

                for item in party_tickets:
                    item_t0 = time.time()
                    t_id = item.get("ticketId")
                    o_id = item.get("orderId")
                    r_num = item.get("rowNum")
                    processed_count += 1

                    try:
                        t_res = await _scrape_ticket_in_page(page, t_id, o_id)
                        item_elapsed = round(time.time() - item_t0, 2)
                        t_res["partyCode"] = party_code
                        t_res["rowNum"] = r_num
                        t_res["currentIndex"] = processed_count
                        t_res["totalCount"] = total_count
                        t_res["elapsedSeconds"] = item_elapsed
                        results.append(t_res)
                        logger.info(f"[BATCH ITEM {processed_count}/{total_count}] Ticket #{t_id} -> Status: {t_res.get('status')} (Took {item_elapsed}s)")
                        if progress_callback:
                            progress_callback(t_res)
                    except Exception as ex:
                        item_elapsed = round(time.time() - item_t0, 2)
                        err_res = {
                            "ok": False,
                            "partyCode": party_code,
                            "ticketId": t_id,
                            "rowNum": r_num,
                            "status": "UNKNOWN",
                            "lastComment": "",
                            "error": str(ex),
                            "currentIndex": processed_count,
                            "totalCount": total_count,
                            "elapsedSeconds": item_elapsed
                        }
                        results.append(err_res)
                        logger.error(f"[BATCH ITEM ERROR {processed_count}/{total_count}] Ticket #{t_id}: {ex}")
                        if progress_callback:
                            progress_callback(err_res)

            except Exception as e:
                logger.error(f"Error in party {party_code} session: {e}")
            finally:
                await context.close()

        await browser.close()
    
    total_batch_time = round(time.time() - batch_t0, 2)
    logger.info(f"[BATCH FINISHED] All {total_count} tickets completed in {total_batch_time}s")
    return results

def scrape_single_ticket_sync(party_code: str, ticket_id: str, order_id: str = None) -> dict:
    """Synchronous entry point for Flask routes."""
    return asyncio.run(scrape_single_ticket_async(party_code, ticket_id, order_id))

def scrape_batch_tickets_sync(tickets: list, progress_callback=None) -> list:
    """Synchronous entry point for batch sync."""
    return asyncio.run(scrape_batch_tickets_async(tickets, progress_callback))
