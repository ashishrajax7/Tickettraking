import os
import io
import re
import sys
import json
import base64
import time

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')
from datetime import datetime, date
from flask import Flask, request, jsonify, send_from_directory, send_file, make_response
from flask_cors import CORS
from dotenv import load_dotenv
import openpyxl
import requests
from PIL import Image

load_dotenv()

app = Flask(__name__, static_folder="static", static_url_path="/static")
CORS(app)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SAMPLE_FILE = os.path.join(BASE_DIR, "sample_upload.xlsx")
CONFIG_FILE = os.path.join(BASE_DIR, "config.json")
CACHE_FILE = os.path.join(BASE_DIR, "drive_cache.json")
SHEETS_DB_FILE = os.path.join(BASE_DIR, "saved_sheets_db.json")
SCRIPT_FILE = os.path.join(BASE_DIR, "google_apps_script.js")
LOCAL_IMG_DIR = os.path.join(BASE_DIR, "saved_images")

DRIVE_FOLDER_ID = os.environ.get("DRIVE_FOLDER_ID", "15c6PLPDtr9I3F8POfHUdwYv-ttdFtesA").strip()
DRIVE_FOLDER_URL = f"https://drive.google.com/drive/folders/{DRIVE_FOLDER_ID}?usp=sharing"

SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "1e9OyvaLMVyRh84adrx-zxuWwz2ZmT-WddZB6VY3fw7Y").strip()
SPREADSHEET_URL = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit"

DEFAULT_HEADERS = [
    "Channel",
    "Seller Name",
    "Return Date",
    "Return Id",
    "MP Date",
    "Days Left",
    "Invoice No.",
    "Order id",
    "Item SKU",
    "Amt",
    "AWB No.",
    "Courier Partner",
    "Platform Status",
    "Status",
    "Last Location",
    "Timestamp",
    "Last Sync",
    "Image",
    "Ticket no",
    "Dispute Reason",
    "Ticket Status",
    "Remark",
    "Action",
]

MONTH_NAMES = [
    "",
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
]

MONTH_ABBR_MAP = {
    "jan": 1, "january": 1,
    "feb": 2, "february": 2,
    "mar": 3, "march": 3,
    "apr": 4, "april": 4,
    "may": 5,
    "jun": 6, "june": 6,
    "jul": 7, "july": 7,
    "aug": 8, "august": 8,
    "sep": 9, "sept": 9, "september": 9,
    "oct": 10, "october": 10,
    "nov": 11, "november": 11,
    "dec": 12, "december": 12,
}

os.makedirs(LOCAL_IMG_DIR, exist_ok=True)

HYPERLINK_FORMULA_RE = re.compile(
    r'^=HYPERLINK\(\s*["\']([^"\']+)["\']\s*(?:[,;]\s*["\']([^"\']*)["\'])?\s*\)',
    re.IGNORECASE,
)
URL_RE = re.compile(r'(https?://[^\s"\'<>]+)', re.IGNORECASE)


def load_json_file(path, default):
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return default
    return default


def save_json_file(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def get_config():
    cfg = load_json_file(CONFIG_FILE, {})
    env_url = os.environ.get("APPS_SCRIPT_URL", "").strip()
    apps_script_url = env_url or cfg.get("apps_script_url", "").strip()
    return {
        "apps_script_url": apps_script_url,
        "from_env": bool(env_url),
        "drive_folder_id": DRIVE_FOLDER_ID,
        "drive_folder_url": DRIVE_FOLDER_URL,
        "spreadsheet_id": SPREADSHEET_ID,
        "spreadsheet_url": SPREADSHEET_URL,
    }


def get_cache():
    return load_json_file(CACHE_FILE, {})


def update_cache_entry(awb, entry):
    cache = get_cache()
    cache[str(awb).strip()] = entry
    save_json_file(CACHE_FILE, cache)


def sanitize_name(name, default=""):
    cleaned = re.sub(r'[<>:"/\\|?*]+', "_", str(name or "").strip())
    return cleaned or default


def build_image_filename(awb, courier):
    safe_awb = sanitize_name(awb, "UNKNOWN_AWB")
    safe_courier = sanitize_name(courier, "")
    if safe_courier:
        return f"{safe_awb}_{safe_courier}.png"
    return f"{safe_awb}.png"


def resolve_month_year_sheet_name(last_sync_str):
    """
    Parse 'Last Sync' string (e.g. '27-09-2026 12:20:50 PM', '2026-10-05', '28-Aug-2026')
    and return '<FullMonth>-<Year>' like 'September-2026', 'October-2026'.
    """
    s = str(last_sync_str or "").strip()
    if s:
        # 1. DD-MM-YYYY or DD/MM/YYYY
        m1 = re.search(r'\b(\d{1,2})[-/\.](\d{1,2})[-/\.](\d{4})\b', s)
        if m1:
            p1, p2, yr = int(m1.group(1)), int(m1.group(2)), int(m1.group(3))
            month_num = p2 if 1 <= p2 <= 12 else (p1 if 1 <= p1 <= 12 else 0)
            if 1 <= month_num <= 12:
                return f"{MONTH_NAMES[month_num]}-{yr}"

        # 2. YYYY-MM-DD or YYYY/MM/DD
        m2 = re.search(r'\b(\d{4})[-/\.](\d{1,2})[-/\.](\d{1,2})\b', s)
        if m2:
            yr, month_num = int(m2.group(1)), int(m2.group(2))
            if 1 <= month_num <= 12:
                return f"{MONTH_NAMES[month_num]}-{yr}"

        # 3. Text month like '24-Sep-2026' or 'Aug 26, 2026'
        m3 = re.search(r'\b([A-Za-z]{3,9})\b.*?\b(20\d{2})\b', s)
        if m3:
            m_name = m3.group(1).lower()
            yr = int(m3.group(2))
            if m_name in MONTH_ABBR_MAP:
                return f"{MONTH_NAMES[MONTH_ABBR_MAP[m_name]]}-{yr}"

    now = datetime.now()
    return f"{MONTH_NAMES[now.month]}-{now.year}"


def is_mock_row(values):
    val_str = " ".join(str(v or "").lower() for v in values)
    return "brand central" in val_str or any(f"inv-100{i}" in val_str for i in range(1, 6))


def build_fghj_key(values, headers=None, row_index=None):
    """
    Build unique composite key from columns F, G, H, J:
    F = Invoice No.
    G = Order id
    H = Item SKU
    J = AWB No.
    """
    idx_f, idx_g, idx_h, idx_j = 5, 6, 7, 9
    if headers:
        for i, h in enumerate(headers):
            hl = str(h or "").strip().lower()
            if "invoice" in hl:
                idx_f = i
            elif "order id" in hl or hl == "order":
                idx_g = i
            elif "sku" in hl:
                idx_h = i
            elif "awb" in hl or "tracking id" in hl:
                idx_j = i

    def get_v(idx):
        if 0 <= idx < len(values) and values[idx] is not None:
            v = str(values[idx]).strip()
            if v.endswith(".0"):
                try:
                    v = str(int(float(v)))
                except Exception:
                    pass
            return v.lower()
        return ""

    key = f"{get_v(idx_f)}|{get_v(idx_g)}|{get_v(idx_h)}|{get_v(idx_j)}"
    if key == "|||":
        c = get_v(0)
        s = get_v(1)
        if row_index is not None:
            return f"blank_{c}_{s}_{row_index}"
        return f"blank_{c}_{s}"
    return key


def format_cell_value(val):
    if val is None:
        return ""
    if isinstance(val, (datetime, date)):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, float):
        if val.is_integer():
            return str(int(val))
        return str(round(val, 2))
    return str(val).strip()


def extract_cell_value_and_link(val_cell, form_cell):
    raw_val = val_cell.value
    form_val = form_cell.value if form_cell is not None else None
    link_target = None

    if val_cell.hyperlink and val_cell.hyperlink.target:
        link_target = str(val_cell.hyperlink.target).strip()
    elif form_cell and form_cell.hyperlink and form_cell.hyperlink.target:
        link_target = str(form_cell.hyperlink.target).strip()

    if isinstance(form_val, str) and form_val.strip().upper().startswith("=HYPERLINK"):
        m = HYPERLINK_FORMULA_RE.match(form_val.strip())
        if m:
            if not link_target:
                link_target = m.group(1).strip()
            if raw_val is None or str(raw_val).strip().upper().startswith("=HYPERLINK"):
                raw_val = m.group(2) if m.group(2) is not None else "View Image"

    display_str = format_cell_value(raw_val)

    if not link_target and display_str:
        url_match = URL_RE.search(display_str)
        if url_match:
            link_target = url_match.group(1).strip()

    return display_str, link_target


def get_saved_lookup_by_key():
    """Return a dict mapping F&G&H&J key -> saved row dict across all saved month sheets."""
    db = load_json_file(SHEETS_DB_FILE, {"sheets": {}})
    lookup = {}
    for s_name, s_data in db.get("sheets", {}).items():
        headers = s_data.get("headers", DEFAULT_HEADERS)
        for r in s_data.get("rows", []):
            vals = r.get("values", [])
            key = r.get("key") or build_fghj_key(vals, headers)
            if key and key != "|||":
                lookup[key] = {
                    "values": vals,
                    "headers": headers,
                    "links": r.get("links", {}),
                    "sheet_name": s_name,
                }
    return lookup


def enrich_row_metadata(headers, formatted, links, row_num, cache, saved_lookup=None):
    channel_col_idx = 0
    awb_col_idx = -1
    courier_col_idx = -1
    image_col_idx = -1
    last_sync_col_idx = -1
    ticket_col_idx = -1
    reason_col_idx = -1
    remark_col_idx = -1

    for i, h in enumerate(headers):
        hl = str(h or "").lower()
        if "channel" in hl:
            channel_col_idx = i
        if "awb" in hl and awb_col_idx == -1:
            awb_col_idx = i
        if "courier" in hl and courier_col_idx == -1:
            courier_col_idx = i
        if any(k in hl for k in ["image", "photo", "pod", "media"]) and image_col_idx == -1:
            image_col_idx = i
        if "last sync" in hl and last_sync_col_idx == -1:
            last_sync_col_idx = i
        if "ticket" in hl and ticket_col_idx == -1:
            ticket_col_idx = i
        if ("dispute" in hl or "reason" in hl) and reason_col_idx == -1:
            reason_col_idx = i
        if "remark" in hl and remark_col_idx == -1:
            remark_col_idx = i

    row_key = build_fghj_key(formatted, headers)

    # Pre-fill Ticket no, Dispute Reason, Remark from previously saved sheet data if currently blank
    if saved_lookup and row_key in saved_lookup:
        saved_item = saved_lookup[row_key]
        s_headers = saved_item.get("headers", [])
        s_vals = saved_item.get("values", [])
        s_map = {str(s_headers[i]).lower(): s_vals[i] for i in range(min(len(s_headers), len(s_vals)))}

        if ticket_col_idx != -1 and not formatted[ticket_col_idx]:
            formatted[ticket_col_idx] = s_map.get("ticket no", "")
        if reason_col_idx != -1 and not formatted[reason_col_idx]:
            formatted[reason_col_idx] = s_map.get("dispute reason", "")
        if remark_col_idx != -1 and not formatted[remark_col_idx]:
            formatted[remark_col_idx] = s_map.get("remark", "")

    awb_val = formatted[awb_col_idx].strip() if (awb_col_idx != -1 and awb_col_idx < len(formatted)) else ""
    courier_val = (
        formatted[courier_col_idx].strip() if (courier_col_idx != -1 and courier_col_idx < len(formatted)) else ""
    )
    channel_val = (
        formatted[channel_col_idx].strip()
        if (channel_col_idx != -1 and channel_col_idx < len(formatted) and formatted[channel_col_idx].strip())
        else "General"
    )
    last_sync_val = (
        formatted[last_sync_col_idx].strip()
        if (last_sync_col_idx != -1 and last_sync_col_idx < len(formatted))
        else ""
    )
    target_month_sheet = resolve_month_year_sheet_name(last_sync_val)
    cached_drive = cache.get(awb_val) if awb_val else None

    if cached_drive and cached_drive.get("viewUrl") and image_col_idx != -1:
        links[str(image_col_idx)] = cached_drive["viewUrl"]

    row_dict = {headers[i]: formatted[i] for i in range(len(headers))}
    return {
        "_row_num": row_num,
        "key": row_key,
        "values": formatted,
        "links": links,
        "awb": awb_val,
        "courier": courier_val,
        "channel": channel_val,
        "last_sync": last_sync_val,
        "target_sheet": target_month_sheet,
        "image_filename": build_image_filename(awb_val, courier_val) if awb_val else "",
        "cached_drive": cached_drive,
        "data": row_dict,
    }


def parse_excel_workbook(raw_bytes, filename="uploaded.xlsx", prefill_from_saved=True):
    wb_val = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True)
    wb_form = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=False)
    cache = get_cache()
    saved_lookup = get_saved_lookup_by_key() if prefill_from_saved else {}
    sheets_data = {}

    for sheet_name in wb_val.sheetnames:
        ws_val = wb_val[sheet_name]
        ws_form = wb_form[sheet_name]

        rows_val = list(ws_val.iter_rows(values_only=False))
        rows_form = list(ws_form.iter_rows(values_only=False))

        if not rows_val:
            sheets_data[sheet_name] = {"headers": [], "rows": []}
            continue

        header_row_idx = 0
        for idx, r in enumerate(rows_val):
            if any(c.value is not None and str(c.value).strip() != "" for c in r):
                header_row_idx = idx
                break

        raw_headers = rows_val[header_row_idx]
        headers = [
            format_cell_value(c.value) if (c.value is not None and str(c.value).strip() != "") else f"Column {i+1}"
            for i, c in enumerate(raw_headers)
        ]

        # Ensure editable columns and Action exist if missing
        for req_col in ["Ticket no", "Dispute Reason", "Remark", "Action"]:
            if not any(h.strip().lower() == req_col.lower() for h in headers):
                headers.append(req_col)

        parsed_rows = []
        for offset, row_v in enumerate(rows_val[header_row_idx + 1 :], start=1):
            actual_row_idx = header_row_idx + offset
            row_f = rows_form[actual_row_idx] if actual_row_idx < len(rows_form) else row_v

            formatted = []
            links = {}

            for col_idx in range(len(headers)):
                c_val = row_v[col_idx] if col_idx < len(row_v) else None
                c_form = row_f[col_idx] if col_idx < len(row_f) else None
                if c_val is None:
                    formatted.append("")
                    continue
                disp, link_url = extract_cell_value_and_link(c_val, c_form)
                formatted.append(disp)
                if link_url:
                    links[str(col_idx)] = link_url

            if not any(cell != "" for cell in formatted) and not links:
                continue

            row_obj = enrich_row_metadata(
                headers, formatted, links, len(parsed_rows) + 1, cache, saved_lookup
            )
            parsed_rows.append(row_obj)

        sheets_data[sheet_name] = {
            "headers": headers,
            "rows": parsed_rows,
            "total_rows": len(parsed_rows),
            "total_columns": len(headers),
        }

    return {
        "filename": filename,
        "source_type": "excel_upload",
        "sheet_names": wb_val.sheetnames,
        "active_sheet": wb_val.sheetnames[0] if wb_val.sheetnames else None,
        "sheets": sheets_data,
        "config": get_config(),
    }


def download_and_convert_to_lossless_png(image_url):
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
        ),
        "Accept": "image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8",
    }

    drive_match = re.search(r"/file/d/([a-zA-Z0-9_-]+)", image_url)
    if drive_match:
        fid = drive_match.group(1)
        image_url = f"https://drive.google.com/uc?export=download&id={fid}"

    resp = requests.get(image_url, headers=headers, timeout=30, stream=True)
    resp.raise_for_status()
    raw_bytes = resp.content

    if raw_bytes.startswith(b"\x89PNG\r\n\x1a\n"):
        return raw_bytes

    img = Image.open(io.BytesIO(raw_bytes))
    if img.mode not in ("RGB", "RGBA", "L"):
        img = img.convert("RGBA" if "A" in img.mode or "transparency" in img.info else "RGB")

    out_buf = io.BytesIO()
    img.save(out_buf, format="PNG", compress_level=1)
    return out_buf.getvalue()


@app.route("/")
def index():
    return send_from_directory("static", "index.html")


@app.route("/icon.png")
def serve_brand_icon():
    icon_path = os.path.join(BASE_DIR, "icon.png")
    if os.path.exists(icon_path):
        return send_file(icon_path, mimetype="image/png")
    return "", 404


@app.route("/api/version", methods=["GET"])
def api_version():
    return jsonify({
        "ok": True,
        "version": "2.2.0",
        "description": "Ajio fast sync with step timings & comment extraction",
        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S")
    })


@app.route("/api/config", methods=["GET", "POST"])
def api_config():
    if request.method == "POST":
        body = request.get_json(silent=True) or {}
        cfg = load_json_file(CONFIG_FILE, {})
        if "apps_script_url" in body:
            cfg["apps_script_url"] = str(body.get("apps_script_url", "")).strip()
        if "spreadsheet_url" in body:
            cfg["spreadsheet_url"] = str(body.get("spreadsheet_url", "")).strip()
        save_json_file(CONFIG_FILE, cfg)
        return jsonify({"ok": True, "config": get_config()})
    return jsonify(get_config())


@app.route("/api/script-code", methods=["GET"])
def get_script_code():
    if os.path.exists(SCRIPT_FILE):
        with open(SCRIPT_FILE, "r", encoding="utf-8") as f:
            return jsonify({"ok": True, "code": f.read(), "folder_url": DRIVE_FOLDER_URL})
    return jsonify({"ok": False, "error": "Script file not found"}), 404


def get_active_spreadsheet_id():
    cfg = load_json_file(CONFIG_FILE, {})
    url_or_id = cfg.get("spreadsheet_url") or os.environ.get("SPREADSHEET_URL") or os.environ.get("SPREADSHEET_ID") or SPREADSHEET_ID
    url_or_id = str(url_or_id).strip()
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url_or_id)
    if m:
        return m.group(1).strip()
    return url_or_id


@app.route("/api/sheet-data", methods=["GET"])
def get_google_sheet_data():
    """
    Fetches LIVE data directly in real-time from the Google Sheet link (Anyone with the link).
    Pure in-memory streaming: does NOT store or cache data in any local file!
    """
    cache = get_cache()
    cfg = get_config()
    sheet_id = get_active_spreadsheet_id()

    sheets_out = {}
    sheet_names = []

    # Fetch live directly from Google Sheet export URL
    try:
        export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
        exp_resp = requests.get(export_url, timeout=15)
        if exp_resp.status_code == 200:
            wb = openpyxl.load_workbook(io.BytesIO(exp_resp.content), data_only=True)
            month_re = re.compile(
                r"^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec|January|February|March|April|May|June|July|August|September|October|November|December)(?:-\d{4})?$",
                re.IGNORECASE,
            )
            for s_name in wb.sheetnames:
                s_clean = s_name.strip()
                if month_re.match(s_clean):
                    sheet_names.append(s_clean)

            if not sheet_names:
                sheet_names = [s.strip() for s in wb.sheetnames]

            for s_name in sheet_names:
                if s_name not in wb.sheetnames:
                    continue
                ws = wb[s_name]
                all_r = list(ws.iter_rows(values_only=True))
                if not all_r:
                    continue

                raw_hdrs = [format_cell_value(c) for c in all_r[0] if c is not None and str(c).strip() != ""]
                headers = raw_hdrs.copy() if raw_hdrs else DEFAULT_HEADERS.copy()

                # Add Action column right next to Remark
                if not any(str(h).strip().lower() == "action" for h in headers):
                    rem_idx = next((i for i, h in enumerate(headers) if "remark" in str(h).lower()), -1)
                    if rem_idx != -1:
                        headers.insert(rem_idx + 1, "Action")
                    else:
                        headers.append("Action")

                # Data comes ONLY from Google Sheet — no local DB merge
                ts_col_idx = headers.index("Ticket Status") if "Ticket Status" in headers else -1
                rem_col_idx = headers.index("Remark") if "Remark" in headers else -1

                sheet_rows = []
                for idx, r_vals in enumerate(all_r[1:], 1):
                    if not any(c is not None and str(c).strip() != "" for c in r_vals):
                        continue

                    val_map = {}
                    for hi, hname in enumerate(raw_hdrs):
                        if hi < len(r_vals):
                            val_map[hname] = format_cell_value(r_vals[hi])
                        else:
                            val_map[hname] = ""

                    formatted = []
                    for h in headers:
                        if h == "Action":
                            formatted.append("")
                        else:
                            formatted.append(val_map.get(h, ""))

                    if is_mock_row(formatted):
                        continue

                    links = {}
                    row_obj = enrich_row_metadata(headers, formatted, links, len(sheet_rows) + 1, cache, None)
                    sheet_rows.append(row_obj)

                sheets_out[s_name] = {
                    "headers": headers,
                    "rows": sheet_rows,
                    "total_rows": len(sheet_rows),
                    "total_columns": len(headers),
                }
    except Exception as e:
        print(f"[LIVE GOOGLE SHEET FETCH ERROR] {e}")

    current_month_sheet = resolve_month_year_sheet_name("")
    if not sheet_names:
        sheet_names = [current_month_sheet]
        sheets_out[current_month_sheet] = {
            "headers": DEFAULT_HEADERS,
            "rows": [],
            "total_rows": 0,
            "total_columns": len(DEFAULT_HEADERS),
        }

    active_sheet = "September-2026" if "September-2026" in sheet_names else sheet_names[0]

    return jsonify(
        {
            "ok": True,
            "filename": "Google Sheet (Live Linked Data)",
            "source_type": "google_sheet",
            "sheet_names": sheet_names,
            "active_sheet": active_sheet,
            "sheets": sheets_out,
            "config": cfg,
        }
    )


@app.route("/api/save-sheet", methods=["POST"])
def save_to_google_sheet():
    """
    Saves/Updates rows into Month-Year tabs (e.g., 'September-2026', 'October-2026') based on 'Last Sync' column,
    using composite key F&G&H&J (Invoice No. | Order id | Item SKU | AWB No.) so existing keys are updated
    and never duplicated.
    """
    body = request.get_json(silent=True) or {}
    headers = body.get("headers") or DEFAULT_HEADERS
    req_sheet_name = (body.get("sheet_name") or "").strip()
    incoming_rows = body.get("rows") or []

    if not incoming_rows:
        return jsonify({"ok": False, "error": "No rows to save"}), 400

    # Locate Last Sync and Image columns
    last_sync_idx = -1
    image_idx = -1
    ticket_idx = -1
    ticket_status_idx = -1
    reason_idx = -1
    remark_idx = -1

    for i, h in enumerate(headers):
        hl = str(h or "").lower()
        if "last sync" in hl:
            last_sync_idx = i
        if any(k in hl for k in ["image", "photo", "pod", "media"]) and image_idx == -1:
            image_idx = i
        if "ticket status" in hl:
            ticket_status_idx = i
        elif "status" in hl and "platform" not in hl and ticket_status_idx == -1:
            ticket_status_idx = i
        elif "ticket" in hl and ticket_idx == -1:
            ticket_idx = i
        if ("dispute" in hl or "reason" in hl) and reason_idx == -1:
            reason_idx = i
        if "remark" in hl and remark_idx == -1:
            remark_idx = i

    db = load_json_file(SHEETS_DB_FILE, {"sheets": {}})
    if "sheets" not in db:
        db["sheets"] = {}

    rows_by_sheet_for_remote = {}
    inserted_count = 0
    updated_count = 0
    touched_sheets = set()

    for r in incoming_rows:
        vals = list(r.get("values") or [])
        while len(vals) < len(headers):
            vals.append("")
        links = dict(r.get("links") or {})
        image_url = r.get("image_url") or (links.get(str(image_idx)) if image_idx != -1 else "") or ""

        if image_idx != -1 and image_url:
            links[str(image_idx)] = image_url
            if not vals[image_idx] or "🖼️" in vals[image_idx]:
                vals[image_idx] = "View Image"

        last_sync_val = vals[last_sync_idx] if (last_sync_idx != -1 and last_sync_idx < len(vals)) else ""
        target_sheet = req_sheet_name or resolve_month_year_sheet_name(last_sync_val)
        touched_sheets.add(target_sheet)

        row_key = build_fghj_key(vals, headers)

        # Upsert into local mirror DB
        if target_sheet not in db["sheets"]:
            db["sheets"][target_sheet] = {"headers": headers, "rows": []}

        sheet_rows = db["sheets"][target_sheet]["rows"]
        db["sheets"][target_sheet]["headers"] = headers

        found_existing = False
        for existing_r in sheet_rows:
            ek = existing_r.get("key") or build_fghj_key(existing_r.get("values", []), headers)
            if ek == row_key and row_key != "|||":
                old_vals = existing_r.get("values", [])
                # Preserve existing Ticket / Reason / Remark / Status if incoming is empty
                for c_idx in [ticket_idx, ticket_status_idx, reason_idx, remark_idx]:
                    if c_idx != -1 and c_idx < len(vals) and not str(vals[c_idx]).strip():
                        if c_idx < len(old_vals) and str(old_vals[c_idx]).strip():
                            vals[c_idx] = old_vals[c_idx]
                existing_r["values"] = vals
                existing_r["links"] = links
                existing_r["key"] = row_key
                found_existing = True
                updated_count += 1
                break

        if not found_existing:
            sheet_rows.append({"key": row_key, "values": vals, "links": links})
            inserted_count += 1

        if target_sheet not in rows_by_sheet_for_remote:
            rows_by_sheet_for_remote[target_sheet] = []
        rows_by_sheet_for_remote[target_sheet].append(
            {
                "key": row_key,
                "values": vals,
                "image_url": image_url,
            }
        )

    save_json_file(SHEETS_DB_FILE, db)

    # Also push to Google Spreadsheet via Apps Script Web App
    cfg = get_config()
    apps_script_url = cfg.get("apps_script_url", "")
    remote_synced = False
    remote_message = ""

    if apps_script_url:
        try:
            gs_resp = requests.post(
                apps_script_url,
                json={
                    "action": "save_sheet_rows",
                    "headers": headers,
                    "rows_by_sheet": rows_by_sheet_for_remote,
                },
                timeout=45,
            )
            if gs_resp.ok:
                gs_json = gs_resp.json()
                if gs_json.get("ok"):
                    remote_synced = True
                    inserted_count = gs_json.get("inserted", inserted_count)
                    updated_count = gs_json.get("updated", updated_count)
                else:
                    remote_message = gs_json.get("error", "")
        except Exception as e:
            remote_message = str(e)

    return jsonify(
        {
            "ok": True,
            "inserted": inserted_count,
            "updated": updated_count,
            "sheets": list(touched_sheets),
            "remote_synced": remote_synced,
            "remote_message": remote_message,
        }
    )


@app.route("/api/save-synced-ticket", methods=["POST"])
def save_synced_ticket():
    """
    Saves an individually synced ticket status and remark immediately into saved_sheets_db.json
    and pushes the update to Google Sheets via Apps Script Web App.
    """
    body = request.get_json(force=True, silent=True) or {}
    ticket_id = str(body.get("ticket_id") or body.get("ticketId") or "").strip()
    order_id = str(body.get("order_id") or body.get("orderId") or "").strip()
    status = str(body.get("status") or "").strip()
    remark = str(body.get("remark") or body.get("lastComment") or "").strip()
    party_code = str(body.get("party_code") or body.get("partyCode") or "").strip()
    sheet_name = str(body.get("sheet_name") or body.get("sheetName") or "").strip()
    row_num = body.get("row_num") or body.get("rowNum")

    if not ticket_id and not order_id and not row_num:
        return jsonify({"ok": False, "error": "Missing ticket_id, order_id, or row_num"}), 400

    db = load_json_file(SHEETS_DB_FILE, {"sheets": {}})
    sheets = db.get("sheets", {})

    target_sheet = sheet_name
    if not target_sheet or target_sheet not in sheets:
        target_sheet = "September-2026" if "September-2026" in sheets else (list(sheets.keys())[0] if sheets else "September-2026")

    if target_sheet not in sheets:
        sheets[target_sheet] = {"headers": DEFAULT_HEADERS.copy(), "rows": []}

    sheet_obj = sheets[target_sheet]
    headers = sheet_obj.get("headers") or DEFAULT_HEADERS.copy()
    rows = sheet_obj.get("rows", [])

    # Find column indices
    ticket_idx = -1
    ticket_status_idx = -1
    remark_idx = -1
    order_idx = -1

    for i, h in enumerate(headers):
        hl = str(h or "").lower().strip()
        if "ticket status" in hl:
            ticket_status_idx = i
        elif "status" in hl and "platform" not in hl and ticket_status_idx == -1:
            ticket_status_idx = i
        if "ticket" in hl and "status" not in hl and ticket_idx == -1:
            ticket_idx = i
        if "remark" in hl and remark_idx == -1:
            remark_idx = i
        if "order" in hl and order_idx == -1:
            order_idx = i

    # Locate the target row in rows
    matched_row = None
    matched_idx = -1

    # 1. Match by Ticket ID
    if ticket_id and ticket_idx != -1:
        for idx, r in enumerate(rows):
            vals = r.get("values", [])
            if ticket_idx < len(vals) and str(vals[ticket_idx]).strip() == ticket_id:
                matched_row = r
                matched_idx = idx
                break

    # 2. Match by Order ID
    if not matched_row and order_id and order_idx != -1:
        for idx, r in enumerate(rows):
            vals = r.get("values", [])
            if order_idx < len(vals) and str(vals[order_idx]).strip() == order_id:
                matched_row = r
                matched_idx = idx
                break

    # 3. Match by row_num (1-based index)
    if not matched_row and row_num is not None:
        try:
            r_int = int(row_num) - 1
            if 0 <= r_int < len(rows):
                matched_row = rows[r_int]
                matched_idx = r_int
        except Exception:
            pass

    if not matched_row:
        return jsonify({"ok": False, "error": f"Row not found for ticket {ticket_id}"}), 404

    vals = list(matched_row.get("values", []))
    while len(vals) < len(headers):
        vals.append("")

    if ticket_status_idx != -1 and status:
        vals[ticket_status_idx] = status
    if remark_idx != -1 and remark:
        vals[remark_idx] = remark

    matched_row["values"] = vals
    save_json_file(SHEETS_DB_FILE, db)

    # Push to Google Spreadsheet via Apps Script Web App
    cfg = get_config()
    apps_script_url = cfg.get("apps_script_url", "")
    remote_synced = False
    if apps_script_url:
        try:
            row_key = matched_row.get("key") or build_fghj_key(vals, headers)
            gs_resp = requests.post(
                apps_script_url,
                json={
                    "action": "save_sheet_rows",
                    "headers": headers,
                    "rows_by_sheet": {
                        target_sheet: [
                            {
                                "key": row_key,
                                "ticket_id": ticket_id,
                                "order_id": order_id,
                                "values": vals,
                                "image_url": matched_row.get("links", {}).get(str(headers.index("Image") if "Image" in headers else -1), "")
                            }
                        ]
                    }
                },
                timeout=20,
            )
            if gs_resp.ok and gs_resp.json().get("ok"):
                remote_synced = True
        except Exception as e:
            print(f"[SAVE SYNCED TICKET REMOTE ERROR] {e}")

    print(f"[SAVE SYNCED TICKET] Saved Row {matched_idx + 1}: Ticket={ticket_id}, Status={status}, RemoteSynced={remote_synced}", flush=True)

    return jsonify({
        "ok": True,
        "sheet_name": target_sheet,
        "row_num": matched_idx + 1,
        "ticket_id": ticket_id,
        "status": status,
        "remark": remark,
        "remote_synced": remote_synced
    })


@app.route("/api/upload", methods=["POST"])
def upload_excel():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]
    if not file.filename:
        return jsonify({"error": "Empty filename"}), 400

    try:
        raw_bytes = file.read()
        result = parse_excel_workbook(raw_bytes, filename=file.filename, prefill_from_saved=True)
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": f"Failed to parse Excel file: {str(e)}"}), 500


@app.route("/api/local-image/<channel>/<filename>", methods=["GET"])
def serve_local_image(channel, filename):
    safe_channel = sanitize_name(channel, "General")
    safe_filename = sanitize_name(filename, "image.png")
    if not safe_filename.lower().endswith(".png"):
        safe_filename += ".png"
    img_path = os.path.join(LOCAL_IMG_DIR, safe_channel, safe_filename)
    if os.path.exists(img_path):
        return send_file(img_path, mimetype="image/png")
    return jsonify({"error": "Image not found locally"}), 404


@app.route("/api/sync-image", methods=["POST"])
def sync_image():
    body = request.get_json(silent=True) or {}
    awb = str(body.get("awb", "")).strip()
    courier = str(body.get("courier", "")).strip()
    channel = str(body.get("channel", "")).strip() or "General"
    image_url = str(body.get("image_url", "")).strip()
    force_drive = bool(body.get("force_drive", False))

    if not awb:
        return jsonify({"ok": False, "error": "Missing AWB number"}), 400

    file_name = build_image_filename(awb, courier)
    cache = get_cache()
    cfg = get_config()
    apps_script_url = cfg.get("apps_script_url", "")

    if awb in cache and not force_drive:
        cached = cache[awb]
        if cached.get("drive_synced") or not apps_script_url:
            return jsonify({"ok": True, "source": "cache", "data": cached})

    if apps_script_url and not image_url:
        try:
            check_resp = requests.post(
                apps_script_url,
                json={"action": "check", "awb": awb, "courier": courier, "channel": channel},
                timeout=20,
            )
            if check_resp.ok:
                check_data = check_resp.json()
                if check_data.get("ok") and check_data.get("found") and check_data.get("data"):
                    d = check_data["data"]
                    entry = {
                        "awb": awb,
                        "courier": courier,
                        "fileName": d.get("fileName", file_name),
                        "channel": d.get("channel", channel),
                        "fileId": d.get("fileId"),
                        "viewUrl": d.get("viewUrl"),
                        "directUrl": d.get("directUrl"),
                        "drive_synced": True,
                        "updated_at": datetime.now().isoformat(),
                    }
                    update_cache_entry(awb, entry)
                    return jsonify({"ok": True, "source": "drive_existing", "data": entry})
        except Exception:
            pass
        return jsonify({"ok": False, "error": "No image link and not found in Drive"}), 404

    if not image_url:
        return jsonify({"ok": False, "error": "No image URL provided"}), 400

    try:
        safe_channel = sanitize_name(channel, "General")
        channel_dir = os.path.join(LOCAL_IMG_DIR, safe_channel)
        os.makedirs(channel_dir, exist_ok=True)
        local_png_path = os.path.join(channel_dir, file_name)

        if os.path.exists(local_png_path):
            with open(local_png_path, "rb") as f:
                png_bytes = f.read()
        else:
            png_bytes = download_and_convert_to_lossless_png(image_url)
            with open(local_png_path, "wb") as f:
                f.write(png_bytes)

        local_view_url = f"/api/local-image/{safe_channel}/{file_name}"

        if apps_script_url:
            b64_png = base64.b64encode(png_bytes).decode("ascii")
            file_stem = file_name[:-4] if file_name.lower().endswith(".png") else file_name
            payload = {
                "action": "upload",
                "awb": file_stem,
                "rawAwb": awb,
                "courier": courier,
                "fileName": file_name,
                "channel": channel,
                "base64Data": b64_png,
                "imageUrl": image_url,
            }
            gs_resp = requests.post(apps_script_url, json=payload, timeout=45)
            gs_json = gs_resp.json()
            if gs_json.get("ok"):
                entry = {
                    "awb": awb,
                    "courier": courier,
                    "fileName": gs_json.get("fileName", file_name),
                    "channel": gs_json.get("channel", channel),
                    "fileId": gs_json.get("fileId"),
                    "viewUrl": gs_json.get("viewUrl"),
                    "directUrl": gs_json.get("directUrl"),
                    "localUrl": local_view_url,
                    "drive_synced": True,
                    "status": gs_json.get("status", "uploaded"),
                    "updated_at": datetime.now().isoformat(),
                }
                update_cache_entry(awb, entry)
                return jsonify({"ok": True, "source": "drive_upload", "data": entry})
            else:
                return jsonify(
                    {
                        "ok": False,
                        "error": gs_json.get("error", "Apps Script upload error"),
                        "localUrl": local_view_url,
                    }
                ), 500

        entry = {
            "awb": awb,
            "courier": courier,
            "fileName": file_name,
            "channel": channel,
            "viewUrl": local_view_url,
            "directUrl": local_view_url,
            "localUrl": local_view_url,
            "drive_synced": False,
            "status": "saved_locally_pending_script_url",
            "updated_at": datetime.now().isoformat(),
        }
        update_cache_entry(awb, entry)
        return jsonify({"ok": True, "source": "local_cache", "data": entry})

    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/sync-ticket-backend", methods=["POST", "OPTIONS"])
def api_sync_ticket_backend():
    if request.method == "OPTIONS":
        resp = make_response(jsonify({"ok": True}))
        resp.headers["Access-Control-Allow-Origin"] = "*"
        resp.headers["Access-Control-Allow-Methods"] = "POST, OPTIONS"
        resp.headers["Access-Control-Allow-Headers"] = "Content-Type"
        return resp

    body = request.get_json(silent=True) or {}
    party_code = str(body.get("partyCode") or body.get("party_code") or "").strip()
    ticket_id = str(body.get("ticketId") or body.get("ticket_id") or "").strip()
    order_id = str(body.get("orderId") or body.get("order_id") or "").strip()
    row_num = body.get("rowNum") or body.get("row_num")

    if not ticket_id:
        return jsonify({"ok": False, "error": "Missing ticketId"}), 400

    print(f"\n[API RECV] POST /api/sync-ticket-backend | Party: {party_code} | Ticket: {ticket_id} | Order: {order_id} | Row: {row_num}", flush=True)
    t_start = time.time()
    try:
        from ajio_scraper import scrape_single_ticket_sync
        res = scrape_single_ticket_sync(party_code, ticket_id, order_id)
        res["rowNum"] = row_num
        elapsed = round(time.time() - t_start, 2)
        res["serverElapsedSeconds"] = elapsed
        print(f"[API RESP] Ticket #{ticket_id} -> Status: {res.get('status')} | Comment: {res.get('lastComment', '')[:60]}... | Total Time: {elapsed}s\n", flush=True)
        resp = make_response(jsonify(res))
        resp.headers["Access-Control-Allow-Origin"] = "*"
        return resp
    except Exception as e:
        elapsed = round(time.time() - t_start, 2)
        print(f"[API ERR] Ticket #{ticket_id} Error: {e} | Time: {elapsed}s\n", flush=True)
        resp = make_response(jsonify({"ok": False, "error": str(e), "ticketId": ticket_id, "rowNum": row_num, "serverElapsedSeconds": elapsed}), 500)
        resp.headers["Access-Control-Allow-Origin"] = "*"
        return resp


@app.route("/api/batch-sync-backend", methods=["POST", "OPTIONS"])
def api_batch_sync_backend():
    if request.method == "OPTIONS":
        resp = make_response(jsonify({"ok": True}))
        resp.headers["Access-Control-Allow-Origin"] = "*"
        resp.headers["Access-Control-Allow-Methods"] = "POST, OPTIONS"
        resp.headers["Access-Control-Allow-Headers"] = "Content-Type"
        return resp

    body = request.get_json(silent=True) or {}
    tickets = body.get("tickets") or []
    if not tickets:
        return jsonify({"ok": False, "error": "No tickets provided"}), 400

    from flask import Response
    import queue
    import threading
    from ajio_scraper import scrape_batch_tickets_sync

    msg_queue = queue.Queue()

    def progress_callback(item_res):
        msg_queue.put(item_res)

    def worker():
        try:
            scrape_batch_tickets_sync(tickets, progress_callback)
        except Exception as e:
            msg_queue.put({"error": str(e)})
        finally:
            msg_queue.put(None) # Signal completion

    threading.Thread(target=worker, daemon=True).start()

    def generate():
        while True:
            try:
                item = msg_queue.get(timeout=90)
                if item is None:
                    yield f"data: {json.dumps({'type': 'BATCH_COMPLETE'})}\n\n"
                    break
                if "error" in item and len(item) == 1:
                    yield f"data: {json.dumps({'type': 'BATCH_ERROR', 'error': item['error']})}\n\n"
                    break
                yield f"data: {json.dumps({'type': 'TICKET_RESULT', **item})}\n\n"
            except queue.Empty:
                yield f"data: {json.dumps({'type': 'HEARTBEAT'})}\n\n"

    resp = Response(generate(), mimetype="text/event-stream")
    resp.headers["Access-Control-Allow-Origin"] = "*"
    resp.headers["Cache-Control"] = "no-cache"
    resp.headers["X-Accel-Buffering"] = "no"
@app.route("/api/sync-vendors", methods=["GET", "POST", "OPTIONS"])
def api_sync_vendors():
    if request.method == "OPTIONS":
        resp = make_response(jsonify({"ok": True}))
        resp.headers["Access-Control-Allow-Origin"] = "*"
        resp.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
        resp.headers["Access-Control-Allow-Headers"] = "Content-Type"
        return resp

    try:
        from ajio_scraper import sync_vendor_credentials_from_sheet
        vendors = sync_vendor_credentials_from_sheet(force=True)
        party_keys = sorted(
            list(vendors.keys()),
            key=lambda x: int(re.sub(r'\D', '', str(x))) if re.sub(r'\D', '', str(x)) else 999999
        )
        resp = make_response(jsonify({
            "ok": True,
            "count": len(vendors),
            "parties": party_keys,
            "message": f"Successfully synced {len(vendors)} party credentials from Google Sheet"
        }))
        resp.headers["Access-Control-Allow-Origin"] = "*"
        return resp
    except Exception as e:
        resp = make_response(jsonify({"ok": False, "error": str(e)}), 500)
        resp.headers["Access-Control-Allow-Origin"] = "*"
        return resp


@app.route("/api/extension-log", methods=["POST", "OPTIONS"])
def extension_log():
    if request.method == "OPTIONS":
        resp = make_response(jsonify({"ok": True}))
        resp.headers["Access-Control-Allow-Origin"] = "*"
        resp.headers["Access-Control-Allow-Methods"] = "POST, OPTIONS"
        resp.headers["Access-Control-Allow-Headers"] = "Content-Type"
        return resp

    try:
        body = request.get_json(force=True, silent=True) or {}
        level = body.get("level", "INFO")
        tag = body.get("tag", "EXT")
        message = body.get("message", "")
        details = body.get("details", "")
        url = body.get("url", "")
        print(f"[{level}][{tag}] {message} | {details} | URL: {url}", flush=True)
        resp = make_response(jsonify({"ok": True}))
        resp.headers["Access-Control-Allow-Origin"] = "*"
        return resp
    except Exception as e:
        resp = make_response(jsonify({"ok": False, "error": str(e)}), 200)
        resp.headers["Access-Control-Allow-Origin"] = "*"
        return resp


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
