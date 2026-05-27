"""
Spending pipeline — chạy hàng tháng, fetch giao dịch từ Gmail (VCB + TPBank),
parse sao kê TPBank TKTT PDF, categorize, upload Google Sheet, gửi email.

MODES:
    preliminary  (default): full pipeline + detect ambiguous + Telegram ask + PRELIM email
    final                 : re-run with updated overrides + FINAL email + clear Pending tab

USAGE (local):
    python spending_pipeline.py                          # auto previous month, prelim
    python spending_pipeline.py --mode final
    python spending_pipeline.py --month 2026-04 --include-quarter

USAGE (cloud — env vars):
    TOKEN_JSON      = nội dung token.json
    SHEET_ID        = Google Sheet master ID
    NOTIFY_EMAILS   = "to1@gmail.com,to2@gmail.com" (comma-sep)
    OVERRIDES_JSON  = (optional) overrides per txn_id
    TELEGRAM_TOKEN  = bot token
    TELEGRAM_CHAT_ID = chat ID

AMBIGUOUS DETECTION:
    category == "Khác / Chuyển khoản cá nhân" AND amount >= 2M VND
    → list in Telegram + sheet tab "Pending YYYY-MM" + PRELIM email note
"""
import argparse
import base64
import calendar
import csv
import io
import json
import os
import re
import sys
import tempfile
from collections import defaultdict
from datetime import date, datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

import pdfplumber
import pypdf
import requests
from bs4 import BeautifulSoup  # for HTML email body

HERE = Path(__file__).parent

# ---------------- Config ----------------
SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
DEFAULT_USD_VND = 26000
TPBANK_CARDS = {"7606": "Visa Platinum", "8272": "JCB Cashback"}
AMBIGUOUS_CATEGORY = "Khác / Chuyển khoản cá nhân"
AMBIGUOUS_MIN_VND = 2_000_000


# ---------------- Auth ----------------
def get_creds():
    """Load creds from env (cloud) or token.json (local)."""
    token_data = os.environ.get("TOKEN_JSON")
    if token_data:
        info = json.loads(token_data)
        creds = Credentials.from_authorized_user_info(info, SCOPES)
    else:
        token_path = HERE / "token.json"
        if not token_path.exists():
            print("ERROR: no token. Run auth.py first or set TOKEN_JSON env.", file=sys.stderr)
            sys.exit(1)
        creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)
    if not creds.valid:
        if creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            raise SystemExit("Token invalid — re-run auth.py")
    return creds


def svc(name, ver, creds): return build(name, ver, credentials=creds, cache_discovery=False)


# ---------------- Gmail helpers ----------------
def b64url_decode(data):
    pad = (4 - len(data) % 4) % 4
    return base64.urlsafe_b64decode(data + "=" * pad)


def walk_parts(payload):
    yield payload
    for p in payload.get("parts", []) or []:
        yield from walk_parts(p)


def get_text(payload):
    """Extract text/plain + text/html body concatenated."""
    out = []
    for part in walk_parts(payload):
        mt = part.get("mimeType", "")
        body = part.get("body", {})
        data = body.get("data")
        if not data:
            continue
        if mt.startswith("text/"):
            raw = b64url_decode(data).decode("utf-8", errors="replace")
            if mt == "text/html":
                soup = BeautifulSoup(raw, "html.parser")
                raw = soup.get_text("\n", strip=True)
            out.append(raw)
    return "\n".join(out)


def gmail_search(gmail, query, max_results=500):
    """List message IDs matching query (paginated)."""
    ids = []
    token = None
    while True:
        kw = dict(userId="me", q=query, maxResults=min(500, max_results - len(ids)))
        if token: kw["pageToken"] = token
        res = gmail.users().messages().list(**kw).execute()
        ids += [m["id"] for m in res.get("messages", [])]
        token = res.get("nextPageToken")
        if not token or len(ids) >= max_results:
            break
    return ids


def gmail_get(gmail, mid):
    return gmail.users().messages().get(userId="me", id=mid, format="full").execute()


def gmail_get_attachment(gmail, mid, aid):
    att = gmail.users().messages().attachments().get(userId="me", messageId=mid, id=aid).execute()
    return b64url_decode(att["data"])


def gmail_send(gmail, to_list, subject, html_body, text_body):
    """to_list: list of email addresses (first = To, rest = Cc)."""
    if isinstance(to_list, str):
        to_list = [to_list]
    msg = MIMEMultipart("alternative")
    msg["to"] = to_list[0]
    if len(to_list) > 1:
        msg["cc"] = ", ".join(to_list[1:])
    msg["subject"] = subject
    msg.attach(MIMEText(text_body, "plain", "utf-8"))
    msg.attach(MIMEText(html_body, "html", "utf-8"))
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    gmail.users().messages().send(userId="me", body={"raw": raw}).execute()


# ---------------- Telegram ----------------
def telegram_send(text, parse_mode="HTML"):
    """Send Telegram message. Returns message_id."""
    token = os.environ.get("TELEGRAM_TOKEN")
    chat_id = os.environ.get("TELEGRAM_CHAT_ID")
    if not token or not chat_id:
        print("  Telegram skipped (no TELEGRAM_TOKEN/CHAT_ID env)")
        return None
    r = requests.post(
        f"https://api.telegram.org/bot{token}/sendMessage",
        data={"chat_id": chat_id, "text": text, "parse_mode": parse_mode,
              "disable_web_page_preview": "true"},
        timeout=20,
    )
    r.raise_for_status()
    return r.json()["result"]["message_id"]


# ---------------- Date helpers ----------------
def previous_month(today=None):
    today = today or date.today()
    first = today.replace(day=1)
    last_prev = first - timedelta(days=1)
    return last_prev.year, last_prev.month


def month_range(y, m):
    first = date(y, m, 1)
    _, last_day = calendar.monthrange(y, m)
    return first, date(y, m, last_day)


def gmail_range_q(start, end):
    """Inclusive start, EXCLUSIVE end. Add 1 day to end for Gmail query."""
    return f"after:{start.strftime('%Y/%m/%d')} before:{(end + timedelta(days=1)).strftime('%Y/%m/%d')}"


# ---------------- Parsers ----------------
def parse_vcb_transfer(text, msg_date):
    """Parse 1 VCB Digibank chuyển khoản email.

    Body has bilingual labels: "Tên người hưởng\nBeneficiary Name\n<VALUE>".
    Anchor on English label (more reliable separator).
    """
    def grab(label):
        m = re.search(rf"{label}\s*\n+([^\n]+)", text)
        return m.group(1).strip() if m else ""

    amt_m = re.search(r"([\d,\.]+)\s*VND", text)
    if not amt_m: return None
    amount = int(amt_m.group(1).replace(",", "").replace(".", ""))

    time_m = re.search(r"(\d{2}:\d{2})\s+\S+\s+(\d{2}/\d{2}/\d{4})", text)
    if time_m:
        t = time_m.group(1)
        dd, mm, yy = time_m.group(2).split("/")
        iso_date = f"{yy}-{mm}-{dd}"
    else:
        t, iso_date = "", msg_date

    recipient = grab("Beneficiary Name")
    bank = grab("Beneficiary Bank")
    content = grab("Content")
    order_m = re.search(r"Order Number[^\d]*(\d+)", text) or re.search(r"(\d{10,14})", text)
    return {
        "source": "VCB_TRANSFER", "date": iso_date, "time": t,
        "amount_vnd": amount,
        "recipient": recipient, "bank": bank, "content": content,
        "order_number": order_m.group(1) if order_m else "",
    }


def parse_vcb_card(text, msg_date):
    """Parse 1 VCB card transaction email."""
    card_m = re.search(r"Visa\s+(\d{6})\D+(\d{4})", text)
    used_m = re.search(r"Sử dụng tại\s+At\s+([^\n]+)", text)
    amt_m = re.search(r"Transaction Amount\s+([\d,\.]+)\s+(\w+)", text) or \
            re.search(r"Số tiền[^\d]*([\d,\.]+)\s+(\w+)", text)
    time_m = re.search(r"(\d{2}:\d{2})\s+\S+\s+(\d{2}/\d{2}/\d{4})", text)
    if not amt_m: return None
    a = amt_m.group(1).replace(",", "")
    cur = amt_m.group(2)
    amount = float(a) if cur == "USD" else int(float(a))
    if time_m:
        t, d = time_m.group(1), time_m.group(2)
        dd, mm, yy = d.split("/")
        iso_date = f"{yy}-{mm}-{dd}"
    else:
        t, iso_date = "", msg_date
    return {
        "source": "VCB_CARD", "date": iso_date, "time": t,
        "amount": amount, "currency": cur,
        "card_last4": card_m.group(2) if card_m else "",
        "card_prefix": card_m.group(1) if card_m else "",
        "merchant_raw": (used_m.group(1) if used_m else "").strip(),
    }


def parse_tpbank_card(text, msg_date):
    """Parse 1 TPBank credit card transaction email."""
    card_m = re.search(r"Số thẻ:\s*(\d{6})\*+(\d{4})", text)
    if not card_m: return None
    # TPBank format: "Giá trị giao dịch:\n529,000VND" or "27.50USD"
    amt_m = re.search(r"Giá trị giao dịch:\s*([\d,\.]+)\s*(VND|USD|EUR|JPY)", text)
    merch_m = re.search(r"Nội dung giao dịch:\s*([^\n]+)", text)
    time_m = re.search(r"Thời gian thực hiện giao dịch:\s*(\d{2}/\d{2}/\d{4})\s+(\d{2}:\d{2})(?::\d{2})?", text)
    type_m = re.search(r"Loại giao dịch:\s*([^\n]+)", text)
    if not amt_m: return None
    a = amt_m.group(1)
    cur = amt_m.group(2)
    if cur == "USD":
        amount = float(a.replace(",", ""))
    else:
        # VND amount: 529,000 or 1,420,000 (comma thousands sep)
        amount = int(a.replace(",", "").replace(".", ""))
    if time_m:
        dd, mm, yy = time_m.group(1).split("/")
        iso_date = f"{yy}-{mm}-{dd}"
        t = time_m.group(2)
    else:
        t, iso_date = "", msg_date
    last4 = card_m.group(2)
    return {
        "source": "TPBANK_VISA" if last4 == "7606" else "TPBANK_JCB",
        "card_last4": last4, "card_prefix": card_m.group(1),
        "date": iso_date, "time": t,
        "amount": amount, "currency": cur,
        "merchant_raw": (merch_m.group(1) if merch_m else "").strip(),
        "txn_type": (type_m.group(1) if type_m else "").strip(),
    }


def parse_tpbank_tktt_pdf(pdf_bytes):
    """Extract giao dịch từ sao kê TPBank TKTT PDF."""
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f:
        f.write(pdf_bytes); tmp = f.name
    try:
        reader = pypdf.PdfReader(tmp)
        if reader.is_encrypted:
            # try common passwords
            for pw in ["07091992", "7091992", "19920907"]:
                if reader.decrypt(pw): break
        with pdfplumber.open(tmp) as pdf:
            tables = []
            for p in pdf.pages:
                for t in (p.extract_tables() or []):
                    tables.append(t)
    finally:
        os.unlink(tmp)

    rows = []
    for t in tables:
        for row in t:
            if not row: continue
            first = (row[0] or "").strip()
            if not re.match(r"^\d+$", first): continue
            if len(row) < 8: continue
            try:
                dd_mm_yy = (row[1] or "").strip().split("/")
                iso = f"{dd_mm_yy[2]}-{dd_mm_yy[1]}-{dd_mm_yy[0]}" if len(dd_mm_yy) == 3 else ""
                debit = float((row[5] or "0").replace(",", ""))
                if debit <= 0: continue
                descr = (row[4] or "").replace("\n", " ").strip()
                is_cc = "TT QUA TPBANK EBANKING" in descr.upper()
                rows.append({
                    "source": "TPBANK_TKTT", "date": iso, "time": "",
                    "amount_vnd": int(debit),
                    "merchant_raw": descr,
                    "txn_id": (row[3] or "").strip(),
                    "exclude_from_spending": is_cc,
                    "exclude_reason": "Trả dư nợ thẻ tín dụng" if is_cc else "",
                })
            except (ValueError, IndexError):
                continue
    return rows


# ---------------- Categorization ----------------
RULES = [
    ("Coffee",              [r"viva ?star", r"bits ?coffee", r"dan vinh", r"highland", r"coffee"]),
    ("AI / Tools (work)",   [r"anthropic", r"claude\.ai", r"perplexity", r"kanbanflow", r"google one", r"openai", r"github\.com/sponsors"]),
    ("Project cá nhân (Cloud/Hosting)", [r"cloudzone", r"tryopenclaw", r"openclaw", r"digitalocean", r"vultr", r"aws", r"hetzner"]),
    ("Subscription",        [r"youtube", r"spotify", r"netflix"]),
    ("Internet/FPT",        [r"fpt"]),
    ("Phí nhà / chung cư",  [r"vinhomes", r"phí dịch vụ", r"nguyen duc dung"]),
    ("Vé giải trí / Du lịch", [r"airbnb", r"nam cat tien", r"ba den", r"leo nui", r"vincom retail", r"vé máy bay", r"vietjet", r"vietnamair", r"agoda", r"booking\.com"]),
    ("Đồ gia dụng / Điện tử", [r"imac", r"macbook", r"may loc khong khi", r"the gioi di dong", r"fpt shop", r"phong vu"]),
    ("Y tế / Kính mắt",     [r"pharmacity", r"mat hai yen", r"benh vien", r"phong kham", r"long chau", r"megane", r"matkinh", r"optic"]),
    ("Siêu thị / Tạp hoá",  [r"king food", r"kfm_", r"kfmkp", r"co-opxtra", r"saigon co-op", r"ministop", r"fairprice", r"co\.op", r"winmart", r"bach hoa"]),
    ("Quần áo / Thể thao",  [r"uniqlo", r"decathlon", r"hdl", r"zara", r"h&m"]),
    ("Di chuyển (Grab/Xanh)", [r"grab", r"xanhsm", r"xanh sm", r"green sm", r".op\*"]),
    ("Mua sắm online",      [r"shopee", r"etsy", r"lazada", r"tiki"]),
    ("Ăn uống ngoài",       [r"quan an", r"hu tieu", r"^pho", r" pho ", r"co loc da lat", r"cot dien", r"han yen", r"nha hang", r"restaurant"]),
    ("Xăng / Cao tốc",      [r"plx", r"petrol", r"vetc"]),
    ("Chuyển khoản (Zalo)", [r"zalo"]),
    ("Chuyển QR / Thanh toán", [r"chuyen tien qr", r"thanh toan qr", r"chuyen tien nhanh qua qr", r"vqrloamb", r"kovqr"]),
]


def categorize(text):
    t = (text or "").lower()
    for cat, patterns in RULES:
        for p in patterns:
            if re.search(p, t):
                return cat
    return "Khác / Chuyển khoản cá nhân"


# ---------------- Pipeline ----------------
def normalize_row(rec, usd_vnd=DEFAULT_USD_VND):
    """Convert to standardized row format."""
    vnd = rec.get("amount_vnd")
    if vnd is None:
        cur = rec.get("currency", "VND")
        amt = rec.get("amount", 0)
        vnd = int(round(amt * usd_vnd)) if cur == "USD" else int(amt)
    merch = rec.get("merchant_raw") or rec.get("recipient") or ""
    return {
        "date": rec["date"], "time": rec.get("time", ""),
        "source": rec["source"],
        "category": categorize(merch),
        "amount_vnd": vnd,
        "amount_orig": rec.get("amount", rec.get("amount_vnd")),
        "currency": rec.get("currency", "VND"),
        "merchant": merch,
        "note": rec.get("content", "") or rec.get("txn_id", ""),
        "exclude_from_spending": rec.get("exclude_from_spending", False),
        "exclude_reason": rec.get("exclude_reason", ""),
        "raw": rec,
    }


def fetch_month(gmail, year, month, overrides):
    """Fetch all transactions in given month."""
    start, end = month_range(year, month)
    print(f"== Fetching {year}-{month:02d} ({start} → {end}) ==")
    rng = gmail_range_q(start, end)
    sources = []

    # 1. VCB transfers
    q = f"from:VCBDigibank@info.vietcombank.com.vn {rng}"
    ids = gmail_search(gmail, q)
    print(f"  VCB transfers: {len(ids)} emails")
    for mid in ids:
        msg = gmail_get(gmail, mid)
        txt = get_text(msg.get("payload", {}))
        rec = parse_vcb_transfer(txt, f"{year}-{month:02d}-01")
        if rec: sources.append(rec)

    # 2. VCB card
    q = f'from:info@info.vietcombank.com.vn subject:"Thông báo giao dịch thẻ" {rng}'
    ids = gmail_search(gmail, q)
    print(f"  VCB card: {len(ids)} emails")
    for mid in ids:
        msg = gmail_get(gmail, mid)
        txt = get_text(msg.get("payload", {}))
        rec = parse_vcb_card(txt, f"{year}-{month:02d}-01")
        if rec: sources.append(rec)

    # 3. TPBank credit cards (both 7606 and 8272)
    q = f'from:tpbank@tpb.com.vn subject:"giao dịch thẻ tín dụng" {rng}'
    ids = gmail_search(gmail, q)
    print(f"  TPBank credit cards: {len(ids)} emails")
    for mid in ids:
        msg = gmail_get(gmail, mid)
        txt = get_text(msg.get("payload", {}))
        rec = parse_tpbank_card(txt, f"{year}-{month:02d}-01")
        if rec: sources.append(rec)

    # 4. TPBank TKTT PDF (gửi mùng 1-5 tháng sau)
    next_start = (end + timedelta(days=1))
    next_end = next_start + timedelta(days=15)
    pdf_q = f'from:tpbank@tpb.com.vn subject:"sao kê tài khoản" after:{next_start.strftime("%Y/%m/%d")} before:{next_end.strftime("%Y/%m/%d")}'
    pdf_ids = gmail_search(gmail, pdf_q)
    print(f"  TPBank TKTT statement: {len(pdf_ids)} email(s)")
    for mid in pdf_ids:
        msg = gmail_get(gmail, mid)
        for part in walk_parts(msg.get("payload", {})):
            if part.get("filename", "").lower().endswith(".pdf"):
                aid = part["body"]["attachmentId"]
                pdf_bytes = gmail_get_attachment(gmail, mid, aid)
                tktt_rows = parse_tpbank_tktt_pdf(pdf_bytes)
                # filter to month
                tktt_rows = [r for r in tktt_rows if r["date"].startswith(f"{year}-{month:02d}")]
                print(f"    Extracted {len(tktt_rows)} TKTT rows")
                sources.extend(tktt_rows)
                break

    # Normalize + apply overrides
    rows = [normalize_row(r) for r in sources]
    for r in rows:
        tid = r["raw"].get("txn_id")
        if tid and tid in overrides:
            ov = overrides[tid]
            if ov.get("category"): r["category"] = ov["category"]
            if ov.get("user_note"): r["note"] = f"{r['note']} | {ov['user_note']}"
            if ov.get("exclude"):
                r["exclude_from_spending"] = True
                r["exclude_reason"] = ov.get("user_note", "user override")

    rows.sort(key=lambda r: (r["date"], r["time"]))
    return rows


# ---------------- Sheet upload ----------------
def upload_month_tab(sheets, sheet_id, rows, year, month):
    tab_name = f"Raw {year}-{month:02d}"
    # Create tab if not exist
    meta = sheets.spreadsheets().get(spreadsheetId=sheet_id).execute()
    existing = {s["properties"]["title"]: s["properties"]["sheetId"] for s in meta["sheets"]}
    if tab_name not in existing:
        sheets.spreadsheets().batchUpdate(spreadsheetId=sheet_id, body={
            "requests": [{"addSheet": {"properties": {"title": tab_name}}}]
        }).execute()
        meta = sheets.spreadsheets().get(spreadsheetId=sheet_id).execute()
        existing = {s["properties"]["title"]: s["properties"]["sheetId"] for s in meta["sheets"]}

    # Build values
    vals = [["Date","Time","Source","Category","Amount VND","Amount Orig","Currency","Merchant","Note","Excluded?","Reason"]]
    for r in rows:
        vals.append([r["date"], r["time"], r["source"], r["category"],
                     r["amount_vnd"], r.get("amount_orig",""), r.get("currency","VND"),
                     r["merchant"], r["note"],
                     "EXCL" if r["exclude_from_spending"] else "",
                     r.get("exclude_reason","")])

    sheets.spreadsheets().values().clear(spreadsheetId=sheet_id, range=f"{tab_name}!A:Z").execute()
    sheets.spreadsheets().values().update(
        spreadsheetId=sheet_id, range=f"{tab_name}!A1",
        valueInputOption="USER_ENTERED", body={"values": vals}
    ).execute()

    # Format
    gid = existing[tab_name]
    sheets.spreadsheets().batchUpdate(spreadsheetId=sheet_id, body={"requests": [
        {"updateSheetProperties": {"properties":{"sheetId":gid,"gridProperties":{"frozenRowCount":1}},"fields":"gridProperties.frozenRowCount"}},
        {"repeatCell": {"range":{"sheetId":gid,"startRowIndex":0,"endRowIndex":1},
                        "cell":{"userEnteredFormat":{"backgroundColor":{"red":0.85,"green":0.9,"blue":1.0},"textFormat":{"bold":True}}},
                        "fields":"userEnteredFormat(backgroundColor,textFormat)"}},
        {"autoResizeDimensions":{"dimensions":{"sheetId":gid,"dimension":"COLUMNS","startIndex":0,"endIndex":11}}},
        {"repeatCell":{"range":{"sheetId":gid,"startRowIndex":1,"startColumnIndex":4,"endColumnIndex":5},
                       "cell":{"userEnteredFormat":{"numberFormat":{"type":"NUMBER","pattern":"#,##0"}}},
                       "fields":"userEnteredFormat.numberFormat"}},
    ]}).execute()
    print(f"  Uploaded → tab {tab_name}")
    return tab_name


def write_summary_tab(sheets, sheet_id, year, month, rows):
    """Write/replace Summary YYYY-MM tab."""
    tab = f"Summary {year}-{month:02d}"
    spending = [r for r in rows if not r["exclude_from_spending"]]
    excluded = [r for r in rows if r["exclude_from_spending"]]
    total = sum(r["amount_vnd"] for r in spending)
    excl_total = sum(r["amount_vnd"] for r in excluded)

    cat = defaultdict(lambda: {"count":0,"total":0})
    for r in spending:
        c = cat[r["category"]]; c["count"] += 1; c["total"] += r["amount_vnd"]
    cat_rows = sorted([(k,v["count"],v["total"]) for k,v in cat.items()], key=lambda x: -x[2])

    src = defaultdict(lambda: {"count":0,"total":0})
    for r in spending:
        s = src[r["source"]]; s["count"] += 1; s["total"] += r["amount_vnd"]

    vals = [
        [f"CHI TIÊU THÁNG {month:02d}/{year}", "", "", ""],
        ["", "", "", ""],
        ["TỔNG (đã loại double-count)", f"{total:,} VND", "", ""],
        ["Số giao dịch", str(len(spending)), "", ""],
        ["Trung bình / ngày", f"{total // calendar.monthrange(year, month)[1]:,} VND", "", ""],
        ["", "", "", ""],
        ["LOẠI TRỪ", f"{excl_total:,} VND", f"{len(excluded)} GD", "(trả nợ CC / nội bộ)"],
        ["", "", "", ""],
        ["BY SOURCE", "Số GD", "Total VND", "%"],
    ]
    for s, v in sorted(src.items(), key=lambda x: -x[1]["total"]):
        vals.append([s, v["count"], v["total"], f"{v['total']/total*100:.1f}%" if total else "0%"])
    vals.append(["", "", "", ""])
    vals.append(["BY CATEGORY", "Số GD", "Total VND", "%"])
    for c, n, t in cat_rows:
        vals.append([c, n, t, f"{t/total*100:.1f}%" if total else "0%"])

    meta = sheets.spreadsheets().get(spreadsheetId=sheet_id).execute()
    existing = {s["properties"]["title"]: s["properties"]["sheetId"] for s in meta["sheets"]}
    if tab not in existing:
        sheets.spreadsheets().batchUpdate(spreadsheetId=sheet_id, body={
            "requests": [{"addSheet": {"properties": {"title": tab}}}]
        }).execute()
    sheets.spreadsheets().values().clear(spreadsheetId=sheet_id, range=f"{tab}!A:Z").execute()
    sheets.spreadsheets().values().update(
        spreadsheetId=sheet_id, range=f"{tab}!A1",
        valueInputOption="USER_ENTERED", body={"values": vals}
    ).execute()
    print(f"  Summary tab updated: {tab}")
    return tab, total, cat_rows, src


def write_pending_tab(sheets, sheet_id, year, month, ambiguous_rows, tg_msg_id=None):
    """Write tab `Pending YYYY-MM` with ambiguous transactions."""
    tab = f"Pending {year}-{month:02d}"
    vals = [["#", "Date", "Source", "Amount VND", "Merchant", "Note", "txn_id", "Current Category", "Status"]]
    for i, r in enumerate(ambiguous_rows, 1):
        vals.append([i, r["date"], r["source"], r["amount_vnd"], r["merchant"], r["note"],
                     r["raw"].get("txn_id", ""), r["category"], "PENDING"])
    if tg_msg_id:
        vals.append([])
        vals.append(["telegram_msg_id", tg_msg_id])
    vals.append([])
    vals.append(["timestamp", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])

    meta = sheets.spreadsheets().get(spreadsheetId=sheet_id).execute()
    existing = {s["properties"]["title"]: s["properties"]["sheetId"] for s in meta["sheets"]}
    if tab not in existing:
        sheets.spreadsheets().batchUpdate(spreadsheetId=sheet_id, body={
            "requests": [{"addSheet": {"properties": {"title": tab}}}]
        }).execute()
    sheets.spreadsheets().values().clear(spreadsheetId=sheet_id, range=f"{tab}!A:Z").execute()
    sheets.spreadsheets().values().update(
        spreadsheetId=sheet_id, range=f"{tab}!A1",
        valueInputOption="USER_ENTERED", body={"values": vals}
    ).execute()
    print(f"  Pending tab updated: {tab} ({len(ambiguous_rows)} ambiguous)")
    return tab


def detect_ambiguous(rows):
    """Return rows with category==Khác AND amount >= AMBIGUOUS_MIN_VND, not excluded."""
    return sorted(
        [r for r in rows
         if r["category"] == AMBIGUOUS_CATEGORY
         and r["amount_vnd"] >= AMBIGUOUS_MIN_VND
         and not r["exclude_from_spending"]],
        key=lambda r: -r["amount_vnd"]
    )


def write_quarter_tab(sheets, sheet_id, year, quarter, all_rows_by_month):
    """all_rows_by_month: {(y,m): [rows]} for 3 months of quarter."""
    tab = f"Quarter Q{quarter}-{year}"
    spending_all = []
    for (y,m), rows in sorted(all_rows_by_month.items()):
        spending_all.extend([r for r in rows if not r["exclude_from_spending"]])
    total = sum(r["amount_vnd"] for r in spending_all)

    cat = defaultdict(lambda: {"count":0,"total":0})
    for r in spending_all:
        c = cat[r["category"]]; c["count"] += 1; c["total"] += r["amount_vnd"]
    cat_rows = sorted([(k,v["count"],v["total"]) for k,v in cat.items()], key=lambda x: -x[2])

    monthly_totals = {}
    for (y,m), rows in sorted(all_rows_by_month.items()):
        sp = [r for r in rows if not r["exclude_from_spending"]]
        monthly_totals[(y,m)] = sum(r["amount_vnd"] for r in sp)

    vals = [
        [f"QUARTER Q{quarter}/{year}", "", "", ""],
        ["", "", "", ""],
        ["TỔNG Q", f"{total:,} VND", "", ""],
        ["Số giao dịch", str(len(spending_all)), "", ""],
        ["Trung bình / tháng", f"{total // 3:,} VND", "", ""],
        ["", "", "", ""],
        ["BY MONTH", "Total VND", "%", ""],
    ]
    for (y,m), t in monthly_totals.items():
        vals.append([f"{y}-{m:02d}", t, f"{t/total*100:.1f}%" if total else "0%", ""])
    vals.append(["", "", "", ""])
    vals.append(["BY CATEGORY (Q tổng)", "Số GD", "Total VND", "%"])
    for c, n, t in cat_rows:
        vals.append([c, n, t, f"{t/total*100:.1f}%" if total else "0%"])

    meta = sheets.spreadsheets().get(spreadsheetId=sheet_id).execute()
    existing = {s["properties"]["title"]: s["properties"]["sheetId"] for s in meta["sheets"]}
    if tab not in existing:
        sheets.spreadsheets().batchUpdate(spreadsheetId=sheet_id, body={
            "requests": [{"addSheet": {"properties": {"title": tab}}}]
        }).execute()
    sheets.spreadsheets().values().clear(spreadsheetId=sheet_id, range=f"{tab}!A:Z").execute()
    sheets.spreadsheets().values().update(
        spreadsheetId=sheet_id, range=f"{tab}!A1",
        valueInputOption="USER_ENTERED", body={"values": vals}
    ).execute()
    print(f"  Quarter tab: {tab}, total={total:,}")
    return tab, total, cat_rows, monthly_totals


# ---------------- Email ----------------
def build_email_html(year, month, total, cat_rows, src, sheet_url, quarter_info=None,
                     ambiguous_rows=None, mode="final"):
    cat_html = "\n".join(
        f"<tr><td>{c}</td><td align=right>{n}</td><td align=right>{t:,}</td><td align=right>{t/total*100:.1f}%</td></tr>"
        for c, n, t in cat_rows[:15]
    )
    src_html = "\n".join(
        f"<tr><td>{s}</td><td align=right>{v['count']}</td><td align=right>{v['total']:,}</td></tr>"
        for s, v in sorted(src.items(), key=lambda x: -x[1]['total'])
    )
    extra = ""
    if quarter_info:
        qtab, qtotal, qcat, monthly = quarter_info
        m_html = "\n".join(f"<tr><td>{y}-{m:02d}</td><td align=right>{t:,}</td></tr>" for (y,m),t in monthly.items())
        qc_html = "\n".join(f"<tr><td>{c}</td><td align=right>{n}</td><td align=right>{t:,}</td></tr>" for c,n,t in qcat[:10])
        extra = f"""
        <h2>📊 QUARTERLY REPORT</h2>
        <p><b>Total quarter: {qtotal:,} VND</b> | Avg/month: {qtotal//3:,} VND</p>
        <h3>By Month</h3>
        <table border=1 cellpadding=5><tr><th>Month</th><th>Total</th></tr>{m_html}</table>
        <h3>Top 10 Category (Quarter)</h3>
        <table border=1 cellpadding=5><tr><th>Category</th><th>#</th><th>Total</th></tr>{qc_html}</table>
        """
    amb_html = ""
    if ambiguous_rows:
        amb_rows_html = "\n".join(
            f"<tr><td>{i}</td><td>{r['date']}</td><td align=right>{r['amount_vnd']:,}</td><td>{r['merchant']}</td></tr>"
            for i, r in enumerate(ambiguous_rows, 1)
        )
        amb_html = f"""
        <div style="background:#fff8e1;padding:15px;border-left:4px solid #ff9800;margin:15px 0">
        <h3>⏳ {len(ambiguous_rows)} giao dịch cần clarify</h3>
        <p>Tao đã gửi Telegram hỏi chi tiết. Mày reply Telegram → tao sẽ update sheet + gửi mail FINAL với category đúng.</p>
        <table border=1 cellpadding=5><tr><th>#</th><th>Date</th><th>Amount VND</th><th>Merchant</th></tr>{amb_rows_html}</table>
        </div>
        """
    title = "💸 Báo cáo chi tiêu (PRELIMINARY)" if mode == "preliminary" else "💸 Báo cáo chi tiêu (FINAL)"
    return f"""<html><body style="font-family:Arial,sans-serif">
    <h1>{title} — {month:02d}/{year}</h1>
    <p><b>Tổng chi tiêu: {total:,} VND</b></p>
    <p>📊 <a href="{sheet_url}">Mở Google Sheet đầy đủ</a></p>
    {amb_html}

    <h2>Phân bổ theo nguồn</h2>
    <table border=1 cellpadding=5><tr><th>Nguồn</th><th>Số GD</th><th>Total VND</th></tr>{src_html}</table>

    <h2>Top 15 danh mục</h2>
    <table border=1 cellpadding=5><tr><th>Category</th><th>#</th><th>Total VND</th><th>%</th></tr>{cat_html}</table>

    {extra}

    <hr>
    <p><small>Auto-generated by spending_pipeline.py • {datetime.now().strftime('%Y-%m-%d %H:%M')}</small></p>
    </body></html>"""


# ---------------- Main ----------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--month", help="YYYY-MM (default: previous month)")
    ap.add_argument("--mode", choices=["preliminary", "final"], default="preliminary",
                    help="preliminary = first run, detect ambiguous + Telegram ask + PRELIM email; final = after user reply, send FINAL email")
    ap.add_argument("--include-quarter", action="store_true", help="Force quarterly report")
    ap.add_argument("--auto-quarter", action="store_true", help="Auto-detect end-of-quarter")
    ap.add_argument("--no-email", action="store_true", help="Skip sending email")
    ap.add_argument("--no-telegram", action="store_true", help="Skip Telegram")
    ap.add_argument("--overrides", help="Path to overrides JSON file")
    args = ap.parse_args()

    # Determine month
    if args.month:
        y, m = map(int, args.month.split("-"))
    else:
        y, m = previous_month()
    print(f"Target month: {y}-{m:02d}")

    creds = get_creds()
    gmail = svc("gmail", "v1", creds)
    sheets = svc("sheets", "v4", creds)

    sheet_id = os.environ.get("SHEET_ID", "1tnwApdgGLF1qG_b_lIclO78hJHGo6ZUtXmpnKYrKb8A")
    notify_emails_str = os.environ.get("NOTIFY_EMAILS", "hung216mc@gmail.com,duongngoccongkhanh@gmail.com")
    notify_emails = [e.strip() for e in notify_emails_str.split(",") if e.strip()]
    overrides = {}
    # Layer 1: default file in repo
    if (HERE / "overrides.json").exists():
        overrides.update(json.loads((HERE / "overrides.json").read_text(encoding="utf-8")))
    # Layer 2: runtime additions (CLI > env)
    if args.overrides:
        overrides.update(json.loads(Path(args.overrides).read_text(encoding="utf-8")))
    elif os.environ.get("OVERRIDES_JSON"):
        overrides.update(json.loads(os.environ["OVERRIDES_JSON"]))

    # Fetch + upload current month
    rows = fetch_month(gmail, y, m, overrides)
    upload_month_tab(sheets, sheet_id, rows, y, m)
    _, total, cat_rows, src = write_summary_tab(sheets, sheet_id, y, m, rows)

    # Quarterly logic
    quarter_info = None
    is_quarter_start = (m in [12, 3, 6, 9])  # if reporting Mar/Jun/Sep/Dec = end of quarter
    if args.include_quarter or (args.auto_quarter and is_quarter_start):
        # Find quarter of target month
        q = (m - 1) // 3 + 1
        q_first_m = (q - 1) * 3 + 1
        all_q = {}
        for qm in range(q_first_m, q_first_m + 3):
            print(f"== Fetching quarter month {y}-{qm:02d} ==")
            if (y, qm) == (y, m):
                all_q[(y, qm)] = rows  # reuse
            else:
                all_q[(y, qm)] = fetch_month(gmail, y, qm, overrides)
                upload_month_tab(sheets, sheet_id, all_q[(y, qm)], y, qm)
                write_summary_tab(sheets, sheet_id, y, qm, all_q[(y, qm)])
        quarter_info = write_quarter_tab(sheets, sheet_id, y, q, all_q)

    # Detect ambiguous (only spending rows where category is "Khác" and amount large)
    ambiguous = detect_ambiguous(rows)
    if ambiguous:
        print(f"\nAmbiguous ({len(ambiguous)}):")
        for r in ambiguous:
            print(f"  {r['date']} | {r['amount_vnd']:>12,} | {r['merchant'][:60]} | src={r['source']}")
    sheet_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit"

    # Preliminary mode + has ambiguous → Telegram + Pending tab
    tg_msg_id = None
    if args.mode == "preliminary" and ambiguous and not args.no_telegram:
        lines = [f"💸 <b>Chi tiêu {m:02d}/{y}: {total:,} VND</b>",
                 f"📊 <a href='{sheet_url}'>Sheet</a>",
                 "",
                 f"⏳ <b>{len(ambiguous)} giao dịch cần clarify</b> (Khác + ≥2M):",
                 ""]
        for i, r in enumerate(ambiguous, 1):
            mer = r["merchant"][:50]
            lines.append(f"<b>{i}.</b> {r['date']} — <b>{r['amount_vnd']:,}</b> VND")
            lines.append(f"   <code>{mer}</code>")
        lines.append("")
        lines.append("👉 <i>Reply tin nhắn này, tao tự hiểu (vd: '1 tiền nhà, 2 cho mẹ đi SG'). Sau khi rep, mai/2 hôm sau tao auto update email FINAL.</i>")
        tg_msg_id = telegram_send("\n".join(lines))
        print(f"Telegram sent (msg_id={tg_msg_id})")
        write_pending_tab(sheets, sheet_id, y, m, ambiguous, tg_msg_id)
    elif args.mode == "final":
        # Clear pending tab status
        try:
            tab = f"Pending {y}-{m:02d}"
            sheets.spreadsheets().values().clear(spreadsheetId=sheet_id, range=f"{tab}!A:Z").execute()
            sheets.spreadsheets().values().update(
                spreadsheetId=sheet_id, range=f"{tab}!A1", valueInputOption="USER_ENTERED",
                body={"values": [["RESOLVED", datetime.now().strftime("%Y-%m-%d %H:%M:%S")]]}
            ).execute()
            print(f"Cleared pending tab")
        except Exception as e:
            print(f"  (no pending tab to clear: {e})")

    # Email
    if not args.no_email:
        mode_tag = "PRELIM" if args.mode == "preliminary" and ambiguous else "FINAL"
        subject = f"💸 [{mode_tag}] Chi tiêu {m:02d}/{y}: {total:,} VND"
        if quarter_info:
            subject += f" + Q{(m-1)//3+1}"
        html = build_email_html(y, m, total, cat_rows, src, sheet_url, quarter_info,
                                ambiguous_rows=ambiguous if mode_tag == "PRELIM" else None,
                                mode="preliminary" if mode_tag == "PRELIM" else "final")
        text = f"[{mode_tag}] Chi tiêu {m:02d}/{y}: {total:,} VND\nSheet: {sheet_url}"
        if mode_tag == "PRELIM":
            text += f"\n\nCòn {len(ambiguous)} GD chờ clarify. Check Telegram."
        gmail_send(gmail, notify_emails, subject, html, text)
        print(f"Email sent → {notify_emails}")

    print(f"\n✅ DONE [{args.mode}] — Total: {total:,} VND, Ambiguous: {len(ambiguous)}\n   Sheet: {sheet_url}")


if __name__ == "__main__":
    main()
