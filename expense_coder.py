"""
Expense Coder - free tool that codes bank payments to expense accounts.

Inputs (put them in the "input" folder):
  input/trial_balance.csv  (or .xlsx)  - your trial balance / chart of accounts
  input/bank/*.csv                     - bank statement CSVs, one file per bank account
  input/bills/*.jpg|png|pdf            - scanned expense bills

Commands:
  py expense_coder.py setup   -> reads the trial balance and creates rules.xlsx for you to check
  py expense_coder.py run     -> reads bank CSVs + bills, writes output/coded_expenses.xlsx
  py expense_coder.py learn   -> reads your corrections in output/coded_expenses.xlsx and
                                 saves them as vendor rules for next time
"""
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from dateutil import parser as dateparser
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent
INPUT = BASE / "input"
BANK_DIR = INPUT / "bank"
BILL_DIR = INPUT / "bills"
OUTPUT = BASE / "output"
RULES = BASE / "rules.xlsx"
RESULT = OUTPUT / "coded_expenses.xlsx"

# --------------------------------------------------------------------------------------
# Built-in keyword lists. If an account NAME contains one of the trigger words, the
# account gets these keywords pre-filled in rules.xlsx. You can edit them there freely.
# --------------------------------------------------------------------------------------
KEYWORD_LIBRARY = [
    (["travel", "subsistence"], "travel, taxi, uber, bolt, train, rail, trainline, flight, airline, airways, easyjet, ryanair, hotel, premier inn, travelodge, airbnb, bus, coach, tube, tfl, parking, toll"),
    (["motor", "vehicle", "car "], "fuel, petrol, diesel, shell, bp, esso, texaco, parking, mot, tyre, tyres, garage, car wash"),
    (["entertain", "hospitality"], "restaurant, cafe, coffee, bar, pub, dinner, lunch, meal, food, catering, starbucks, costa, pret, nando, pizza, wine"),
    (["building", "premises", "repair", "maintenance", "property"], "plumber, plumbing, electrician, electrical, repair, maintenance, locksmith, paint, b&q, screwfix, toolstation, wickes, roof, boiler, builder, glazing"),
    (["computer", "it ", "software", "technology"], "computer, laptop, pc, software, microsoft, adobe, apple, dell, hp, lenovo, it support, printer, ink, toner, aws, hosting, domain, google workspace, antivirus, server, network"),
    (["telephone", "phone", "internet", "communication"], "vodafone, ee, o2, three, bt, sky, virgin media, broadband, mobile, phone, telephone, sim"),
    (["stationery", "office", "printing"], "staples, ryman, stationery, paper, pens, envelopes, office supplies, amazon, viking, printing"),
    (["postage", "courier", "carriage"], "royal mail, post office, postage, stamps, dhl, fedex, ups, dpd, hermes, evri, courier"),
    (["light", "heat", "utilit", "electric", "power"], "electricity, gas, water, british gas, edf, octopus, eon, e.on, ovo, scottish power, thames water"),
    (["rent", "lease"], "rent, lease, landlord"),
    (["insurance"], "insurance, aviva, axa, hiscox, zurich, direct line, premium"),
    (["bank charge", "bank fee", "charges"], "bank charge, service charge, fee, commission, overdraft"),
    (["advertis", "marketing", "promotion"], "advertising, facebook, meta, google ads, linkedin, marketing, flyers, promotion"),
    (["subscription"], "subscription, membership, netflix, spotify, dropbox, zoom, slack"),
    (["training", "course"], "training, course, seminar, conference, workshop, exam"),
    (["professional", "legal", "accountan", "consult"], "solicitor, legal, accountant, accountancy, consultancy, consultant, audit"),
    (["cleaning"], "cleaning, cleaner, janitorial, hygiene"),
    (["wage", "salar", "payroll"], "salary, salaries, payroll, wages"),
]
BANK_NAME_HINTS = ["bank", "current a", "current account", "savings", "barclays", "hsbc", "lloyds",
                   "natwest", "santander", "monzo", "starling", "revolut", "metro bank", "cash at",
                   "petty cash", "credit card"]


# ======================================================================================
# Small helpers
# ======================================================================================
def say(msg):
    print(msg, flush=True)


def norm(s):
    return re.sub(r"\s+", " ", str(s or "")).strip().lower()


def to_amount(v):
    """'£1,234.50' / '(12.00)' / '12.00 DR' -> float; blank -> None"""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if not s:
        return None
    neg = s.startswith("(") and s.endswith(")") or s.startswith("-") or s.upper().endswith("DR")
    s = re.sub(r"[^\d.]", "", s)
    if not s or s == ".":
        return None
    try:
        n = float(s)
    except ValueError:
        return None
    return -n if neg else n


def to_date(v, dayfirst):
    if v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    try:
        return dateparser.parse(str(v), dayfirst=dayfirst, fuzzy=True).date()
    except (ValueError, OverflowError):
        return None


def find_col(columns, *words):
    for w in words:
        for c in columns:
            if w in norm(c):
                return c
    return None


def read_table(path, header_word="date"):
    """Read CSV/XLSX, skipping any title lines above the real header row."""
    if path.suffix.lower() in (".xlsx", ".xls"):
        raw = pd.read_excel(path, header=None, dtype=str)
    else:
        import csv
        text = None
        for enc in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                text = path.read_text(encoding=enc)
                break
            except UnicodeDecodeError:
                continue
        lines = [r for r in csv.reader(text.splitlines()) if any(x.strip() for x in r)]
        width = max((len(r) for r in lines), default=0)
        raw = pd.DataFrame([r + [""] * (width - len(r)) for r in lines], dtype=str)
    raw = raw.fillna("")
    hdr = 0
    for i in range(min(25, len(raw))):
        cells = [norm(x) for x in raw.iloc[i].tolist()]
        if any(header_word in c for c in cells) and sum(1 for c in cells if c) >= 2:
            hdr = i
            break
    df = raw.iloc[hdr + 1:].copy()
    df.columns = [str(c).strip() or f"col{j}" for j, c in enumerate(raw.iloc[hdr].tolist())]
    return df.reset_index(drop=True)


# ======================================================================================
# SETUP: trial balance -> rules.xlsx
# ======================================================================================
def guess_type(code, name):
    n = norm(name)
    try:
        c = int(re.sub(r"\D", "", str(code)) or -1)
    except ValueError:
        c = -1
    if 0 <= c <= 999 or "suspense" in n:   # fixed assets (Sage UK) / suspense
        return "Other"
    looks_like_cost = any(w in n for w in ("charge", "fee", "interest", "commission"))
    if any(h in n for h in BANK_NAME_HINTS) and not looks_like_cost:
        return "Bank"
    if 1200 <= c <= 1299:            # Sage UK default bank range
        return "Bank"
    if 5000 <= c <= 9998:            # Sage UK default purchases/overheads range
        return "Expense"
    if any(t in n + " " for triggers, _ in KEYWORD_LIBRARY for t in triggers):
        return "Expense"
    return "Other"


def default_keywords(name):
    n = norm(name) + " "
    kws = []
    for triggers, words in KEYWORD_LIBRARY:
        if any(t in n for t in triggers):
            kws.append(words)
    return ", ".join(kws)


def setup():
    tb = next((p for p in INPUT.glob("trial_balance.*")), None)
    if not tb:
        sys.exit("Put your trial balance in input/trial_balance.csv (or .xlsx) first.")
    df = read_table(tb, header_word="code")
    if df.empty:
        df = read_table(tb, header_word="account")
    code_col = find_col(df.columns, "code", "ref", "n/c", "nominal", "account no", "number")
    name_col = find_col(df.columns, "name", "description", "account")
    if code_col == name_col:
        name_col = next(c for c in df.columns if c != code_col)
    say(f"Trial balance: using '{code_col}' as code and '{name_col}' as name")

    wb = Workbook()
    ws = wb.active
    ws.title = "Accounts"
    ws.append(["Code", "Name", "Type (Expense/Bank/Other)", "Keywords (comma separated)"])
    for _, r in df.iterrows():
        code, name = str(r[code_col]).strip(), str(r[name_col]).strip()
        if not code or not name or norm(code) in ("total", "totals"):
            continue
        t = guess_type(code, name)
        ws.append([code, name, t, default_keywords(name) if t == "Expense" else ""])

    wb.create_sheet("Banks").append(["Bank account code", "Text in CSV file name"])
    for row in list(ws.iter_rows(min_row=2, values_only=True)):
        if row[2] == "Bank":
            wb["Banks"].append([row[0], row[0]])

    wb.create_sheet("Vendors").append(["If text contains", "Use expense code"])

    st = wb.create_sheet("Settings")
    for row in [["Setting", "Value", "Meaning"],
                ["DayFirst", "Yes", "Yes = dates are DD/MM/YYYY, No = MM/DD/YYYY"],
                ["DateWindowDays", 10, "How many days apart a bill and bank payment may be"],
                ["SuspenseCode", "9999", "Code used when the tool cannot decide (review these)"],
                ["DefaultTaxCode", "T9", "Tax code put in the Sage import (T9 = outside VAT in Sage UK)"]]:
        st.append(row)
    style(wb)
    wb.save(RULES)
    say(f"\nCreated {RULES.name}. Open it and check:\n"
        "  - Accounts sheet: Type column (Expense / Bank / Other) and Keywords\n"
        "  - Banks sheet: which text in each bank CSV file name identifies that bank\n"
        "Then run:  py expense_coder.py run")


# ======================================================================================
# Load rules
# ======================================================================================
def load_rules():
    if not RULES.exists():
        sys.exit("rules.xlsx not found - run:  py expense_coder.py setup")
    wb = load_workbook(RULES, data_only=True)
    accounts, banks, vendors, settings = {}, [], [], {}
    for code, name, typ, kw in wb["Accounts"].iter_rows(min_row=2, max_col=4, values_only=True):
        if code is None:
            continue
        accounts[str(code).strip()] = {
            "name": str(name or ""), "type": norm(typ),
            "keywords": [k.strip().lower() for k in str(kw or "").split(",") if k.strip()]}
    for code, text in wb["Banks"].iter_rows(min_row=2, max_col=2, values_only=True):
        if code is not None:
            banks.append((str(code).strip(), norm(text or code)))
    for text, code in wb["Vendors"].iter_rows(min_row=2, max_col=2, values_only=True):
        if text and code is not None:
            vendors.append((norm(text), str(code).strip()))
    for k, v, *_ in wb["Settings"].iter_rows(min_row=2, values_only=True):
        if k:
            settings[str(k)] = v
    settings["DayFirst"] = norm(settings.get("DayFirst", "yes")) in ("yes", "y", "true", "1")
    settings["DateWindowDays"] = int(settings.get("DateWindowDays") or 10)
    settings["SuspenseCode"] = str(settings.get("SuspenseCode") or "9999")
    settings["DefaultTaxCode"] = str(settings.get("DefaultTaxCode") or "T9")
    return accounts, banks, vendors, settings


# ======================================================================================
# Bank statements
# ======================================================================================
def read_bank_files(banks, dayfirst):
    rows = []
    files = sorted(p for p in BANK_DIR.glob("*") if p.suffix.lower() in (".csv", ".xlsx", ".xls"))
    if not files:
        sys.exit("No bank statements found in input/bank/")
    for f in files:
        fname = norm(f.stem)
        bank_code = next((c for c, t in banks if t and t in fname), None)
        if not bank_code:
            say(f"  ! {f.name}: no bank code matched (add a row in rules.xlsx > Banks). Using '?'.")
            bank_code = "?"
        df = read_table(f)
        dcol = find_col(df.columns, "date")
        desc = find_col(df.columns, "description", "narrative", "details", "particular",
                        "memo", "payee", "transaction", "reference", "name")
        out_col = find_col(df.columns, "paid out", "money out", "debit", "withdrawal", "payments")
        in_col = find_col(df.columns, "paid in", "money in", "credit", "deposit", "receipts")
        amt_col = find_col(df.columns, "amount", "value")
        n = 0
        for _, r in df.iterrows():
            d = to_date(r[dcol], dayfirst) if dcol else None
            if not d:
                continue
            if out_col and in_col and out_col != in_col:
                paid = to_amount(r[out_col])
                if not paid:
                    continue                        # money in -> not an expense
                paid = abs(paid)
            elif amt_col:
                a = to_amount(r[amt_col])
                if a is None or a >= 0:
                    continue                        # positive = money in
                paid = abs(a)
            else:
                continue
            rows.append({"date": d, "desc": str(r[desc]).strip() if desc else "",
                         "amount": round(paid, 2), "bank": bank_code, "file": f.name})
            n += 1
        say(f"  {f.name}: {n} payments (bank code {bank_code})")
    return rows


# ======================================================================================
# Bills (free OCR with RapidOCR)
# ======================================================================================
_ocr = None


def ocr_lines(path):
    global _ocr
    if _ocr is None:
        from rapidocr import RapidOCR
        _ocr = RapidOCR()
    images = []
    if path.suffix.lower() == ".pdf":
        import pypdfium2 as pdfium
        pdf = pdfium.PdfDocument(str(path))
        for i in range(len(pdf)):
            img = pdf[i].render(scale=2).to_pil()
            tmp = OUTPUT / f"_page_{path.stem}_{i}.png"
            img.save(tmp)
            images.append(tmp)
    else:
        images.append(path)
    lines = []
    for img in images:
        res = _ocr(str(img))
        if res is not None and res.txts:
            lines.extend(res.txts)
        if img != path:
            img.unlink(missing_ok=True)
    return [l.strip() for l in lines if l and l.strip()]


MONEY = re.compile(r"(?<![\d.])(\d{1,3}(?:[,\s]\d{3})*\.\d{2}|\d+\.\d{2})(?![\d])")
DATE_RX = re.compile(r"\b(\d{1,2}[/\-.]\d{1,2}[/\-.]\d{2,4}|\d{4}-\d{2}-\d{2}|"
                     r"\d{1,2}\s+[A-Za-z]{3,9}\.?\s+\d{2,4}|[A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4})\b")


def parse_bill(lines, dayfirst):
    text = "\n".join(lines)
    total = None
    # Prefer amount on (or right after) a line that says total / amount due
    for pri in ("grand total", "amount due", "total due", "balance due", "total paid", "total"):
        for i, l in enumerate(lines):
            if pri in l.lower() and "sub" not in l.lower():
                found = MONEY.findall(l) or (MONEY.findall(lines[i + 1]) if i + 1 < len(lines) else [])
                if found:
                    total = float(found[-1].replace(",", "").replace(" ", ""))
                    break
        if total is not None:
            break
    if total is None:
        nums = [float(m.replace(",", "").replace(" ", "")) for m in MONEY.findall(text)]
        total = max(nums) if nums else None
    bill_date = None
    for m in DATE_RX.findall(text):
        bill_date = to_date(m, dayfirst)
        if bill_date:
            break
    vendor = next((l for l in lines if re.search(r"[A-Za-z]{3,}", l)), "")
    return {"vendor": vendor[:60], "total": total, "date": bill_date, "text": text}


def read_bills(dayfirst):
    bills = []
    files = sorted(p for p in BILL_DIR.glob("*")
                   if p.suffix.lower() in (".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".pdf"))
    if files:
        say(f"Reading {len(files)} bills (first run downloads the free OCR model, please wait)...")
    for f in files:
        try:
            b = parse_bill(ocr_lines(f), dayfirst)
        except Exception as e:                     # keep going if one scan is unreadable
            say(f"  ! {f.name}: could not read ({e})")
            b = {"vendor": "", "total": None, "date": None, "text": ""}
        b["file"] = f.name
        say(f"  {f.name}: vendor='{b['vendor']}' total={b['total']} date={b['date']}")
        bills.append(b)
    return bills


# ======================================================================================
# Match + code
# ======================================================================================
def match_bills(payments, bills, window):
    for b in bills:
        if b["total"] is None:
            continue
        best, best_gap = None, None
        for p in payments:
            if p.get("bill") or abs(p["amount"] - b["total"]) > 0.01:
                continue
            gap = abs((p["date"] - b["date"]).days) if b["date"] else 0
            if b["date"] and gap > window:
                continue
            if best is None or gap < best_gap:
                best, best_gap = p, gap
        if best:
            best["bill"] = b
            b["matched"] = True


def word_hit(keyword, text):
    return re.search(r"(?<![a-z0-9])" + re.escape(keyword) + r"(?![a-z0-9])", text) is not None


def choose_code(text, accounts, vendors, suspense):
    t = norm(text)
    for key, code in vendors:                       # 1. your saved vendor rules
        if key and key in t:
            return code, "High", f"Vendor rule: '{key}'"
    scores = {}
    for code, a in accounts.items():                # 2. keyword rules
        if a["type"] != "expense":
            continue
        hits = [k for k in a["keywords"] if word_hit(k, t)]
        if hits:
            scores[code] = (len(hits), hits)
    if scores:
        ranked = sorted(scores.items(), key=lambda kv: -kv[1][0])
        code, (n, hits) = ranked[0]
        tie = len(ranked) > 1 and ranked[1][1][0] == n
        conf = "Low" if tie else "Medium"
        why = f"Keywords: {', '.join(hits[:4])}" + (f" (tie with {ranked[1][0]})" if tie else "")
        return code, conf, why
    return suspense, "Low", "No rule matched - please choose the code"   # 3. review


def vendor_key(row):
    """Short text used to save a vendor rule when you correct a code."""
    src = row.get("bill_vendor") or row.get("bank_desc") or ""
    words = re.sub(r"[^a-z& ]", " ", norm(src)).split()
    return " ".join(words[:2])


def run():
    accounts, banks, vendors, s = load_rules()
    OUTPUT.mkdir(exist_ok=True)
    say("Reading bank statements...")
    payments = read_bank_files(banks, s["DayFirst"])
    bills = read_bills(s["DayFirst"])
    match_bills(payments, bills, s["DateWindowDays"])

    rows = []
    for p in payments:
        b = p.get("bill")
        text = p["desc"] + " " + (b["text"] if b else "")
        code, conf, why = choose_code(text, accounts, vendors, s["SuspenseCode"])
        narrative = (f"{b['vendor']} - {p['desc']}" if b else p["desc"])[:100]
        rows.append({"date": p["date"], "code": code, "bank": p["bank"], "narr": narrative,
                     "amount": p["amount"], "status": "Matched to bill" if b else "Bank only - no bill",
                     "conf": conf, "why": why, "bill": b["file"] if b else "",
                     "bank_desc": p["desc"], "bill_vendor": b["vendor"] if b else ""})
    for b in bills:
        if not b.get("matched"):
            code, conf, why = choose_code(b["text"], accounts, vendors, s["SuspenseCode"])
            rows.append({"date": b["date"], "code": code, "bank": "?", "narr": b["vendor"],
                         "amount": b["total"], "status": "Bill only - no bank payment found",
                         "conf": "Low", "why": why, "bill": b["file"], "bank_desc": "",
                         "bill_vendor": b["vendor"]})
    rows.sort(key=lambda r: (r["date"] is None, r["date"] or datetime.min.date()))
    write_output(rows, accounts, s)


# ======================================================================================
# Output workbook
# ======================================================================================
YELLOW = PatternFill("solid", fgColor="FFF2CC")
RED = PatternFill("solid", fgColor="F8CBAD")


def style(wb):
    for ws in wb.worksheets:
        for c in ws[1]:
            c.font = Font(bold=True)
        ws.freeze_panes = "A2"
        for i, col in enumerate(ws.columns, 1):
            width = max((len(str(c.value)) for c in col if c.value is not None), default=8)
            ws.column_dimensions[get_column_letter(i)].width = min(max(10, width + 2), 60)


def write_output(rows, accounts, s):
    wb = Workbook()
    ws = wb.active
    ws.title = "Coded expenses"
    ws.append(["Date", "Expense code", "Bank code", "Narrative", "Amount",
               "Status", "Confidence", "Why this code", "Bill file", "Expense account name",
               "Suggested code", "Bank description", "Bill vendor"])
    for r in rows:
        ws.append([r["date"], r["code"], r["bank"], r["narr"], r["amount"], r["status"],
                   r["conf"], r["why"], r["bill"], accounts.get(r["code"], {}).get("name", ""),
                   r["code"], r["bank_desc"], r["bill_vendor"]])
        fill = RED if r["status"].startswith("Bill only") else YELLOW if r["conf"] == "Low" else None
        if fill:
            for c in ws[ws.max_row]:
                c.fill = fill
    for c in ws["A"][1:]:
        c.number_format = "DD/MM/YYYY" if s["DayFirst"] else "MM/DD/YYYY"
    for c in ws["E"][1:]:
        c.number_format = "#,##0.00"
    ws.auto_filter.ref = ws.dimensions
    for col in ("K", "L", "M"):                     # helper columns used by "learn"
        ws.column_dimensions[col].hidden = True

    ready = [r for r in rows if r["bank"] != "?" and r["amount"] is not None and r["date"]]

    # Sage 50 "audit trail" import: Type BP = bank payment (one line per payment)
    sg = wb.create_sheet("Sage 50 import")
    sg.append(["Type", "Account Reference", "Nominal A/C Ref", "Department Code", "Date",
               "Reference", "Details", "Net Amount", "Tax Code", "Tax Amount"])
    for i, r in enumerate(ready, 1):
        sg.append(["BP", r["bank"], r["code"], 0, r["date"].strftime("%d/%m/%Y"),
                   f"AUTO{i:04d}", r["narr"][:60], r["amount"], s["DefaultTaxCode"], 0])

    # Generic double-entry journal (Sage Accounting cloud, QuickBooks, others)
    jn = wb.create_sheet("Journal import")
    jn.append(["Journal No", "Journal Date", "Account Code", "Account Name", "Debits",
               "Credits", "Description"])
    for i, r in enumerate(ready, 1):
        d = r["date"].strftime("%d/%m/%Y")
        jn.append([f"AUTO{i:04d}", d, r["code"], accounts.get(r["code"], {}).get("name", ""),
                   r["amount"], None, r["narr"][:100]])
        jn.append([f"AUTO{i:04d}", d, r["bank"], accounts.get(r["bank"], {}).get("name", ""),
                   None, r["amount"], r["narr"][:100]])

    sm = wb.create_sheet("Summary")
    sm.append(["Item", "Count", "Total"])
    for label, f in [("All rows", lambda r: True),
                     ("Matched to bill", lambda r: r["status"].startswith("Matched")),
                     ("Bank only - no bill", lambda r: r["status"].startswith("Bank only")),
                     ("Bill only - no bank payment", lambda r: r["status"].startswith("Bill only")),
                     ("Needs review (Low confidence)", lambda r: r["conf"] == "Low")]:
        sel = [r for r in rows if f(r)]
        sm.append([label, len(sel), round(sum(r["amount"] or 0 for r in sel), 2)])
    style(wb)
    for col in ("K", "L", "M"):
        ws.column_dimensions[col].hidden = True
    wb.save(RESULT)
    for sheet, name in (("Sage 50 import", "sage50_import.csv"), ("Journal import", "journal_import.csv")):
        pd.DataFrame(list(wb[sheet].values)[1:], columns=list(wb[sheet].values)[0]) \
            .to_csv(OUTPUT / name, index=False)
    say(f"\nDone. {len(rows)} rows -> {RESULT.relative_to(BASE)}")
    say("  Yellow rows = please check the code.  Red rows = bill with no bank payment.")
    say("  Change any wrong code in column B, save, then run:  py expense_coder.py learn")


# ======================================================================================
# LEARN: your corrections -> vendor rules
# ======================================================================================
def learn():
    if not RESULT.exists():
        sys.exit("Nothing to learn from yet - run first.")
    ws = load_workbook(RESULT, data_only=True)["Coded expenses"]
    wb = load_workbook(RULES)
    vs = wb["Vendors"]
    existing = {norm(r[0]) for r in vs.iter_rows(min_row=2, max_col=1, values_only=True) if r[0]}
    added = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        code, suggested, bank_desc, bill_vendor = row[1], row[10], row[11], row[12]
        if code is None or str(code).strip() == str(suggested).strip():
            continue
        key = vendor_key({"bill_vendor": bill_vendor, "bank_desc": bank_desc})
        if key and key not in existing:
            vs.append([key, str(code).strip()])
            existing.add(key)
            added += 1
            say(f"  learned: '{key}' -> {code}")
    wb.save(RULES)
    say(f"Saved {added} new vendor rule(s) to rules.xlsx. They apply on the next run.")


if __name__ == "__main__":
    cmd = sys.argv[1].lower() if len(sys.argv) > 1 else "run"
    {"setup": setup, "run": run, "learn": learn}.get(cmd, run)()
