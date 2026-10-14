"""Import existing Excel files (employees, companies, daily bank expense workbook,
client statements) into the database."""
import re
from datetime import date, datetime

from openpyxl import load_workbook

from db import iso, parse_date, to_float

EXPENSE_RULES = [
    ("Bank Charges", ["transfer fee", "vat fee", "bank charge", "sms fee"]),
    ("Salary", ["salary", "salaries", "wages"]),
    ("Iqama / Visa / Govt Fees", ["iqama", "sadad", "exit", "re-entry", "reentry", "visa", "qiwa",
                                  "tamm", "muqeem", "gosi", "jawazat", "passport", "work permit",
                                  "baladiya", "chamber", "absher", "fee"]),
    ("Violations", ["violation", "mukhalfa", "fine"]),
    ("Fuel", ["fuel", "petrol", "diesel"]),
    ("Vehicles & Maintenance", ["bus", "tire", "tyre", "fahas", "coaster", "vehicle", "car ", "oil",
                                "workshop", "repair", "battery", "seat cover", "radiator", "readitor",
                                "a/c", "insurance", "istimara"]),
    ("Rent", ["rent"]),
    ("Daily Cash Expense", ["daily expense", "cash"]),
    ("Purchase / Assets", ["purchase", "payment for the"]),
]


def guess_category(desc, ttype="Expense"):
    d = (desc or "").lower()
    if ttype == "Income":
        return "Refund" if "refund" in d else "Client Payment"
    for cat, words in EXPENSE_RULES:
        if any(w in d for w in words):
            return cat
    return "Other Expense"


def _clean(v):
    return re.sub(r"\s+", " ", str(v)).strip() if v is not None else ""


def _find_header(ws, must, limit=15):
    """Return (row_number, {lower_header: col_index}) of first row containing all `must` words."""
    for r in range(1, min(ws.max_row, limit) + 1):
        vals = [_clean(c.value).lower() for c in ws[r]]
        if all(any(m in v for v in vals) for m in must):
            return r, {v: i for i, v in enumerate(vals) if v}
    return None, {}


def _col(headers, *keys):
    for h, i in headers.items():
        if any(k in h for k in keys):
            return i
    return None


def sheet_names(path):
    wb = load_workbook(path, read_only=True)
    names = wb.sheetnames
    wb.close()
    return names


# ------------------------------------------------------------ employees
def header_names(path, sheet):
    ws = load_workbook(path, data_only=True)[sheet]
    hr, h = _find_header(ws, ["name"])
    return [k for k, _ in sorted(h.items(), key=lambda x: x[1])] if hr else []


def import_employees(db, path, sheet, salary_header=None):
    ws = load_workbook(path, data_only=True)[sheet]
    hr, h = _find_header(ws, ["name"])
    if not hr:
        raise ValueError("Could not find a header row with a 'Name' column.")
    c_code = _col(h, "id no", "id")
    c_iq = _col(h, "iqama")
    c_name = _col(h, "name")
    c_sal = h.get(salary_header) if salary_header else _col(h, "salary", "bs ")
    added = updated = 0
    for row in ws.iter_rows(min_row=hr + 1, values_only=True):
        name = _clean(row[c_name]) if c_name is not None else ""
        if not name:
            continue
        text = " ".join(_clean(v).lower() for v in row)
        status = "Active"
        if "final exit" in text:
            status = "Final Exit"
        elif "remove" in text:
            status = "Removed"
        sal = row[c_sal] if c_sal is not None else None
        sal = sal if isinstance(sal, (int, float)) else None
        data = {"name": name,
                "emp_code": _clean(row[c_code]) if c_code is not None else "",
                "iqama_no": _clean(row[c_iq]) if c_iq is not None else ""}
        if sal is not None:
            data["basic_salary"] = float(sal)
        data["status"] = status
        existing = None
        if data["iqama_no"]:
            existing = db.one("SELECT id FROM employees WHERE iqama_no=?", (data["iqama_no"],))
        if not existing and data["emp_code"]:
            existing = db.one("SELECT id FROM employees WHERE emp_code=?", (data["emp_code"],))
        if existing:
            db.save_employee(data, existing["id"])
            updated += 1
        else:
            data.setdefault("basic_salary", 0)
            db.save_employee(data)
            added += 1
    return f"Employees: {added} added, {updated} updated."


# ------------------------------------------------------------ companies
def import_companies(db, path, sheet):
    ws = load_workbook(path, data_only=True)[sheet]
    hr, h = _find_header(ws, ["company"])
    if not hr:
        raise ValueError("Could not find a header row with a 'Company' column.")
    c_code = _col(h, "id no", "id")
    c_name = _col(h, "company")
    c_cat = _col(h, "category")
    added = updated = 0
    for row in ws.iter_rows(min_row=hr + 1, values_only=True):
        name = _clean(row[c_name]) if c_name is not None else ""
        code = _clean(row[c_code]) if c_code is not None else ""
        if not name:
            continue
        if code and name.upper().startswith(code.upper() + " "):
            name = name[len(code) + 1:].strip()
        data = {"name": name}
        if code:
            data["code"] = code
        if c_cat is not None and row[c_cat]:
            data["category"] = _clean(row[c_cat]).title()
        existing = db.one("SELECT id FROM companies WHERE UPPER(REPLACE(name,' ',''))=?",
                          (name.upper().replace(" ", ""),))
        if existing:
            db.save_company(data, existing["id"])
            updated += 1
        else:
            data.setdefault("code", db.next_code("companies", "code", "C"))
            data.setdefault("category", "Current Client")
            db.save_company(data)
            added += 1
    return f"Companies: {added} added, {updated} updated."


def _match_company(db, text):
    t = (text or "").lower().replace(" ", "")
    best = None
    for c in db.q("SELECT id, name FROM companies"):
        first = re.split(r"[\s\-]+", c["name"].lower())
        key = first[0] if first and len(first[0]) > 3 else "".join(first[:2])
        if key and key[:5] in t:
            best = c["id"]
    return best


# ------------------------------------------------------------ bank expense workbook
SKIP_LABELS = ("opening", "closing", "difference", "total amount", "list of daily", "through bank")


def _sheet_date(name):
    m = re.fullmatch(r"\s*(\d{2})(\d{2})(\d{4})\s*", name)     # e.g. 15012026
    if m:
        try:
            return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            return None
    m = re.search(r"(\d{1,2})\s*-\s*([A-Za-z]+)\s*-\s*(\d{4})", name)
    if not m:
        return None
    try:
        return datetime.strptime(f"{m.group(1)}-{m.group(2)[:3]}-{m.group(3)}", "%d-%b-%Y").date()
    except ValueError:
        return None


def parse_bank_workbook(path):
    """Returns list of dicts: {date, start_balance, end_balance, items:[...]} per sheet."""
    wb = load_workbook(path, data_only=True)
    days = []
    for ws in wb.worksheets:
        sd = _sheet_date(ws.title)
        hr, _ = _find_header(ws, ["description"], limit=20)
        if not hr:
            continue
        balances, income, items = {}, [], []
        for row in ws.iter_rows(min_row=1, max_row=hr - 1, values_only=True):
            vals = [v for v in row if v is not None and _clean(v) != ""]
            if len(vals) < 2 or not isinstance(vals[0], str):
                continue
            label = _clean(vals[0])
            num = next((v for v in vals[1:] if isinstance(v, (int, float))), None)
            if num is None:
                continue
            low = label.lower()
            if low.startswith("opening"):
                balances["opening"] = num
            elif low.startswith("closing"):
                balances["closing"] = num
            elif low.startswith("difference"):
                continue
            else:
                income.append({"type": "Income", "description": label, "amount": float(num)})
        for row in ws.iter_rows(min_row=hr + 1, values_only=True):
            vals = list(row)
            first = next((v for v in vals if v is not None and _clean(v)), None)
            if first is None:
                continue
            if isinstance(first, str) and first.strip().lower().startswith("total"):
                break
            # Sr. No. is optional (some rows leave it blank); the date column is optional too
            start = 1 if isinstance(vals[0], (int, float)) else 0
            d = next((v for v in vals if isinstance(v, datetime)), None)
            texts = [(i, v) for i, v in enumerate(vals) if i >= start and isinstance(v, str) and _clean(v)]
            if texts:
                di, desc = texts[0]
            else:
                di = vals.index(d) if d is not None else start - 1
                desc = "(No description in old sheet)"
            amount = next((v for v in vals[di + 1:] if isinstance(v, (int, float))), None)
            if amount is None:
                continue
            ai = vals.index(amount, di + 1)
            remark = next((_clean(v) for v in vals[ai + 1:] if isinstance(v, str) and _clean(v)), "")
            items.append({"type": "Expense", "description": _clean(desc), "amount": float(amount),
                          "remark": remark, "date": d.date() if d else None})
        day = sd or next((i["date"] for i in items if i.get("date")), None)
        if not day:
            continue
        if sd:  # the sheet's own date is what its balance refers to
            for it in items:
                it["date"] = sd
        # The old sheets are not consistent about which of "Opening"/"Closing" is the balance
        # before the day, so work it out from the numbers: after = before + income - expense.
        before = after = None
        o, c = balances.get("opening"), balances.get("closing")
        if o is not None and c is not None:
            move = sum(i["amount"] for i in income) - sum(i["amount"] for i in items)
            if abs(o + move - c) <= abs(c + move - o):
                before, after = o, c
            else:
                before, after = c, o
        for it in income:
            it["date"] = day
        for it in items:
            it["date"] = it["date"] or day
        days.append({"date": day, "sheet": ws.title, "before": before, "after": after, "items": income + items})
    days.sort(key=lambda x: x["date"])
    return days


def import_bank_workbook(db, path, set_opening=True):
    days = parse_bank_workbook(path)
    if not days:
        raise ValueError("No daily expense sheets were found in this workbook.")
    done_dates = {r[0] for r in db.q("SELECT DISTINCT date FROM transactions WHERE source='import'")}
    n_items = n_skip = 0
    for day in days:
        dates_in_sheet = {iso(i["date"]) for i in day["items"]}
        if dates_in_sheet & done_dates:
            n_skip += 1
            continue
        for it in day["items"]:
            db.conn.execute(
                "INSERT INTO transactions(date,type,category,description,amount,remark,method,"
                "company_id,source,created_at) VALUES (?,?,?,?,?,?,?,?,?,?)",
                (iso(it["date"]), it["type"], guess_category(it["description"], it["type"]),
                 it["description"], it["amount"], it.get("remark", ""), "Bank",
                 _match_company(db, it["description"]) if it["type"] == "Income" else None,
                 "import", datetime.now().isoformat(timespec="seconds")))
            n_items += 1
    db.conn.commit()
    msg = [f"Bank workbook: {len(days) - n_skip} day-sheets imported ({n_items} entries)."]
    if n_skip:
        msg.append(f"{n_skip} sheets skipped (already imported).")
    first = days[0]
    has_manual_before = db.scalar("SELECT COUNT(*) FROM transactions WHERE source!='import' AND date<?",
                                  (iso(first["date"]),))
    if set_opening and first.get("before") is not None and not has_manual_before:
        db.set("opening_balance", round(first["before"], 2))
        db.set("opening_date", iso(first["date"]))
        msg.append(f"Opening balance set to {first['before']:,.2f} on {first['date']:%d-%b-%Y}.")
    # check balances against the old sheets
    mism = []
    for day in days:
        if day.get("after") is None:
            continue
        ours = db.day_summary(iso(day["date"]))["closing"]
        if abs(ours - day["after"]) > 0.05:
            mism.append(f"{day['date']:%d-%b-%Y}: sheet {day['after']:,.2f} vs computed {ours:,.2f}")
    last = days[-1]
    msg.append(f"Balance after {last['date']:%d-%b-%Y}: {db.day_summary(iso(last['date']))['closing']:,.2f}")
    if mism:
        msg.append(f"{len(mism)} day(s) where the old sheet balance differs from the computed one "
                   f"(usually a typing slip in the old file):")
        msg += ["   " + m for m in mism[:15]]
        if len(mism) > 15:
            msg.append(f"   ... and {len(mism) - 15} more")
    return "\n".join(msg)


# ------------------------------------------------------------ client statement (Spectra style)
def import_statement(db, path, sheet):
    ws = load_workbook(path, data_only=True)[sheet]
    hr, h = _find_header(ws, ["amount", "description"], limit=30)
    company = next((_clean(v) for row in ws.iter_rows(max_row=hr or 1, values_only=True) for v in row
                    if isinstance(v, str) and _clean(v) and "date" not in v.lower()), sheet)
    if not hr:
        raise ValueError("Could not find the invoice table (needs Description and Amount columns).")
    c_inv = _col(h, "inovice", "invoice", "inv")
    c_month = _col(h, "month", "period")
    c_desc = _col(h, "description")
    c_amt = _col(h, "amount")
    c_rec = _col(h, "received")
    c_rem = _col(h, "remark")
    comp = db.one("SELECT id FROM companies WHERE UPPER(name)=?", (company.upper(),))
    cid = comp["id"] if comp else db.save_company(
        {"code": db.next_code("companies", "code", "C"), "name": company, "category": "Current Client"})
    hdr_date = None
    for row in ws.iter_rows(min_row=1, max_row=hr, values_only=True):
        for v in row:
            m = re.search(r"date[:\s]*([\d]{1,2}-[A-Za-z]+-\d{4})", _clean(v), re.I)
            if m:
                hdr_date = parse_date(m.group(1))
    n = 0
    for row in ws.iter_rows(min_row=hr + 1, values_only=True):
        amt = row[c_amt] if c_amt is not None else None
        if not isinstance(row[0], (int, float)) or not isinstance(amt, (int, float)):
            continue
        inv_no = _clean(row[c_inv]) if c_inv is not None else ""
        if inv_no and db.one("SELECT 1 FROM invoices WHERE company_id=? AND invoice_no=?", (cid, inv_no)):
            continue
        inv_id = db.save_invoice({"company_id": cid, "invoice_no": inv_no,
                                  "issue_date": iso(hdr_date or date.today()),
                                  "period": _clean(row[c_month]) if c_month is not None else "",
                                  "description": _clean(row[c_desc]) if c_desc is not None else "",
                                  "amount": float(amt),
                                  "remark": _clean(row[c_rem]) if c_rem is not None else ""})
        rec = row[c_rec] if c_rec is not None else None
        if isinstance(rec, (int, float)) and rec:
            db.run("INSERT INTO receipts(invoice_id,date,amount,remark) VALUES (?,?,?,?)",
                   (inv_id, iso(hdr_date or date.today()), float(rec), "Imported"))
        n += 1
    return f"Statement for {company}: {n} invoice(s) imported."


def detect_kind(path):
    """Best guess of what an Excel file contains."""
    wb = load_workbook(path, data_only=True)
    try:
        dated = sum(1 for n in wb.sheetnames if _sheet_date(n))
        if dated >= 1:
            ws = wb[wb.sheetnames[0]]
            for row in ws.iter_rows(max_row=10, values_only=True):
                if any("expense" in _clean(v).lower() or "opening" in _clean(v).lower() for v in row):
                    return "bank"
        for name in wb.sheetnames:
            ws = wb[name]
            text = " ".join(_clean(v).lower() for row in ws.iter_rows(max_row=20, values_only=True)
                            for v in row if v is not None)
            if "iqama" in text:
                return "employees"
            if "inovice" in text or "invoice" in text:
                return "statement"
            if "company" in text:
                return "companies"
        return "unknown"
    finally:
        wb.close()
