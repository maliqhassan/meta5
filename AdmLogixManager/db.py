"""SQLite storage layer. All data lives in one file: data/office.db"""
import json
import os
import re
import shutil
import sqlite3
import sys
from datetime import date, datetime

if getattr(sys, "frozen", False):          # running as packaged .exe
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DATA_DIR = os.path.join(BASE_DIR, "data")
EXPORT_DIR = os.path.join(BASE_DIR, "exports")
BACKUP_DIR = os.path.join(BASE_DIR, "backups")
DB_PATH = os.path.join(DATA_DIR, "office.db")

EMPLOYEE_STATUSES = ["Active", "On Leave", "Removed", "Final Exit"]
COMPANY_CATEGORIES = ["Current Client", "Previous Client", "Supplier", "Other"]
TXN_TYPES = ["Expense", "Income"]
PAY_METHODS = ["Bank", "Cash"]

DEFAULT_CATEGORIES = [
    ("Salary", "Expense"),
    ("Iqama / Visa / Govt Fees", "Expense"),
    ("Bank Charges", "Expense"),
    ("Vehicles & Maintenance", "Expense"),
    ("Fuel", "Expense"),
    ("Violations", "Expense"),
    ("Rent", "Expense"),
    ("Daily Cash Expense", "Expense"),
    ("Purchase / Assets", "Expense"),
    ("Office & Admin", "Expense"),
    ("Other Expense", "Expense"),
    ("Client Payment", "Income"),
    ("Refund", "Income"),
    ("Other Income", "Income"),
]

DEFAULT_SETTINGS = {
    "company_name": "Adm Logix Rental Equipment",
    "currency": "SAR",
    "report_title": "List of Daily Expenses",
    "report_subtitle": "Through Bank Account",
    "opening_balance": "0",
    "opening_date": "2026-01-01",
    "cash_opening": "0",
    "transfer_fee": "0.5",
    "vat_percent": "15",
    "prepared_by": "",
}

SCHEMA = """
CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT);
CREATE TABLE IF NOT EXISTS categories (
    id INTEGER PRIMARY KEY, name TEXT NOT NULL, type TEXT NOT NULL,
    UNIQUE(name, type));
CREATE TABLE IF NOT EXISTS custom_fields (
    id INTEGER PRIMARY KEY, entity TEXT NOT NULL, name TEXT NOT NULL,
    UNIQUE(entity, name));
CREATE TABLE IF NOT EXISTS employees (
    id INTEGER PRIMARY KEY,
    emp_code TEXT, iqama_no TEXT, name TEXT NOT NULL, short_name TEXT,
    basic_salary REAL DEFAULT 0, status TEXT DEFAULT 'Active',
    iqama_expiry TEXT, phone TEXT, nationality TEXT, join_date TEXT,
    notes TEXT, extra TEXT DEFAULT '{}');
CREATE TABLE IF NOT EXISTS companies (
    id INTEGER PRIMARY KEY,
    code TEXT, name TEXT NOT NULL, category TEXT DEFAULT 'Current Client',
    contact_person TEXT, phone TEXT, vat_no TEXT, notes TEXT,
    extra TEXT DEFAULT '{}');
CREATE TABLE IF NOT EXISTS transactions (
    id INTEGER PRIMARY KEY,
    date TEXT NOT NULL, type TEXT NOT NULL, category TEXT,
    description TEXT NOT NULL, amount REAL NOT NULL, remark TEXT,
    method TEXT DEFAULT 'Bank',
    employee_id INTEGER, company_id INTEGER,
    source TEXT DEFAULT 'manual', ref_id INTEGER,
    extra TEXT DEFAULT '{}', created_at TEXT);
CREATE INDEX IF NOT EXISTS ix_txn_date ON transactions(date);
CREATE TABLE IF NOT EXISTS payroll (
    id INTEGER PRIMARY KEY, month TEXT NOT NULL, employee_id INTEGER NOT NULL,
    basic_salary REAL DEFAULT 0, allowance REAL DEFAULT 0,
    overtime REAL DEFAULT 0, deduction REAL DEFAULT 0, advance REAL DEFAULT 0,
    paid_amount REAL DEFAULT 0, paid_date TEXT, method TEXT DEFAULT 'Bank',
    remark TEXT, UNIQUE(month, employee_id));
CREATE TABLE IF NOT EXISTS invoices (
    id INTEGER PRIMARY KEY, company_id INTEGER NOT NULL,
    invoice_no TEXT, issue_date TEXT, period TEXT, description TEXT,
    amount REAL DEFAULT 0, remark TEXT);
CREATE TABLE IF NOT EXISTS receipts (
    id INTEGER PRIMARY KEY, invoice_id INTEGER NOT NULL, date TEXT,
    amount REAL NOT NULL, txn_id INTEGER, remark TEXT);
"""


# ---------------------------------------------------------------- dates
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def parse_date(text):
    """Accepts 10-Oct-2026, 10-Oct-26, 10/10/2026, 2026-10-10, 10.10.2026."""
    if isinstance(text, (date, datetime)):
        return text if isinstance(text, date) and not isinstance(text, datetime) else text.date()
    text = (text or "").strip()
    if not text:
        return None
    for fmt in ("%d-%b-%Y", "%d-%b-%y", "%d-%B-%Y", "%d/%m/%Y", "%Y-%m-%d",
                "%d.%m.%Y", "%d-%m-%Y", "%d %b %Y", "%d/%m/%y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None


def fmt_date(value):
    """ISO string / date -> 10-Oct-2026 for display."""
    if not value:
        return ""
    d = parse_date(value) if not isinstance(value, date) else value
    return d.strftime("%d-%b-%Y") if d else str(value)


def iso(d):
    return d.isoformat() if d else None


def money(v):
    try:
        return f"{float(v or 0):,.2f}"
    except (TypeError, ValueError):
        return str(v)


def to_float(text):
    if text is None:
        return 0.0
    if isinstance(text, (int, float)):
        return float(text)
    text = str(text).replace(",", "").strip()
    if not text:
        return 0.0
    try:
        return float(text)
    except ValueError:
        pass
    # allow simple sums like "400+120" (as written in the old sheets' Remark column)
    if re.fullmatch(r"[0-9.+\-*/() ]+", text):
        try:
            return float(eval(text, {"__builtins__": {}}, {}))
        except Exception:
            pass
    raise ValueError(f"'{text}' is not a valid amount")


# ---------------------------------------------------------------- database
class Database:
    def __init__(self, path=DB_PATH):
        path = os.path.abspath(path)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        self.path = path
        self.conn = sqlite3.connect(path)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys = ON")
        self.conn.executescript(SCHEMA)
        for k, v in DEFAULT_SETTINGS.items():
            self.conn.execute("INSERT OR IGNORE INTO settings VALUES (?,?)", (k, v))
        if not self.conn.execute("SELECT 1 FROM categories LIMIT 1").fetchone():
            self.conn.executemany("INSERT INTO categories(name,type) VALUES (?,?)",
                                  DEFAULT_CATEGORIES)
        self.conn.commit()

    # -- generic helpers
    def q(self, sql, params=()):
        return self.conn.execute(sql, params).fetchall()

    def one(self, sql, params=()):
        return self.conn.execute(sql, params).fetchone()

    def scalar(self, sql, params=()):
        r = self.conn.execute(sql, params).fetchone()
        return (r[0] if r and r[0] is not None else 0)

    def run(self, sql, params=()):
        cur = self.conn.execute(sql, params)
        self.conn.commit()
        return cur.lastrowid

    def _save(self, table, data, row_id=None):
        data = dict(data)
        if "extra" in data and isinstance(data["extra"], dict):
            data["extra"] = json.dumps(data["extra"], ensure_ascii=False)
        cols = list(data)
        if row_id:
            self.run(f"UPDATE {table} SET {', '.join(c + '=?' for c in cols)} WHERE id=?",
                     [data[c] for c in cols] + [row_id])
            return row_id
        return self.run(f"INSERT INTO {table} ({', '.join(cols)}) VALUES ({', '.join('?' * len(cols))})",
                        [data[c] for c in cols])

    # -- settings
    def get(self, key, default=""):
        r = self.one("SELECT value FROM settings WHERE key=?", (key,))
        return r[0] if r else default

    def set(self, key, value):
        self.run("INSERT OR REPLACE INTO settings VALUES (?,?)", (key, str(value)))

    # -- categories / custom fields
    def categories(self, ttype=None):
        if ttype:
            return [r[0] for r in self.q("SELECT name FROM categories WHERE type=? ORDER BY name", (ttype,))]
        return self.q("SELECT * FROM categories ORDER BY type, name")

    def custom_fields(self, entity):
        return [r[0] for r in self.q("SELECT name FROM custom_fields WHERE entity=? ORDER BY id", (entity,))]

    def add_custom_field(self, entity, name):
        self.run("INSERT OR IGNORE INTO custom_fields(entity,name) VALUES (?,?)", (entity, name.strip()))

    def delete_custom_field(self, entity, name):
        self.run("DELETE FROM custom_fields WHERE entity=? AND name=?", (entity, name))

    @staticmethod
    def extra(row):
        try:
            return json.loads(row["extra"] or "{}")
        except Exception:
            return {}

    # -- employees
    def employees(self, search="", status=""):
        sql, p = "SELECT * FROM employees WHERE 1=1", []
        if status and status != "All":
            sql += " AND status=?"; p.append(status)
        if search:
            sql += " AND (name LIKE ? OR emp_code LIKE ? OR iqama_no LIKE ? OR short_name LIKE ?)"
            p += [f"%{search}%"] * 4
        rows = self.q(sql, p)
        return sorted(rows, key=lambda r: _code_key(r["emp_code"]))

    def save_employee(self, data, row_id=None):
        return self._save("employees", data, row_id)

    def delete_employee(self, row_id):
        self.run("DELETE FROM payroll WHERE employee_id=?", (row_id,))
        self.run("UPDATE transactions SET employee_id=NULL WHERE employee_id=?", (row_id,))
        self.run("DELETE FROM employees WHERE id=?", (row_id,))

    def next_code(self, table, column, prefix):
        nums = []
        for (c,) in self.q(f"SELECT {column} FROM {table}"):
            c = (c or "").upper()
            digits = "".join(ch for ch in c.split("-")[-1] if ch.isdigit())
            if digits:
                nums.append(int(digits))
        return f"{prefix}-{(max(nums) + 1) if nums else 1}"

    # -- companies
    def companies(self, search="", category=""):
        sql, p = "SELECT * FROM companies WHERE 1=1", []
        if category and category != "All":
            sql += " AND category=?"; p.append(category)
        if search:
            sql += " AND (name LIKE ? OR code LIKE ?)"; p += [f"%{search}%"] * 2
        return sorted(self.q(sql, p), key=lambda r: _code_key(r["code"]))

    def save_company(self, data, row_id=None):
        return self._save("companies", data, row_id)

    def delete_company(self, row_id):
        for inv in self.q("SELECT id FROM invoices WHERE company_id=?", (row_id,)):
            self.delete_invoice(inv["id"])
        self.run("UPDATE transactions SET company_id=NULL WHERE company_id=?", (row_id,))
        self.run("DELETE FROM companies WHERE id=?", (row_id,))

    # -- transactions
    def save_txn(self, data, row_id=None):
        data = dict(data)
        if not row_id:
            data.setdefault("created_at", datetime.now().isoformat(timespec="seconds"))
        return self._save("transactions", data, row_id)

    def delete_txn(self, row_id):
        t = self.one("SELECT * FROM transactions WHERE id=?", (row_id,))
        if t and t["source"] == "payroll" and t["ref_id"]:
            self.run("UPDATE payroll SET paid_amount=MAX(0, paid_amount-?) WHERE id=?",
                     (t["amount"], t["ref_id"]))
        self.run("DELETE FROM receipts WHERE txn_id=?", (row_id,))
        self.run("DELETE FROM transactions WHERE id=?", (row_id,))

    def txns(self, date_from=None, date_to=None, ttype="", category="", search="",
             method=""):
        sql = ("SELECT t.*, e.name AS employee_name, c.name AS company_name "
               "FROM transactions t LEFT JOIN employees e ON e.id=t.employee_id "
               "LEFT JOIN companies c ON c.id=t.company_id WHERE 1=1")
        p = []
        if date_from:
            sql += " AND t.date>=?"; p.append(date_from)
        if date_to:
            sql += " AND t.date<=?"; p.append(date_to)
        if ttype and ttype != "All":
            sql += " AND t.type=?"; p.append(ttype)
        if category and category != "All":
            sql += " AND t.category=?"; p.append(category)
        if method and method != "All":
            sql += " AND t.method=?"; p.append(method)
        if search:
            sql += " AND (t.description LIKE ? OR t.remark LIKE ?)"; p += [f"%{search}%"] * 2
        return self.q(sql + " ORDER BY t.date, t.id", p)

    def balance_before(self, day_iso, method="Bank"):
        """Balance at the start of the given day (opening balance + all earlier movements)."""
        opening = to_float(self.get("opening_balance" if method == "Bank" else "cash_opening", "0"))
        inc = self.scalar("SELECT SUM(amount) FROM transactions WHERE type='Income' AND date<? AND method=?",
                          (day_iso, method))
        exp = self.scalar("SELECT SUM(amount) FROM transactions WHERE type='Expense' AND date<? AND method=?",
                          (day_iso, method))
        return opening + inc - exp

    def day_summary(self, day_iso, method="Bank"):
        start = self.balance_before(day_iso, method)
        inc = self.scalar("SELECT SUM(amount) FROM transactions WHERE type='Income' AND date=? AND method=?",
                          (day_iso, method))
        exp = self.scalar("SELECT SUM(amount) FROM transactions WHERE type='Expense' AND date=? AND method=?",
                          (day_iso, method))
        return {"opening": start, "income": inc, "expense": exp, "closing": start + inc - exp}

    def current_balance(self, method="Bank"):
        return self.balance_before("9999-12-31", method)

    def descriptions(self, limit=400):
        return [r[0] for r in self.q(
            "SELECT description FROM transactions GROUP BY description "
            "ORDER BY MAX(date) DESC LIMIT ?", (limit,))]

    # -- payroll
    def generate_payroll(self, month):
        added = 0
        for e in self.q("SELECT * FROM employees WHERE status IN ('Active','On Leave')"):
            if not self.one("SELECT 1 FROM payroll WHERE month=? AND employee_id=?", (month, e["id"])):
                self.run("INSERT INTO payroll(month, employee_id, basic_salary) VALUES (?,?,?)",
                         (month, e["id"], e["basic_salary"] or 0))
                added += 1
        return added

    def payroll(self, month):
        rows = self.q("SELECT p.*, e.emp_code, e.name, e.iqama_no FROM payroll p "
                      "JOIN employees e ON e.id=p.employee_id WHERE p.month=?", (month,))
        return sorted(rows, key=lambda r: _code_key(r["emp_code"]))

    @staticmethod
    def net_salary(r):
        return (r["basic_salary"] or 0) + (r["allowance"] or 0) + (r["overtime"] or 0) \
            - (r["deduction"] or 0) - (r["advance"] or 0)

    def save_payroll(self, data, row_id):
        return self._save("payroll", data, row_id)

    # -- invoices
    def invoices(self, company_id=None):
        sql = ("SELECT i.*, c.name AS company_name, "
               "(SELECT COALESCE(SUM(amount),0) FROM receipts r WHERE r.invoice_id=i.id) AS received "
               "FROM invoices i JOIN companies c ON c.id=i.company_id")
        if company_id:
            return self.q(sql + " WHERE i.company_id=? ORDER BY i.issue_date, i.id", (company_id,))
        return self.q(sql + " ORDER BY i.issue_date, i.id")

    def save_invoice(self, data, row_id=None):
        return self._save("invoices", data, row_id)

    def delete_invoice(self, row_id):
        for r in self.q("SELECT txn_id FROM receipts WHERE invoice_id=?", (row_id,)):
            if r["txn_id"]:
                self.run("DELETE FROM transactions WHERE id=?", (r["txn_id"],))
        self.run("DELETE FROM receipts WHERE invoice_id=?", (row_id,))
        self.run("DELETE FROM invoices WHERE id=?", (row_id,))

    def receivable_total(self):
        return self.scalar("SELECT SUM(amount) FROM invoices") - self.scalar("SELECT SUM(amount) FROM receipts")

    # -- backup
    def backup(self, dest=None):
        os.makedirs(BACKUP_DIR, exist_ok=True)
        dest = dest or os.path.join(BACKUP_DIR, f"office_{datetime.now():%Y%m%d_%H%M%S}.db")
        self.conn.commit()
        bck = sqlite3.connect(dest)
        with bck:
            self.conn.backup(bck)
        bck.close()
        return dest

    def restore(self, src):
        self.conn.close()
        self.backup_path = self.path + ".before_restore"
        shutil.copy2(self.path, self.backup_path)
        shutil.copy2(src, self.path)
        self.__init__(self.path)


def _code_key(code):
    code = code or ""
    digits = "".join(ch for ch in code if ch.isdigit())
    return (code.split("-")[0], int(digits) if digits else 10 ** 9, code)
