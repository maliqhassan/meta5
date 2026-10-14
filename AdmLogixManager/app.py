"""Adm Logix Office Manager - employees, salaries, daily bank/cash book, clients,
dashboard and Excel/PDF reports.  Run:  python app.py"""
import os
import tkinter as tk
from decimal import ROUND_HALF_UP, Decimal
from datetime import date, datetime, timedelta
from tkinter import filedialog, messagebox, ttk

import importer
import reports
from db import (BACKUP_DIR, COMPANY_CATEGORIES, DATA_DIR, EMPLOYEE_STATUSES, EXPORT_DIR,
                PAY_METHODS, TXN_TYPES, Database, fmt_date, money, parse_date, to_float)
from widgets import (COLORS, FONT, FONT_B, AutoCombo, DateEntry, Form, ScrollFrame, Table,
                     open_file, setup_style)

APP_NAME = "Office Manager"


def month_label(ym):
    return datetime.strptime(ym, "%Y-%m").strftime("%b-%Y")


def month_from_label(lbl):
    return datetime.strptime(lbl, "%b-%Y").strftime("%Y-%m")


def month_range(ym):
    y, m = map(int, ym.split("-"))
    first = date(y, m, 1)
    last = date(y + (m == 12), m % 12 + 1, 1) - timedelta(days=1)
    return first.isoformat(), last.isoformat()


def card(parent, **kw):
    f = tk.Frame(parent, bg=COLORS["card"], highlightbackground=COLORS["line"], highlightthickness=1, **kw)
    return f


class Page(ttk.Frame):
    title = ""

    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self.db = app.db
        head = ttk.Frame(self)
        head.pack(fill="x", padx=18, pady=(14, 6))
        ttk.Label(head, text=self.title, style="H1.TLabel").pack(side="left")
        self.head_right = ttk.Frame(head)
        self.head_right.pack(side="right")
        self.body = ttk.Frame(self)
        self.body.pack(fill="both", expand=True, padx=18, pady=(0, 14))

    def refresh(self):
        pass

    def export(self, fn, *args):
        try:
            path = fn(*args)
        except PermissionError:
            messagebox.showerror("File is open", "Please close the old copy of this file "
                                 "(it is open in Excel / PDF reader) and try again.")
            return
        except Exception as e:  # noqa
            messagebox.showerror("Export failed", str(e))
            return
        self.app.status(f"Saved: {path}")
        open_file(path)


# =================================================================== dashboard
class DashboardPage(Page):
    title = "Dashboard"

    def __init__(self, master, app):
        super().__init__(master, app)
        ttk.Button(self.head_right, text="+ Add Today's Entry", style="Accent.TButton",
                   command=lambda: app.show("daily")).pack(side="left", padx=4)
        ttk.Button(self.head_right, text="Today's PDF", command=lambda: self.export(
            reports.export_day, self.db, date.today().isoformat(), "pdf")).pack(side="left", padx=4)
        ttk.Button(self.head_right, text="Refresh", command=self.refresh).pack(side="left", padx=4)

        self.kpi_row = ttk.Frame(self.body)
        self.kpi_row.pack(fill="x")
        self.kpis = {}
        for i, (key, label) in enumerate([("bank", "Bank Balance"), ("cash", "Cash in Hand"),
                                          ("today", "Today's Expense"), ("m_exp", "This Month Expense"),
                                          ("m_inc", "This Month Income"), ("salary", "Salary Due (Month)"),
                                          ("recv", "Client Receivable")]):
            c = card(self.kpi_row)
            c.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 8, 0))
            self.kpi_row.columnconfigure(i, weight=1)
            tk.Label(c, text=label, bg=COLORS["card"], fg=COLORS["muted"], font=("Segoe UI", 9)).pack(
                anchor="w", padx=12, pady=(10, 0))
            v = tk.Label(c, text="-", bg=COLORS["card"], fg=COLORS["text"], font=("Segoe UI", 15, "bold"))
            v.pack(anchor="w", padx=12)
            s = tk.Label(c, text="", bg=COLORS["card"], fg=COLORS["muted"], font=("Segoe UI", 8))
            s.pack(anchor="w", padx=12, pady=(0, 10))
            self.kpis[key] = (v, s)

        mid = ttk.Frame(self.body)
        mid.pack(fill="both", expand=True, pady=(12, 0))
        mid.columnconfigure(0, weight=3)
        mid.columnconfigure(1, weight=2)
        mid.rowconfigure(0, weight=1)
        mid.rowconfigure(1, weight=1)

        c1 = card(mid)
        c1.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=(0, 8))
        tk.Label(c1, text="Income vs Expense - last 6 months", bg=COLORS["card"], font=FONT_B).pack(
            anchor="w", padx=12, pady=(10, 0))
        self.chart1 = tk.Canvas(c1, bg=COLORS["card"], highlightthickness=0, height=220)
        self.chart1.pack(fill="both", expand=True, padx=8, pady=8)
        self.chart1.bind("<Configure>", lambda e: self.draw_months())

        c2 = card(mid)
        c2.grid(row=0, column=1, sticky="nsew", pady=(0, 8))
        self.cat_title = tk.Label(c2, text="This month's expense by category", bg=COLORS["card"], font=FONT_B)
        self.cat_title.pack(anchor="w", padx=12, pady=(10, 0))
        self.chart2 = tk.Canvas(c2, bg=COLORS["card"], highlightthickness=0, height=220)
        self.chart2.pack(fill="both", expand=True, padx=8, pady=8)
        self.chart2.bind("<Configure>", lambda e: self.draw_categories())

        c3 = card(mid)
        c3.grid(row=1, column=0, sticky="nsew", padx=(0, 8))
        tk.Label(c3, text="Latest entries", bg=COLORS["card"], font=FONT_B).pack(anchor="w", padx=12, pady=(10, 4))
        self.recent = Table(c3, [("date", "Date", 95, "center"), ("type", "Type", 70, "center"),
                                 ("desc", "Description", 250, "w"), ("amt", "Amount", 95, "e")], height=8)
        self.recent.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        c4 = card(mid)
        c4.grid(row=1, column=1, sticky="nsew")
        tk.Label(c4, text="Reminders", bg=COLORS["card"], font=FONT_B).pack(anchor="w", padx=12, pady=(10, 4))
        self.alerts = tk.Text(c4, height=8, relief="flat", bg=COLORS["card"], font=FONT, wrap="word")
        self.alerts.pack(fill="both", expand=True, padx=12, pady=(0, 10))
        self.month_data = []
        self.cat_data = []

    def refresh(self):
        db = self.db
        today = date.today()
        ym = today.strftime("%Y-%m")
        m1, m2 = month_range(ym)
        cur = db.get("currency")
        bank = db.current_balance("Bank")
        cash = db.current_balance("Cash")
        t = db.day_summary(today.isoformat())
        tc = db.day_summary(today.isoformat(), "Cash")
        m_exp = db.scalar("SELECT SUM(amount) FROM transactions WHERE type='Expense' AND date BETWEEN ? AND ?", (m1, m2))
        m_inc = db.scalar("SELECT SUM(amount) FROM transactions WHERE type='Income' AND date BETWEEN ? AND ?", (m1, m2))
        pr = db.payroll(ym)
        due = sum(max(0, db.net_salary(r) - (r["paid_amount"] or 0)) for r in pr)
        unpaid = sum(1 for r in pr if db.net_salary(r) - (r["paid_amount"] or 0) > 0.009)
        recv = db.receivable_total()
        last_day = db.one("SELECT MAX(date) FROM transactions")[0]
        vals = {
            "bank": (money(bank), f"{cur} · last entry {fmt_date(last_day) or '-'}"),
            "cash": (money(cash), cur),
            "today": (money(t["expense"] + tc["expense"]), f"{today:%d-%b-%Y}"),
            "m_exp": (money(m_exp), f"{today:%B %Y}"),
            "m_inc": (money(m_inc), f"{today:%B %Y}"),
            "salary": (money(due), f"{unpaid} unpaid" if pr else "payroll not generated"),
            "recv": (money(recv), "invoices not yet received"),
        }
        for k, (v, s) in vals.items():
            self.kpis[k][0].config(text=v, fg=COLORS["red"] if k == "bank" and bank < 0 else COLORS["text"])
            self.kpis[k][1].config(text=s)

        # 6 months
        self.month_data = []
        y, m = today.year, today.month
        for i in range(5, -1, -1):
            mm = m - i
            yy = y + (mm - 1) // 12
            mm = (mm - 1) % 12 + 1
            a, b = month_range(f"{yy}-{mm:02d}")
            inc = db.scalar("SELECT SUM(amount) FROM transactions WHERE type='Income' AND date BETWEEN ? AND ?", (a, b))
            exp = db.scalar("SELECT SUM(amount) FROM transactions WHERE type='Expense' AND date BETWEEN ? AND ?", (a, b))
            self.month_data.append((date(yy, mm, 1).strftime("%b %y"), inc, exp))
        # categories (fall back to last month with data)
        rows = db.q("SELECT category, SUM(amount) s FROM transactions WHERE type='Expense' AND date BETWEEN ? AND ? "
                    "GROUP BY category ORDER BY s DESC", (m1, m2))
        label = f"Expense by category - {today:%B %Y}"
        if not rows and last_day:
            ld = parse_date(last_day)
            a, b = month_range(ld.strftime("%Y-%m"))
            rows = db.q("SELECT category, SUM(amount) s FROM transactions WHERE type='Expense' AND date BETWEEN ? "
                        "AND ? GROUP BY category ORDER BY s DESC", (a, b))
            label = f"Expense by category - {ld:%B %Y}"
        self.cat_title.config(text=label)
        self.cat_data = [(r[0] or "Uncategorised", r[1]) for r in rows]
        self.draw_months()
        self.draw_categories()

        recent = db.q("SELECT * FROM transactions ORDER BY date DESC, id DESC LIMIT 25")
        self.recent.load([(r["id"], (fmt_date(r["date"]), r["type"], r["description"], money(r["amount"])),
                           ("income",) if r["type"] == "Income" else ()) for r in recent])

        notes = []
        soon = (today + timedelta(days=30)).isoformat()
        for e in db.q("SELECT emp_code, name, iqama_expiry FROM employees WHERE status='Active' AND "
                      "iqama_expiry IS NOT NULL AND iqama_expiry!='' AND iqama_expiry<=? ORDER BY iqama_expiry", (soon,)):
            word = "EXPIRED" if e["iqama_expiry"] < today.isoformat() else "expires"
            notes.append(f"• Iqama {word} {fmt_date(e['iqama_expiry'])}: {e['emp_code']} {e['name']}")
        if not pr:
            notes.append(f"• Salary sheet for {today:%B %Y} not generated yet (Salaries page).")
        elif unpaid:
            notes.append(f"• {unpaid} employee(s) not fully paid for {today:%B %Y} - total {money(due)}.")
        if not db.scalar("SELECT COUNT(*) FROM transactions WHERE date=?", (today.isoformat(),)):
            notes.append("• No entries recorded for today yet.")
        for inv in db.q("SELECT i.*, c.name cname, (SELECT COALESCE(SUM(amount),0) FROM receipts r WHERE "
                        "r.invoice_id=i.id) rec FROM invoices i JOIN companies c ON c.id=i.company_id"):
            if inv["amount"] - inv["rec"] > 0.009 and inv["issue_date"] and inv["issue_date"] <= (today - timedelta(days=30)).isoformat():
                notes.append(f"• Unpaid > 30 days: {inv['cname']} {inv['invoice_no'] or ''} - "
                             f"{money(inv['amount'] - inv['rec'])}")
        self.alerts.config(state="normal")
        self.alerts.delete("1.0", "end")
        self.alerts.insert("1.0", "\n".join(notes) if notes else "All good - nothing pending.")
        self.alerts.config(state="disabled")

    def draw_months(self):
        cv = self.chart1
        cv.delete("all")
        w, h = cv.winfo_width(), cv.winfo_height()
        if w < 50 or not self.month_data:
            return
        left, bottom, top = 70, h - 28, 22
        mx = max([max(i, e) for _, i, e in self.month_data] + [1])
        for k in range(5):
            yv = bottom - (bottom - top) * k / 4
            cv.create_line(left, yv, w - 10, yv, fill=COLORS["line"])
            cv.create_text(left - 6, yv, text=f"{mx * k / 4:,.0f}", anchor="e", fill=COLORS["muted"],
                           font=("Segoe UI", 8))
        n = len(self.month_data)
        slot = (w - left - 10) / n
        bw = min(26, slot / 3)
        for i, (lbl, inc, exp) in enumerate(self.month_data):
            x = left + slot * i + slot / 2
            for j, (val, col) in enumerate(((inc, COLORS["green"]), (exp, COLORS["accent"]))):
                bh = (bottom - top) * val / mx
                x0 = x - bw - 2 + j * (bw + 4)
                cv.create_rectangle(x0, bottom - bh, x0 + bw, bottom, fill=col, outline="")
            cv.create_text(x, bottom + 12, text=lbl, fill=COLORS["text"], font=("Segoe UI", 9))
        cv.create_rectangle(w - 170, 4, w - 160, 14, fill=COLORS["green"], outline="")
        cv.create_text(w - 156, 9, text="Income", anchor="w", font=("Segoe UI", 8))
        cv.create_rectangle(w - 100, 4, w - 90, 14, fill=COLORS["accent"], outline="")
        cv.create_text(w - 86, 9, text="Expense", anchor="w", font=("Segoe UI", 8))

    def draw_categories(self):
        cv = self.chart2
        cv.delete("all")
        w, h = cv.winfo_width(), cv.winfo_height()
        if w < 50:
            return
        if not self.cat_data:
            cv.create_text(w / 2, h / 2, text="No expenses yet", fill=COLORS["muted"])
            return
        data = self.cat_data[:8]
        mx = max(v for _, v in data) or 1
        total = sum(v for _, v in self.cat_data) or 1
        row_h = min(26, (h - 10) / len(data))
        lw = 150
        for i, (name, v) in enumerate(data):
            y = 6 + i * row_h
            cv.create_text(4, y + row_h / 2, text=name[:22], anchor="w", font=("Segoe UI", 9))
            bw = (w - lw - 120) * v / mx
            cv.create_rectangle(lw, y + 4, lw + bw, y + row_h - 4, fill=COLORS["accent"], outline="")
            cv.create_text(lw + bw + 6, y + row_h / 2, anchor="w", font=("Segoe UI", 8),
                           text=f"{v:,.0f} ({v / total:.0%})", fill=COLORS["muted"])


# =================================================================== daily book
class DailyPage(Page):
    title = "Daily Income / Expense"

    def __init__(self, master, app):
        super().__init__(master, app)
        self.edit_id = None
        bar = card(self.body)
        bar.pack(fill="x")
        inner = tk.Frame(bar, bg=COLORS["card"])
        inner.pack(fill="x", padx=10, pady=8)
        tk.Label(inner, text="Account", bg=COLORS["card"]).pack(side="left")
        self.method = ttk.Combobox(inner, values=PAY_METHODS, state="readonly", width=7)
        self.method.set("Bank")
        self.method.pack(side="left", padx=(4, 14))
        self.method.bind("<<ComboboxSelected>>", lambda e: self.refresh())
        tk.Label(inner, text="Date", bg=COLORS["card"]).pack(side="left")
        ttk.Button(inner, text="◀", width=3, command=lambda: self.shift(-1)).pack(side="left", padx=(4, 0))
        self.day = DateEntry(inner, on_change=self.refresh, style="Card.TFrame")
        self.day.pack(side="left", padx=2)
        ttk.Button(inner, text="▶", width=3, command=lambda: self.shift(1)).pack(side="left")
        ttk.Button(inner, text="Today", command=lambda: (self.day.set(date.today()), self.refresh())).pack(
            side="left", padx=6)
        self.sum_lbl = {}
        for key, lbl in (("closing", "Closing"), ("expense", "Expense"), ("income", "Income"),
                         ("opening", "Opening (B/F)")):
            f = tk.Frame(inner, bg=COLORS["card"])
            f.pack(side="right", padx=10)
            tk.Label(f, text=lbl, bg=COLORS["card"], fg=COLORS["muted"], font=("Segoe UI", 8)).pack(anchor="e")
            v = tk.Label(f, text="0.00", bg=COLORS["card"], font=("Segoe UI", 12, "bold"))
            v.pack(anchor="e")
            self.sum_lbl[key] = v

        self.table = Table(self.body, [("sr", "Sr.", 45, "center"), ("type", "Type", 75, "center"),
                                       ("cat", "Category", 170, "w"), ("desc", "Description", 380, "w"),
                                       ("amt", "Amount", 110, "e"), ("remark", "Remark", 140, "w"),
                                       ("link", "Employee / Client", 180, "w")], height=11)
        self.table.pack(fill="both", expand=True, pady=10)
        self.table.bind_select(self.load_selected)

        fcard = card(self.body)
        fcard.pack(fill="x")
        f = tk.Frame(fcard, bg=COLORS["card"])
        f.pack(fill="x", padx=10, pady=8)
        self.form_title = tk.Label(f, text="New entry", bg=COLORS["card"], font=FONT_B, fg=COLORS["accent"])
        self.form_title.grid(row=0, column=0, columnspan=6, sticky="w")
        self.ttype = tk.StringVar(value="Expense")
        tf = tk.Frame(f, bg=COLORS["card"])
        tf.grid(row=1, column=0, columnspan=2, sticky="w", pady=3)
        for t in TXN_TYPES:
            ttk.Radiobutton(tf, text=t, value=t, variable=self.ttype, command=self.type_changed).pack(side="left", padx=(0, 10))
        tk.Label(f, text="Category", bg=COLORS["card"]).grid(row=1, column=2, sticky="e", padx=4)
        self.cat = ttk.Combobox(f, width=26)
        self.cat.grid(row=1, column=3, sticky="ew")
        tk.Label(f, text="Amount", bg=COLORS["card"]).grid(row=1, column=4, sticky="e", padx=4)
        self.amount = ttk.Entry(f, width=14)
        self.amount.grid(row=1, column=5, sticky="w")
        tk.Label(f, text="Description", bg=COLORS["card"]).grid(row=2, column=0, sticky="w", pady=3)
        self.desc = AutoCombo(f, width=60)
        self.desc.grid(row=2, column=1, columnspan=3, sticky="ew", pady=3)
        self.desc.bind("<<ComboboxSelected>>", self.desc_chosen)
        tk.Label(f, text="Remark", bg=COLORS["card"]).grid(row=2, column=4, sticky="e", padx=4)
        self.remark = ttk.Entry(f, width=18)
        self.remark.grid(row=2, column=5, sticky="ew")
        tk.Label(f, text="Employee", bg=COLORS["card"]).grid(row=3, column=0, sticky="w", pady=3)
        self.emp = AutoCombo(f, width=34)
        self.emp.grid(row=3, column=1, sticky="ew")
        tk.Label(f, text="Client", bg=COLORS["card"]).grid(row=3, column=2, sticky="e", padx=4)
        self.comp = AutoCombo(f, width=26)
        self.comp.grid(row=3, column=3, sticky="ew")
        self.fee_var = tk.BooleanVar(value=False)
        ff = tk.Frame(f, bg=COLORS["card"])
        ff.grid(row=3, column=4, columnspan=2, sticky="w")
        ttk.Checkbutton(ff, text="+ Transfer fee", variable=self.fee_var).pack(side="left")
        self.fee_amt = ttk.Entry(ff, width=6)
        self.fee_amt.pack(side="left", padx=4)
        tk.Label(ff, text="+VAT", bg=COLORS["card"], fg=COLORS["muted"]).pack(side="left")
        self.extra_frame = tk.Frame(f, bg=COLORS["card"])
        self.extra_frame.grid(row=4, column=0, columnspan=6, sticky="ew")
        f.columnconfigure(1, weight=1)
        f.columnconfigure(3, weight=1)

        btns = tk.Frame(f, bg=COLORS["card"])
        btns.grid(row=5, column=0, columnspan=6, sticky="ew", pady=(8, 0))
        ttk.Button(btns, text="Save Entry  (Ctrl+S)", style="Accent.TButton", command=self.save).pack(side="left")
        ttk.Button(btns, text="New", command=self.clear).pack(side="left", padx=6)
        ttk.Button(btns, text="Delete", style="Danger.TButton", command=self.delete).pack(side="left")
        ttk.Button(btns, text="Month Workbook (Excel)", command=self.export_month).pack(side="right")
        ttk.Button(btns, text="Day Excel", command=lambda: self.export(
            reports.export_day, self.db, self.day.get(), "xlsx", self.method.get())).pack(side="right", padx=6)
        ttk.Button(btns, text="Day PDF", command=lambda: self.export(
            reports.export_day, self.db, self.day.get(), "pdf", self.method.get())).pack(side="right")
        self.amount.bind("<Return>", lambda e: self.save())
        self.extra_widgets = {}

    def shift(self, n):
        d = self.day.get_date() or date.today()
        self.day.set(d + timedelta(days=n))
        self.refresh()

    def type_changed(self):
        self.cat["values"] = self.db.categories(self.ttype.get())

    def desc_chosen(self, e=None):
        r = self.db.one("SELECT category, type FROM transactions WHERE description=? ORDER BY id DESC LIMIT 1",
                        (self.desc.get(),))
        if r and not self.edit_id:
            self.ttype.set(r["type"])
            self.type_changed()
            self.cat.set(r["category"] or "")

    def build_extra(self):
        for w in self.extra_frame.winfo_children():
            w.destroy()
        self.extra_widgets = {}
        for i, name in enumerate(self.db.custom_fields("transaction")):
            tk.Label(self.extra_frame, text=name, bg=COLORS["card"]).grid(row=0, column=i * 2, sticky="w", pady=3, padx=(0, 4))
            e = ttk.Entry(self.extra_frame, width=18)
            e.grid(row=0, column=i * 2 + 1, sticky="w", padx=(0, 12))
            self.extra_widgets[name] = e

    def refresh(self):
        db = self.db
        day = self.day.get() or date.today().isoformat()
        method = self.method.get()
        self.type_changed()
        self.desc.set_values(db.descriptions())
        self.emp.set_values([""] + [f"{e['emp_code']} | {e['name']}" for e in db.employees()])
        self.comp.set_values([""] + [f"{c['code']} | {c['name']}" for c in db.companies()])
        if not self.fee_amt.get():
            self.fee_amt.insert(0, db.get("transfer_fee"))
        self.build_extra()
        s = db.day_summary(day, method)
        for k, lbl in self.sum_lbl.items():
            lbl.config(text=money(s[k]))
        self.sum_lbl["closing"].config(fg=COLORS["red"] if s["closing"] < 0 else COLORS["text"])
        rows = db.txns(day, day, method=method)
        out, n = [], 0
        for r in sorted(rows, key=lambda r: (r["type"] != "Income", r["id"])):
            n += 1
            out.append((r["id"], (n, r["type"], r["category"] or "", r["description"], money(r["amount"]),
                                  r["remark"] or "", r["employee_name"] or r["company_name"] or ""),
                        ("income",) if r["type"] == "Income" else ()))
        self.table.load(out)
        if not self.edit_id:
            self.clear(keep_type=True)

    def _id_from(self, combo, table):
        v = combo.get().strip()
        if not v:
            return None
        code = v.split("|")[0].strip()
        col = "emp_code" if table == "employees" else "code"
        r = self.db.one(f"SELECT id FROM {table} WHERE {col}=?", (code,))
        if not r:
            raise ValueError(f"'{v}' is not in the {table} list - pick from the drop-down or leave empty.")
        return r["id"]

    def save(self):
        db = self.db
        try:
            day = self.day.get()
            if not day:
                raise ValueError("Please enter a valid date (e.g. 14-Oct-2026).")
            desc = self.desc.get().strip()
            if not desc:
                raise ValueError("Please enter a description.")
            amount = to_float(self.amount.get())
            if amount <= 0:
                raise ValueError("Amount must be more than zero.")
            emp_id = self._id_from(self.emp, "employees")
            comp_id = self._id_from(self.comp, "companies")
            fee = to_float(self.fee_amt.get()) if self.fee_var.get() else 0
        except ValueError as e:
            messagebox.showwarning("Check entry", str(e))
            return
        extra = {k: w.get().strip() for k, w in self.extra_widgets.items() if w.get().strip()}
        data = {"date": day, "type": self.ttype.get(), "category": self.cat.get().strip(),
                "description": desc, "amount": amount, "remark": self.remark.get().strip(),
                "method": self.method.get(), "employee_id": emp_id, "company_id": comp_id, "extra": extra}
        if data["category"] and data["category"] not in db.categories(data["type"]):
            db.run("INSERT OR IGNORE INTO categories(name,type) VALUES (?,?)", (data["category"], data["type"]))
        db.save_txn(data, self.edit_id)
        if fee and not self.edit_id:
            add_fee(db, day, fee, self.method.get())
        self.app.status(f"Saved: {desc} - {money(amount)}")
        self.edit_id = None
        self.refresh()
        self.desc.focus_set()

    def load_selected(self):
        tid = self.table.selected()
        if not tid:
            return
        r = self.db.one("SELECT t.*, e.emp_code, e.name en, c.code cc, c.name cn FROM transactions t "
                        "LEFT JOIN employees e ON e.id=t.employee_id LEFT JOIN companies c ON c.id=t.company_id "
                        "WHERE t.id=?", (tid,))
        self.edit_id = tid
        self.form_title.config(text=f"Editing entry #{tid}  (press New to add a fresh entry)")
        self.ttype.set(r["type"])
        self.type_changed()
        self.cat.set(r["category"] or "")
        self.desc.set(r["description"])
        self.amount.delete(0, "end")
        self.amount.insert(0, f"{r['amount']:g}")
        self.remark.delete(0, "end")
        self.remark.insert(0, r["remark"] or "")
        self.emp.set(f"{r['emp_code']} | {r['en']}" if r["en"] else "")
        self.comp.set(f"{r['cc']} | {r['cn']}" if r["cn"] else "")
        ex = self.db.extra(r)
        for k, w in self.extra_widgets.items():
            w.delete(0, "end")
            w.insert(0, ex.get(k, ""))
        self.fee_var.set(False)

    def clear(self, keep_type=False):
        self.edit_id = None
        self.form_title.config(text="New entry")
        if not keep_type:
            self.ttype.set("Expense")
            self.type_changed()
        for w in (self.amount, self.remark):
            w.delete(0, "end")
        self.desc.set("")
        self.cat.set("")
        self.emp.set("")
        self.comp.set("")
        for w in self.extra_widgets.values():
            w.delete(0, "end")
        self.table.tree.selection_remove(self.table.tree.selection())

    def delete(self):
        tid = self.table.selected()
        if not tid:
            messagebox.showinfo("Delete", "Select an entry in the list first.")
            return
        r = self.db.one("SELECT * FROM transactions WHERE id=?", (tid,))
        extra = ""
        if r["source"] == "payroll":
            extra = "\n\nThis is a salary payment - the employee's paid amount will be reduced."
        if r["source"] == "invoice":
            extra = "\n\nThis is a client receipt - the invoice will show it as not received."
        if messagebox.askyesno("Delete", f"Delete '{r['description']}' ({money(r['amount'])})?{extra}"):
            self.db.delete_txn(tid)
            self.edit_id = None
            self.refresh()

    def export_month(self):
        d = self.day.get_date() or date.today()
        self.export(reports.export_month_workbook, self.db, d.year, d.month, self.method.get())


def add_fee(db, day, fee, method="Bank", note=""):
    """Bank transfer fee + VAT lines, as written in the old sheets."""
    vat = float(Decimal(str(fee * to_float(db.get("vat_percent", "15")) / 100))
                .quantize(Decimal("0.01"), ROUND_HALF_UP))      # 0.5 -> 0.08, 1 -> 0.15
    db.save_txn({"date": day, "type": "Expense", "category": "Bank Charges", "description": "Transfer Fee",
                 "amount": fee, "remark": note, "method": method})
    if vat:
        db.save_txn({"date": day, "type": "Expense", "category": "Bank Charges", "description": "Vat Fee",
                     "amount": vat, "remark": note, "method": method})


# =================================================================== reports
class ReportsPage(Page):
    title = "Reports & Search"

    def __init__(self, master, app):
        super().__init__(master, app)
        top = card(self.body)
        top.pack(fill="x")
        f = tk.Frame(top, bg=COLORS["card"])
        f.pack(fill="x", padx=10, pady=8)
        tk.Label(f, text="From", bg=COLORS["card"]).grid(row=0, column=0)
        self.d1 = DateEntry(f, value=date.today().replace(day=1), style="Card.TFrame")
        self.d1.grid(row=0, column=1, padx=4)
        tk.Label(f, text="To", bg=COLORS["card"]).grid(row=0, column=2)
        self.d2 = DateEntry(f, style="Card.TFrame")
        self.d2.grid(row=0, column=3, padx=4)
        tk.Label(f, text="Type", bg=COLORS["card"]).grid(row=0, column=4, padx=(10, 2))
        self.ttype = ttk.Combobox(f, values=["All"] + TXN_TYPES, state="readonly", width=9)
        self.ttype.set("All")
        self.ttype.grid(row=0, column=5)
        tk.Label(f, text="Category", bg=COLORS["card"]).grid(row=0, column=6, padx=(10, 2))
        self.cat = ttk.Combobox(f, state="readonly", width=22)
        self.cat.grid(row=0, column=7)
        tk.Label(f, text="Account", bg=COLORS["card"]).grid(row=0, column=8, padx=(10, 2))
        self.method = ttk.Combobox(f, values=["All"] + PAY_METHODS, state="readonly", width=7)
        self.method.set("All")
        self.method.grid(row=0, column=9)
        tk.Label(f, text="Search", bg=COLORS["card"]).grid(row=1, column=0, pady=(6, 0))
        self.search = ttk.Entry(f, width=30)
        self.search.grid(row=1, column=1, columnspan=3, sticky="ew", padx=4, pady=(6, 0))
        self.search.bind("<Return>", lambda e: self.refresh())
        q = tk.Frame(f, bg=COLORS["card"])
        q.grid(row=1, column=4, columnspan=6, sticky="w", pady=(6, 0), padx=(10, 0))
        for lbl, fn in (("Today", self.q_today), ("This Month", self.q_month), ("Last Month", self.q_last),
                        ("This Year", self.q_year)):
            ttk.Button(q, text=lbl, command=fn).pack(side="left", padx=(0, 4))
        ttk.Button(q, text="Show", style="Accent.TButton", command=self.refresh).pack(side="left", padx=(8, 0))

        self.table = Table(self.body, [("date", "Date", 95, "center"), ("type", "Type", 70, "center"),
                                       ("cat", "Category", 160, "w"), ("desc", "Description", 360, "w"),
                                       ("link", "Employee / Client", 160, "w"), ("acc", "Acc.", 55, "center"),
                                       ("inc", "Income", 105, "e"), ("exp", "Expense", 105, "e"),
                                       ("remark", "Remark", 120, "w")])
        self.table.pack(fill="both", expand=True, pady=10)
        bot = ttk.Frame(self.body)
        bot.pack(fill="x")
        self.total_lbl = ttk.Label(bot, text="", font=FONT_B)
        self.total_lbl.pack(side="left")
        ttk.Button(bot, text="Export Excel", command=lambda: self.do_export("xlsx")).pack(side="right")
        ttk.Button(bot, text="Export PDF", style="Accent.TButton",
                   command=lambda: self.do_export("pdf")).pack(side="right", padx=6)
        self.rows = []

    def _set(self, a, b):
        self.d1.set(a)
        self.d2.set(b)
        self.refresh()

    def q_today(self):
        self._set(date.today(), date.today())

    def q_month(self):
        self._set(date.today().replace(day=1), date.today())

    def q_last(self):
        last = date.today().replace(day=1) - timedelta(days=1)
        self._set(last.replace(day=1), last)

    def q_year(self):
        self._set(date(date.today().year, 1, 1), date.today())

    def refresh(self):
        cats = ["All"] + sorted({r["name"] for r in self.db.categories()})
        self.cat["values"] = cats
        if not self.cat.get():
            self.cat.set("All")
        self.rows = self.db.txns(self.d1.get(), self.d2.get(), self.ttype.get(), self.cat.get(),
                                 self.search.get().strip(), self.method.get())
        out, ti, te = [], 0, 0
        for r in self.rows:
            inc = r["amount"] if r["type"] == "Income" else 0
            exp = r["amount"] if r["type"] == "Expense" else 0
            ti += inc
            te += exp
            out.append((r["id"], (fmt_date(r["date"]), r["type"], r["category"] or "", r["description"],
                                  r["employee_name"] or r["company_name"] or "", r["method"],
                                  money(inc) if inc else "", money(exp) if exp else "", r["remark"] or ""),
                        ("income",) if inc else ()))
        self.table.load(out)
        self.total_lbl.config(text=f"{len(self.rows)} entries     Income: {money(ti)}     "
                                   f"Expense: {money(te)}     Net: {money(ti - te)}")

    def do_export(self, fmt):
        if not self.rows:
            messagebox.showinfo("Export", "Nothing to export - press Show first.")
            return
        title = f"Report {fmt_date(self.d1.get()) or 'start'} to {fmt_date(self.d2.get()) or 'today'}"
        parts = [f"{k}: {v}" for k, v in (("Type", self.ttype.get()), ("Category", self.cat.get()),
                                          ("Account", self.method.get())) if v and v != "All"]
        if self.search.get().strip():
            parts.append(f"Search: {self.search.get().strip()}")
        self.export(reports.export_txn_report, self.db, self.rows, title, " | ".join(parts) or "All entries", fmt)


# =================================================================== employees
class EmployeesPage(Page):
    title = "Employees"
    BASE = [("emp_code", "ID No", "entry", None), ("iqama_no", "Iqama No", "entry", None),
            ("name", "Full Name *", "entry", None), ("short_name", "Short Name", "entry", None),
            ("basic_salary", "Basic Salary", "entry", None), ("status", "Status", "readonly", EMPLOYEE_STATUSES),
            ("iqama_expiry", "Iqama Expiry", "date", None), ("phone", "Phone", "entry", None),
            ("nationality", "Nationality", "entry", None), ("join_date", "Joining Date", "date", None),
            ("notes", "Notes", "text", None)]

    def __init__(self, master, app):
        super().__init__(master, app)
        self.edit_id = None
        ttk.Button(self.head_right, text="Excel", command=lambda: self.export(
            reports.export_employees, self.db, self.rows, "xlsx")).pack(side="right")
        ttk.Button(self.head_right, text="PDF", command=lambda: self.export(
            reports.export_employees, self.db, self.rows, "pdf")).pack(side="right", padx=6)
        right = card(self.body, width=430)
        right.pack(side="right", fill="y", padx=(10, 0))
        right.pack_propagate(False)
        left = ttk.Frame(self.body)
        left.pack(side="left", fill="both", expand=True)
        flt = ttk.Frame(left)
        flt.pack(fill="x", pady=(0, 6))
        ttk.Label(flt, text="Search").pack(side="left")
        self.search = ttk.Entry(flt, width=26)
        self.search.pack(side="left", padx=4)
        self.search.bind("<KeyRelease>", lambda e: self.refresh())
        ttk.Label(flt, text="Status").pack(side="left", padx=(10, 2))
        self.status_f = ttk.Combobox(flt, values=["All"] + EMPLOYEE_STATUSES, state="readonly", width=11)
        self.status_f.set("Active")
        self.status_f.pack(side="left")
        self.status_f.bind("<<ComboboxSelected>>", lambda e: self.refresh())
        self.count_lbl = ttk.Label(flt, text="")
        self.count_lbl.pack(side="right")
        self.table = Table(left, [])
        self.table.pack(fill="both", expand=True)
        self.table.bind_select(self.load_selected)

        tk.Label(right, text="Employee details", bg=COLORS["card"], font=FONT_B, fg=COLORS["accent"]).pack(
            anchor="w", padx=10, pady=(10, 0))
        self.scroll = ScrollFrame(right)
        self.scroll.pack(fill="both", expand=True, padx=4, pady=4)
        btns = tk.Frame(right, bg=COLORS["card"])
        btns.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Button(btns, text="Save", style="Accent.TButton", command=self.save).pack(side="left")
        ttk.Button(btns, text="New", command=self.new).pack(side="left", padx=6)
        ttk.Button(btns, text="Delete", style="Danger.TButton", command=self.delete).pack(side="left")
        ttk.Button(btns, text="Payment History", command=self.history).pack(side="right")
        self.form = None
        self.rows = []
        self.field_sig = None

    def build_form(self):
        extras = self.db.custom_fields("employee")
        sig = tuple(extras)
        if sig == self.field_sig:
            return
        self.field_sig = sig
        for w in self.scroll.inner.winfo_children():
            w.destroy()
        spec = self.BASE + [(f"x:{n}", n, "entry", None) for n in extras]
        self.form = Form(self.scroll.inner, spec)
        tk.Label(self.scroll.inner, text="Need another field? Settings → Custom Fields.", bg=COLORS["card"],
                 fg=COLORS["muted"], font=("Segoe UI", 8)).grid(row=self.form.next_row, column=0, columnspan=2,
                                                                sticky="w", padx=8, pady=6)
        cols = [("code", "ID No", 70, "center"), ("iqama", "Iqama No", 95, "center"),
                ("name", "Name", 250, "w"), ("sal", "Basic Salary", 95, "e"), ("status", "Status", 85, "center"),
                ("exp", "Iqama Expiry", 95, "center"), ("phone", "Phone", 100, "w")]
        cols += [(f"x{i}", n, 110, "w") for i, n in enumerate(extras)]
        self.table.set_columns(cols)
        self.new()

    def refresh(self):
        self.build_form()
        extras = self.db.custom_fields("employee")
        self.rows = self.db.employees(self.search.get().strip(), self.status_f.get())
        today = date.today().isoformat()
        out = []
        for r in self.rows:
            ex = self.db.extra(r)
            tags = ()
            if r["status"] in ("Removed", "Final Exit"):
                tags = ("muted",)
            elif r["iqama_expiry"] and r["iqama_expiry"] < today:
                tags = ("warn",)
            out.append((r["id"], (r["emp_code"], r["iqama_no"], r["name"], money(r["basic_salary"]), r["status"],
                                  fmt_date(r["iqama_expiry"]), r["phone"] or "") + tuple(ex.get(n, "") for n in extras),
                        tags))
        self.table.load(out)
        total = sum(r["basic_salary"] or 0 for r in self.rows)
        self.count_lbl.config(text=f"{len(self.rows)} employees · total basic {money(total)}")

    def new(self):
        self.edit_id = None
        self.form.clear()
        self.form.set("status", "Active")
        self.form.set("emp_code", self.db.next_code("employees", "emp_code", self._prefix()))
        self.form.set("basic_salary", "0")

    def _prefix(self):
        r = self.db.one("SELECT emp_code FROM employees WHERE emp_code LIKE '%-%' ORDER BY id DESC LIMIT 1")
        return r[0].split("-")[0] if r else "E"

    def load_selected(self):
        eid = self.table.selected()
        if not eid:
            return
        r = self.db.one("SELECT * FROM employees WHERE id=?", (eid,))
        self.edit_id = eid
        ex = self.db.extra(r)
        for key, *_ in self.BASE:
            v = r[key]
            if key == "basic_salary":
                v = f"{v or 0:g}"
            self.form.set(key, v)
        for n in self.db.custom_fields("employee"):
            self.form.set(f"x:{n}", ex.get(n, ""))

    def save(self):
        v = self.form.values()
        if not v["name"]:
            messagebox.showwarning("Employee", "Name is required.")
            return
        try:
            sal = to_float(v["basic_salary"])
        except ValueError as e:
            messagebox.showwarning("Employee", str(e))
            return
        dup = self.db.one("SELECT id FROM employees WHERE emp_code=? AND id!=?", (v["emp_code"], self.edit_id or 0))
        if v["emp_code"] and dup:
            messagebox.showwarning("Employee", f"ID No {v['emp_code']} is already used.")
            return
        data = {k: v[k] for k, *_ in self.BASE}
        data["basic_salary"] = sal
        data["extra"] = {n: v[f"x:{n}"] for n in self.db.custom_fields("employee") if v.get(f"x:{n}")}
        self.edit_id = self.db.save_employee(data, self.edit_id)
        self.app.status(f"Saved employee {data['name']}")
        keep = self.edit_id
        self.refresh()
        if self.table.tree.exists(str(keep)):
            self.table.tree.selection_set(str(keep))
            self.table.tree.see(str(keep))

    def delete(self):
        if not self.edit_id:
            return
        r = self.db.one("SELECT name FROM employees WHERE id=?", (self.edit_id,))
        if messagebox.askyesno("Delete", f"Delete {r['name']} and their salary sheet rows?\n\n"
                                         "Tip: to keep history, set Status to 'Removed' or 'Final Exit' instead."):
            self.db.delete_employee(self.edit_id)
            self.new()
            self.refresh()

    def history(self):
        if not self.edit_id:
            messagebox.showinfo("History", "Select an employee first.")
            return
        e = self.db.one("SELECT * FROM employees WHERE id=?", (self.edit_id,))
        top = tk.Toplevel(self)
        top.title(f"Payment history - {e['name']}")
        top.geometry("900x500")
        ttk.Label(top, text=f"{e['emp_code']}  {e['name']}", font=FONT_B).pack(anchor="w", padx=10, pady=8)
        t = Table(top, [("date", "Date", 100, "center"), ("desc", "Description", 380, "w"),
                        ("cat", "Category", 150, "w"), ("amt", "Amount", 110, "e"), ("src", "Source", 90, "center")])
        t.pack(fill="both", expand=True, padx=10)
        words = [w for w in (e["short_name"] or e["name"]).split() if len(w) > 2][:2]
        like = "%" + "%".join(words) + "%" if words else "%%%"
        rows = self.db.q("SELECT * FROM transactions WHERE employee_id=? OR description LIKE ? ORDER BY date DESC",
                         (self.edit_id, like))
        t.load([(r["id"], (fmt_date(r["date"]), r["description"], r["category"] or "", money(r["amount"]),
                           r["source"]), ()) for r in rows])
        ttk.Label(top, text=f"{len(rows)} entries · total {money(sum(r['amount'] for r in rows))}   "
                            f"(includes old entries whose description mentions '{' '.join(words)}')").pack(
            anchor="w", padx=10, pady=8)


# =================================================================== payroll
class PayrollPage(Page):
    title = "Salaries"

    def __init__(self, master, app):
        super().__init__(master, app)
        top = card(self.body)
        top.pack(fill="x")
        f = tk.Frame(top, bg=COLORS["card"])
        f.pack(fill="x", padx=10, pady=8)
        tk.Label(f, text="Month", bg=COLORS["card"]).pack(side="left")
        months = []
        d = date.today().replace(day=1)
        for i in range(-1, 18):
            m = d.month - i
            y = d.year + (m - 1) // 12
            months.append(date(y, (m - 1) % 12 + 1, 1).strftime("%b-%Y"))
        self.month = ttk.Combobox(f, values=months, width=10)
        self.month.set(date.today().strftime("%b-%Y"))
        self.month.pack(side="left", padx=6)
        self.month.bind("<<ComboboxSelected>>", lambda e: self.refresh())
        self.month.bind("<Return>", lambda e: self.refresh())
        ttk.Button(f, text="Generate / Update Sheet", style="Accent.TButton", command=self.generate).pack(side="left", padx=6)
        ttk.Button(f, text="Pay All Remaining", command=self.pay_all).pack(side="left", padx=6)
        ttk.Button(f, text="Salary Sheet Excel", command=lambda: self.export(
            reports.export_payroll, self.db, self.ym(), "xlsx")).pack(side="right")
        ttk.Button(f, text="Salary Sheet PDF", command=lambda: self.export(
            reports.export_payroll, self.db, self.ym(), "pdf")).pack(side="right", padx=6)
        ttk.Button(f, text="Payslip PDF", command=self.payslip).pack(side="right")

        self.table = Table(self.body, [("code", "ID", 60, "center"), ("name", "Name", 230, "w"),
                                       ("basic", "Basic", 85, "e"), ("allow", "Allow.", 75, "e"),
                                       ("ot", "Overtime", 75, "e"), ("ded", "Deduct.", 75, "e"),
                                       ("adv", "Advance", 75, "e"), ("net", "Net", 90, "e"),
                                       ("paid", "Paid", 90, "e"), ("bal", "Balance", 90, "e"),
                                       ("pdate", "Paid Date", 95, "center"), ("status", "Status", 75, "center"),
                                       ("remark", "Remark", 120, "w")], height=13)
        self.table.pack(fill="both", expand=True, pady=10)
        self.table.bind_select(self.load_selected)

        bot = card(self.body)
        bot.pack(fill="x")
        g = tk.Frame(bot, bg=COLORS["card"])
        g.pack(fill="x", padx=10, pady=8)
        self.sel_lbl = tk.Label(g, text="Select an employee row above", bg=COLORS["card"], font=FONT_B,
                                fg=COLORS["accent"])
        self.sel_lbl.grid(row=0, column=0, columnspan=10, sticky="w")
        self.fields = {}
        for i, (k, lbl) in enumerate((("basic_salary", "Basic"), ("allowance", "Allowance"), ("overtime", "Overtime"),
                                      ("deduction", "Deduction"), ("advance", "Advance"))):
            tk.Label(g, text=lbl, bg=COLORS["card"]).grid(row=1, column=i * 2, sticky="e", padx=(8, 2), pady=4)
            e = ttk.Entry(g, width=10)
            e.grid(row=1, column=i * 2 + 1, sticky="w")
            self.fields[k] = e
        tk.Label(g, text="Remark", bg=COLORS["card"]).grid(row=1, column=10, sticky="e", padx=(8, 2))
        self.remark = ttk.Entry(g, width=20)
        self.remark.grid(row=1, column=11, sticky="w")

        tk.Label(g, text="Pay amount", bg=COLORS["card"]).grid(row=2, column=0, sticky="e", padx=(8, 2), pady=4)
        self.pay_amt = ttk.Entry(g, width=10)
        self.pay_amt.grid(row=2, column=1, sticky="w")
        tk.Label(g, text="Date", bg=COLORS["card"]).grid(row=2, column=2, sticky="e", padx=(8, 2))
        self.pay_date = DateEntry(g, style="Card.TFrame")
        self.pay_date.grid(row=2, column=3, columnspan=2, sticky="w")
        tk.Label(g, text="Via", bg=COLORS["card"]).grid(row=2, column=5, sticky="e", padx=(8, 2))
        self.pay_method = ttk.Combobox(g, values=PAY_METHODS, state="readonly", width=7)
        self.pay_method.set("Bank")
        self.pay_method.grid(row=2, column=6, sticky="w")
        self.fee_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(g, text="+ Transfer fee & VAT", variable=self.fee_var).grid(row=2, column=7, columnspan=3,
                                                                                   sticky="w", padx=8)
        pb = tk.Frame(g, bg=COLORS["card"])
        pb.grid(row=3, column=0, columnspan=12, sticky="w", padx=8, pady=(6, 0))
        ttk.Button(pb, text="Save Changes", command=self.save_row).pack(side="left")
        ttk.Button(pb, text="Record Payment", style="Accent.TButton", command=self.pay).pack(side="left", padx=8)
        tk.Label(pb, text="Payments are added to the Daily Book automatically (category Salary).",
                 bg=COLORS["card"], fg=COLORS["muted"]).pack(side="left", padx=8)
        self.total_lbl = ttk.Label(self.body, text="", font=FONT_B)
        self.total_lbl.pack(anchor="w", pady=(6, 0))
        self.sel = None

    def ym(self):
        try:
            return month_from_label(self.month.get().strip())
        except ValueError:
            messagebox.showwarning("Month", "Month must look like Oct-2026")
            raise

    def generate(self):
        n = self.db.generate_payroll(self.ym())
        # refresh basic salary for unpaid rows from the employee record
        for r in self.db.payroll(self.ym()):
            if not r["paid_amount"]:
                e = self.db.one("SELECT basic_salary FROM employees WHERE id=?", (r["employee_id"],))
                self.db.run("UPDATE payroll SET basic_salary=? WHERE id=?", (e["basic_salary"] or 0, r["id"]))
        self.app.status(f"{n} employee(s) added to the {self.month.get()} salary sheet")
        self.refresh()

    def refresh(self):
        try:
            ym = self.ym()
        except ValueError:
            return
        rows = self.db.payroll(ym)
        out = []
        tot = {"net": 0, "paid": 0}
        for r in rows:
            net = self.db.net_salary(r)
            paid = r["paid_amount"] or 0
            bal = net - paid
            status = "Paid" if bal <= 0.009 else ("Partial" if paid else "Unpaid")
            tot["net"] += net
            tot["paid"] += paid
            out.append((r["id"], (r["emp_code"], r["name"], money(r["basic_salary"]), money(r["allowance"]),
                                  money(r["overtime"]), money(r["deduction"]), money(r["advance"]), money(net),
                                  money(paid), money(bal), fmt_date(r["paid_date"]), status, r["remark"] or ""),
                        ("income",) if status == "Paid" else (("warn",) if status == "Unpaid" else ())))
        self.table.load(out)
        if not rows:
            self.total_lbl.config(text="No salary sheet for this month yet - press 'Generate / Update Sheet'.")
        else:
            self.total_lbl.config(text=f"{len(rows)} employees     Net total: {money(tot['net'])}     "
                                       f"Paid: {money(tot['paid'])}     Remaining: {money(tot['net'] - tot['paid'])}")
        if self.sel and self.table.tree.exists(str(self.sel)):
            self.table.tree.selection_set(str(self.sel))

    def load_selected(self):
        pid = self.table.selected()
        if not pid:
            return
        self.sel = pid
        r = self.db.one("SELECT p.*, e.name, e.emp_code FROM payroll p JOIN employees e ON e.id=p.employee_id "
                        "WHERE p.id=?", (pid,))
        self.sel_lbl.config(text=f"{r['emp_code']}  {r['name']}  -  {month_label(r['month'])}")
        for k, e in self.fields.items():
            e.delete(0, "end")
            e.insert(0, f"{r[k] or 0:g}")
        self.remark.delete(0, "end")
        self.remark.insert(0, r["remark"] or "")
        bal = self.db.net_salary(r) - (r["paid_amount"] or 0)
        self.pay_amt.delete(0, "end")
        self.pay_amt.insert(0, f"{max(bal, 0):g}")

    def save_row(self):
        if not self.sel:
            return
        try:
            data = {k: to_float(e.get()) for k, e in self.fields.items()}
        except ValueError as e:
            messagebox.showwarning("Salary", str(e))
            return
        data["remark"] = self.remark.get().strip()
        self.db.save_payroll(data, self.sel)
        self.refresh()
        self.load_selected()

    def _pay(self, pid, amount, day, method, fee):
        r = self.db.one("SELECT p.*, e.name FROM payroll p JOIN employees e ON e.id=p.employee_id WHERE p.id=?", (pid,))
        self.db.save_txn({"date": day, "type": "Expense", "category": "Salary",
                          "description": f"{r['name']} Salary {month_label(r['month'])}", "amount": amount,
                          "method": method, "employee_id": r["employee_id"], "source": "payroll", "ref_id": pid})
        self.db.run("UPDATE payroll SET paid_amount=COALESCE(paid_amount,0)+?, paid_date=?, method=? WHERE id=?",
                    (amount, day, method, pid))
        if fee and method == "Bank":
            add_fee(self.db, day, to_float(self.db.get("transfer_fee")), method)

    def pay(self):
        if not self.sel:
            messagebox.showinfo("Pay", "Select an employee row first.")
            return
        try:
            amt = to_float(self.pay_amt.get())
        except ValueError as e:
            messagebox.showwarning("Pay", str(e))
            return
        day = self.pay_date.get()
        if amt <= 0 or not day:
            messagebox.showwarning("Pay", "Enter an amount and a valid date.")
            return
        self.save_row()
        self._pay(self.sel, amt, day, self.pay_method.get(), self.fee_var.get())
        self.app.status(f"Salary payment of {money(amt)} recorded and added to the daily book")
        self.refresh()
        self.load_selected()

    def pay_all(self):
        rows = [r for r in self.db.payroll(self.ym()) if self.db.net_salary(r) - (r["paid_amount"] or 0) > 0.009]
        if not rows:
            messagebox.showinfo("Pay all", "Nothing remaining to pay for this month.")
            return
        total = sum(self.db.net_salary(r) - (r["paid_amount"] or 0) for r in rows)
        day = self.pay_date.get() or date.today().isoformat()
        if not messagebox.askyesno("Pay all", f"Record payment for {len(rows)} employee(s), total {money(total)}, "
                                              f"on {fmt_date(day)} via {self.pay_method.get()}?\n"
                                              f"Each payment is added to the daily book"
                                              f"{' with transfer fee + VAT' if self.fee_var.get() else ''}."):
            return
        for r in rows:
            self._pay(r["id"], self.db.net_salary(r) - (r["paid_amount"] or 0), day, self.pay_method.get(),
                      self.fee_var.get())
        self.refresh()

    def payslip(self):
        if not self.sel:
            messagebox.showinfo("Payslip", "Select an employee row first.")
            return
        self.export(reports.export_payslip, self.db, self.sel)


# =================================================================== clients
class ClientsPage(Page):
    title = "Clients & Invoices"
    BASE = [("code", "ID No", "entry", None), ("name", "Company Name *", "entry", None),
            ("category", "Category", "readonly", COMPANY_CATEGORIES), ("contact_person", "Contact Person", "entry", None),
            ("phone", "Phone", "entry", None), ("vat_no", "VAT / CR No", "entry", None), ("notes", "Notes", "text", None)]

    def __init__(self, master, app):
        super().__init__(master, app)
        self.edit_id = None
        self.inv_id = None
        ttk.Button(self.head_right, text="Client List Excel", command=lambda: self.export(
            reports.export_companies, self.db, self.rows, "xlsx")).pack(side="right")
        ttk.Button(self.head_right, text="Client List PDF", command=lambda: self.export(
            reports.export_companies, self.db, self.rows, "pdf")).pack(side="right", padx=6)
        left = ttk.Frame(self.body, width=500)
        left.pack(side="left", fill="y", padx=(0, 10))
        left.pack_propagate(False)
        flt = ttk.Frame(left)
        flt.pack(fill="x", pady=(0, 6))
        ttk.Label(flt, text="Search").pack(side="left")
        self.search = ttk.Entry(flt, width=20)
        self.search.pack(side="left", padx=4)
        self.search.bind("<KeyRelease>", lambda e: self.refresh())
        self.cat_f = ttk.Combobox(flt, values=["All"] + COMPANY_CATEGORIES, state="readonly", width=14)
        self.cat_f.set("All")
        self.cat_f.pack(side="left", padx=4)
        self.cat_f.bind("<<ComboboxSelected>>", lambda e: self.refresh())
        self.table = Table(left, [], height=10)
        self.table.pack(fill="both", expand=True)
        self.table.bind_select(self.load_selected)
        fc = card(left)
        fc.pack(fill="x", pady=(8, 0))
        self.form_holder = tk.Frame(fc, bg=COLORS["card"])
        self.form_holder.pack(fill="x", pady=6)
        b = tk.Frame(fc, bg=COLORS["card"])
        b.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Button(b, text="Save Client", style="Accent.TButton", command=self.save).pack(side="left")
        ttk.Button(b, text="New", command=self.new).pack(side="left", padx=6)
        ttk.Button(b, text="Delete", style="Danger.TButton", command=self.delete).pack(side="left")

        right = ttk.Frame(self.body)
        right.pack(side="left", fill="both", expand=True)
        ih = ttk.Frame(right)
        ih.pack(fill="x", pady=(0, 6))
        self.inv_title = ttk.Label(ih, text="Invoices - select a client", font=FONT_B)
        self.inv_title.pack(side="left")
        ttk.Button(ih, text="Statement Excel", command=lambda: self.statement("xlsx")).pack(side="right")
        ttk.Button(ih, text="Statement PDF", style="Accent.TButton",
                   command=lambda: self.statement("pdf")).pack(side="right", padx=6)
        self.inv_table = Table(right, [("no", "Invoice #", 100, "w"), ("date", "Date", 90, "center"),
                                       ("period", "Month / Period", 140, "w"), ("desc", "Description", 170, "w"),
                                       ("amt", "Amount", 95, "e"), ("rec", "Received", 95, "e"),
                                       ("bal", "Balance", 95, "e")], height=8)
        self.inv_table.pack(fill="both", expand=True)
        self.inv_table.bind_select(self.load_invoice)
        self.inv_total = ttk.Label(right, text="", font=FONT_B)
        self.inv_total.pack(anchor="w", pady=4)
        ic = card(right)
        ic.pack(fill="x")
        g = tk.Frame(ic, bg=COLORS["card"])
        g.pack(fill="x", pady=6)
        self.inv_form = Form(g, [("invoice_no", "Invoice #", "entry", None), ("issue_date", "Invoice Date", "date", None),
                                 ("period", "Month / Period", "entry", None), ("amount", "Amount", "entry", None),
                                 ("description", "Description", "entry", None), ("remark", "Remark", "entry", None)],
                             columns=2, label_width=12)
        b2 = tk.Frame(ic, bg=COLORS["card"])
        b2.pack(fill="x", padx=8, pady=(0, 6))
        ttk.Button(b2, text="Save Invoice", style="Accent.TButton", command=self.save_invoice).pack(side="left")
        ttk.Button(b2, text="New Invoice", command=self.new_invoice).pack(side="left", padx=6)
        ttk.Button(b2, text="Delete Invoice", style="Danger.TButton", command=self.delete_invoice).pack(side="left")
        b3 = tk.Frame(ic, bg=COLORS["card"])
        b3.pack(fill="x", padx=8, pady=(0, 8))
        tk.Label(b3, text="Receive", bg=COLORS["card"]).pack(side="left")
        self.rec_amt = ttk.Entry(b3, width=12)
        self.rec_amt.pack(side="left", padx=4)
        tk.Label(b3, text="on", bg=COLORS["card"]).pack(side="left")
        self.rec_date = DateEntry(b3, style="Card.TFrame")
        self.rec_date.pack(side="left", padx=4)
        self.rec_method = ttk.Combobox(b3, values=PAY_METHODS, state="readonly", width=6)
        self.rec_method.set("Bank")
        self.rec_method.pack(side="left", padx=4)
        ttk.Button(b3, text="Record Receipt", command=self.receive).pack(side="left", padx=4)
        self.rows = []
        self.form = None
        self.field_sig = None

    def build_form(self):
        extras = self.db.custom_fields("company")
        if tuple(extras) == self.field_sig:
            return
        self.field_sig = tuple(extras)
        for w in self.form_holder.winfo_children():
            w.destroy()
        self.form = Form(self.form_holder, self.BASE[:-1] + [(f"x:{n}", n, "entry", None) for n in extras],
                         columns=1, label_width=14)
        cols = [("code", "ID", 50, "center"), ("name", "Company", 210, "w"), ("cat", "Category", 110, "w"),
                ("due", "Receivable", 95, "e")] + [(f"x{i}", n, 100, "w") for i, n in enumerate(extras)]
        self.table.set_columns(cols)
        self.new()

    def refresh(self):
        self.build_form()
        extras = self.db.custom_fields("company")
        self.rows = self.db.companies(self.search.get().strip(), self.cat_f.get())
        dues = {r["company_id"]: r["due"] for r in self.db.q(
            "SELECT company_id, SUM(amount - (SELECT COALESCE(SUM(amount),0) FROM receipts WHERE invoice_id=i.id)) due "
            "FROM invoices i GROUP BY company_id")}
        out = []
        for r in self.rows:
            ex = self.db.extra(r)
            due = dues.get(r["id"], 0) or 0
            out.append((r["id"], (r["code"], r["name"], r["category"], money(due) if due else "")
                        + tuple(ex.get(n, "") for n in extras),
                        ("muted",) if r["category"] == "Previous Client" else ()))
        self.table.load(out)
        if self.edit_id and self.table.tree.exists(str(self.edit_id)):
            self.table.tree.selection_set(str(self.edit_id))
        self.load_invoices()

    def new(self):
        self.edit_id = None
        self.form.clear()
        self.form.set("code", self.db.next_code("companies", "code", "C"))
        self.form.set("category", "Current Client")
        self.load_invoices()

    def load_selected(self):
        cid = self.table.selected()
        if not cid or cid == self.edit_id:
            return
        self.edit_id = cid
        r = self.db.one("SELECT * FROM companies WHERE id=?", (cid,))
        ex = self.db.extra(r)
        for key, *_ in self.BASE[:-1]:
            self.form.set(key, r[key])
        for n in self.db.custom_fields("company"):
            self.form.set(f"x:{n}", ex.get(n, ""))
        self.new_invoice()
        self.load_invoices()

    def save(self):
        v = self.form.values()
        if not v["name"]:
            messagebox.showwarning("Client", "Company name is required.")
            return
        data = {k: v[k] for k, *_ in self.BASE[:-1]}
        data["extra"] = {n: v[f"x:{n}"] for n in self.db.custom_fields("company") if v.get(f"x:{n}")}
        self.edit_id = self.db.save_company(data, self.edit_id)
        self.refresh()

    def delete(self):
        if not self.edit_id:
            return
        r = self.db.one("SELECT name FROM companies WHERE id=?", (self.edit_id,))
        if messagebox.askyesno("Delete", f"Delete {r['name']} with all its invoices and receipts?\n\n"
                                         "Tip: set Category to 'Previous Client' to keep the history."):
            self.db.delete_company(self.edit_id)
            self.edit_id = None
            self.refresh()
            self.new()

    # invoices
    def load_invoices(self):
        if not self.edit_id:
            self.inv_table.load([])
            self.inv_title.config(text="Invoices - select a client on the left")
            self.inv_total.config(text="")
            return
        c = self.db.one("SELECT name FROM companies WHERE id=?", (self.edit_id,))
        self.inv_title.config(text=f"Invoices - {c['name']}")
        invs = self.db.invoices(self.edit_id)
        self.inv_table.load([(r["id"], (r["invoice_no"], fmt_date(r["issue_date"]), r["period"], r["description"],
                                        money(r["amount"]), money(r["received"]), money(r["amount"] - r["received"])),
                              ("income",) if r["amount"] - r["received"] <= 0.009 else ()) for r in invs])
        ta = sum(r["amount"] for r in invs)
        tr = sum(r["received"] for r in invs)
        self.inv_total.config(text=f"Total {money(ta)}     Received {money(tr)}     Remaining {money(ta - tr)}")

    def new_invoice(self):
        self.inv_id = None
        self.inv_form.clear()
        self.inv_form.set("issue_date", date.today().isoformat())
        self.rec_amt.delete(0, "end")

    def load_invoice(self):
        iid = self.inv_table.selected()
        if not iid:
            return
        self.inv_id = iid
        r = self.db.one("SELECT i.*, (SELECT COALESCE(SUM(amount),0) FROM receipts WHERE invoice_id=i.id) rec "
                        "FROM invoices i WHERE id=?", (iid,))
        for k in ("invoice_no", "issue_date", "period", "description", "remark"):
            self.inv_form.set(k, r[k])
        self.inv_form.set("amount", f"{r['amount']:g}")
        self.rec_amt.delete(0, "end")
        self.rec_amt.insert(0, f"{max(r['amount'] - r['rec'], 0):g}")

    def save_invoice(self):
        if not self.edit_id:
            messagebox.showinfo("Invoice", "Select or save a client first.")
            return
        v = self.inv_form.values()
        try:
            v["amount"] = to_float(v["amount"])
        except ValueError as e:
            messagebox.showwarning("Invoice", str(e))
            return
        v["company_id"] = self.edit_id
        self.inv_id = self.db.save_invoice(v, self.inv_id)
        self.refresh()

    def delete_invoice(self):
        if self.inv_id and messagebox.askyesno("Delete", "Delete this invoice? Receipts recorded for it are "
                                                         "also removed from the daily book."):
            self.db.delete_invoice(self.inv_id)
            self.new_invoice()
            self.refresh()

    def receive(self):
        if not self.inv_id:
            messagebox.showinfo("Receipt", "Select an invoice first.")
            return
        try:
            amt = to_float(self.rec_amt.get())
        except ValueError as e:
            messagebox.showwarning("Receipt", str(e))
            return
        day = self.rec_date.get()
        if amt <= 0 or not day:
            messagebox.showwarning("Receipt", "Enter an amount and date.")
            return
        inv = self.db.one("SELECT i.*, c.name cname FROM invoices i JOIN companies c ON c.id=i.company_id WHERE i.id=?",
                          (self.inv_id,))
        tid = self.db.save_txn({"date": day, "type": "Income", "category": "Client Payment",
                                "description": f"{inv['cname']} payment {inv['invoice_no'] or ''}".strip(),
                                "amount": amt, "method": self.rec_method.get(), "company_id": inv["company_id"],
                                "source": "invoice", "ref_id": self.inv_id})
        self.db.run("INSERT INTO receipts(invoice_id,date,amount,txn_id) VALUES (?,?,?,?)", (self.inv_id, day, amt, tid))
        self.app.status(f"Receipt {money(amt)} recorded and added to the daily book as income")
        self.refresh()

    def statement(self, fmt):
        if not self.edit_id:
            messagebox.showinfo("Statement", "Select a client first.")
            return
        self.export(reports.export_statement, self.db, self.edit_id, fmt)


# =================================================================== settings
class SettingsPage(Page):
    title = "Settings & Data"

    def __init__(self, master, app):
        super().__init__(master, app)
        sf = ScrollFrame(self.body, bg=COLORS["bg"])
        sf.pack(fill="both", expand=True)
        root = sf.inner
        root.configure(style="TFrame")

        c1 = ttk.LabelFrame(root, text="Company & report details")
        c1.pack(fill="x", pady=(0, 10))
        g = ttk.Frame(c1, style="Card.TFrame")
        g.pack(fill="x", padx=6, pady=6)
        self.form = Form(g, [("company_name", "Company Name", "entry", None), ("currency", "Currency", "entry", None),
                             ("report_title", "Daily Report Title", "entry", None),
                             ("report_subtitle", "Bank Report Subtitle", "entry", None),
                             ("opening_balance", "Bank Opening Balance", "entry", None),
                             ("opening_date", "Opening Balance Date", "date", None),
                             ("cash_opening", "Cash Opening Balance", "entry", None),
                             ("transfer_fee", "Default Transfer Fee", "entry", None),
                             ("vat_percent", "VAT % on fee", "entry", None)], columns=2, label_width=20)
        ttk.Button(g, text="Save Settings", style="Accent.TButton", command=self.save).grid(
            row=self.form.next_row, column=1, sticky="w", padx=0, pady=6)

        c2 = ttk.LabelFrame(root, text="Custom fields (add your own fields - they appear on the forms and in exports)")
        c2.pack(fill="x", pady=(0, 10))
        g2 = ttk.Frame(c2, style="Card.TFrame")
        g2.pack(fill="x", padx=6, pady=6)
        ttk.Label(g2, text="Add to", style="Card.TLabel").grid(row=0, column=0, padx=4)
        self.cf_entity = ttk.Combobox(g2, values=["Employees", "Clients", "Daily entries"], state="readonly", width=14)
        self.cf_entity.set("Employees")
        self.cf_entity.grid(row=0, column=1, padx=4)
        self.cf_entity.bind("<<ComboboxSelected>>", lambda e: self.load_fields())
        ttk.Label(g2, text="Field name", style="Card.TLabel").grid(row=0, column=2, padx=4)
        self.cf_name = ttk.Entry(g2, width=24)
        self.cf_name.grid(row=0, column=3, padx=4)
        ttk.Button(g2, text="Add Field", command=self.add_field).grid(row=0, column=4, padx=4)
        ttk.Button(g2, text="Remove Selected", style="Danger.TButton", command=self.del_field).grid(row=0, column=5, padx=4)
        self.cf_list = tk.Listbox(g2, height=4, font=FONT)
        self.cf_list.grid(row=1, column=0, columnspan=6, sticky="ew", padx=4, pady=6)
        ttk.Label(g2, text="Examples: Passport No, Bank Account / IBAN, Visa Expiry, Vehicle No, Project, Invoice No",
                  style="Muted.TLabel").grid(row=2, column=0, columnspan=6, sticky="w", padx=4)

        c3 = ttk.LabelFrame(root, text="Expense / income categories")
        c3.pack(fill="x", pady=(0, 10))
        g3 = ttk.Frame(c3, style="Card.TFrame")
        g3.pack(fill="x", padx=6, pady=6)
        self.cat_type = ttk.Combobox(g3, values=TXN_TYPES, state="readonly", width=9)
        self.cat_type.set("Expense")
        self.cat_type.grid(row=0, column=0, padx=4)
        self.cat_type.bind("<<ComboboxSelected>>", lambda e: self.load_cats())
        self.cat_name = ttk.Entry(g3, width=28)
        self.cat_name.grid(row=0, column=1, padx=4)
        ttk.Button(g3, text="Add Category", command=self.add_cat).grid(row=0, column=2, padx=4)
        ttk.Button(g3, text="Remove Selected", style="Danger.TButton", command=self.del_cat).grid(row=0, column=3, padx=4)
        self.cat_list = tk.Listbox(g3, height=6, font=FONT)
        self.cat_list.grid(row=1, column=0, columnspan=4, sticky="ew", padx=4, pady=6)
        g3.columnconfigure(1, weight=1)

        c4 = ttk.LabelFrame(root, text="Import your existing Excel files / Backup")
        c4.pack(fill="x", pady=(0, 10))
        g4 = ttk.Frame(c4, style="Card.TFrame")
        g4.pack(fill="x", padx=6, pady=6)
        ttk.Button(g4, text="Import Excel File...", style="Accent.TButton", command=self.import_file).pack(side="left", padx=4)
        ttk.Button(g4, text="Backup Now", command=self.backup).pack(side="left", padx=4)
        ttk.Button(g4, text="Restore Backup...", command=self.restore).pack(side="left", padx=4)
        ttk.Button(g4, text="Open Exports Folder", command=lambda: (os.makedirs(EXPORT_DIR, exist_ok=True),
                                                                    open_file(EXPORT_DIR))).pack(side="left", padx=4)
        ttk.Button(g4, text="Open Data Folder", command=lambda: open_file(DATA_DIR)).pack(side="left", padx=4)
        ttk.Label(c4, text="Import understands: Nominal Roll (employees), ID Wise Companies / Client List, the daily "
                           "Bank Expense workbook (all day sheets), and client statements (invoice sheets).\n"
                           "All data is kept in data\\office.db - copy that one file to move to another computer. "
                           "A backup is made automatically every time the program closes.",
                  style="Muted.TLabel", justify="left").pack(anchor="w", padx=10, pady=(0, 8))

    def refresh(self):
        for k in self.form.widgets:
            self.form.set(k, self.db.get(k, "0" if k == "cash_opening" else ""))
        self.load_fields()
        self.load_cats()

    def save(self):
        v = self.form.values()
        try:
            for k in ("opening_balance", "cash_opening", "transfer_fee", "vat_percent"):
                v[k] = f"{to_float(v[k]):g}"
        except ValueError as e:
            messagebox.showwarning("Settings", str(e))
            return
        for k, val in v.items():
            self.db.set(k, val or "")
        self.app.status("Settings saved")
        self.app.title(f"{APP_NAME} - {v['company_name']}")

    def _entity(self):
        return {"Employees": "employee", "Clients": "company", "Daily entries": "transaction"}[self.cf_entity.get()]

    def load_fields(self):
        self.cf_list.delete(0, "end")
        for n in self.db.custom_fields(self._entity()):
            self.cf_list.insert("end", n)

    def add_field(self):
        n = self.cf_name.get().strip()
        if n:
            self.db.add_custom_field(self._entity(), n)
            self.cf_name.delete(0, "end")
            self.load_fields()

    def del_field(self):
        sel = self.cf_list.curselection()
        if sel and messagebox.askyesno("Remove field", "Remove this field from the forms? "
                                                       "(Values already typed stay saved and come back if re-added.)"):
            self.db.delete_custom_field(self._entity(), self.cf_list.get(sel[0]))
            self.load_fields()

    def load_cats(self):
        self.cat_list.delete(0, "end")
        for n in self.db.categories(self.cat_type.get()):
            self.cat_list.insert("end", n)

    def add_cat(self):
        n = self.cat_name.get().strip()
        if n:
            self.db.run("INSERT OR IGNORE INTO categories(name,type) VALUES (?,?)", (n, self.cat_type.get()))
            self.cat_name.delete(0, "end")
            self.load_cats()

    def del_cat(self):
        sel = self.cat_list.curselection()
        if sel:
            self.db.run("DELETE FROM categories WHERE name=? AND type=?", (self.cat_list.get(sel[0]), self.cat_type.get()))
            self.load_cats()

    def backup(self):
        p = self.db.backup()
        messagebox.showinfo("Backup", f"Backup saved:\n{p}")

    def restore(self):
        p = filedialog.askopenfilename(title="Choose backup", initialdir=BACKUP_DIR,
                                       filetypes=[("Backup", "*.db"), ("All", "*.*")])
        if p and messagebox.askyesno("Restore", "Replace ALL current data with this backup?"):
            self.db.backup()
            self.db.restore(p)
            messagebox.showinfo("Restore", "Backup restored.")
            self.app.refresh_all()

    def import_file(self):
        path = filedialog.askopenfilename(title="Choose Excel file", filetypes=[("Excel", "*.xlsx *.xlsm"), ("All", "*.*")])
        if path:
            ImportDialog(self, self.app, path)


class ImportDialog(tk.Toplevel):
    KINDS = {"employees": "Employees (Nominal Roll)", "companies": "Companies / Client List",
             "bank": "Daily Bank Expense workbook", "statement": "Client statement (invoices)"}

    def __init__(self, master, app, path):
        super().__init__(master)
        self.app, self.path = app, path
        self.title("Import Excel")
        self.geometry("620x330")
        self.transient(app)
        self.app.config(cursor="watch")
        self.update()
        try:
            kind = importer.detect_kind(path)
            self.sheets = importer.sheet_names(path)
        except Exception as e:  # noqa
            self.app.config(cursor="")
            messagebox.showerror("Import", f"Could not read file:\n{e}")
            self.destroy()
            return
        self.app.config(cursor="")
        f = ttk.Frame(self, padding=14)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text=os.path.basename(path), font=FONT_B).grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(f, text="This file contains").grid(row=1, column=0, sticky="w", pady=6)
        self.kind = ttk.Combobox(f, values=list(self.KINDS.values()), state="readonly", width=34)
        self.kind.set(self.KINDS.get(kind, self.KINDS["employees"]))
        self.kind.grid(row=1, column=1, sticky="w")
        self.kind.bind("<<ComboboxSelected>>", lambda e: self.update_ui())
        ttk.Label(f, text="Sheet").grid(row=2, column=0, sticky="w", pady=6)
        self.sheet = ttk.Combobox(f, values=self.sheets, state="readonly", width=34)
        self.sheet.set(self.sheets[0])
        self.sheet.grid(row=2, column=1, sticky="w")
        self.sheet.bind("<<ComboboxSelected>>", lambda e: self.update_ui())
        self.sal_lbl = ttk.Label(f, text="Salary column")
        self.sal_lbl.grid(row=3, column=0, sticky="w", pady=6)
        self.salary = ttk.Combobox(f, state="readonly", width=34)
        self.salary.grid(row=3, column=1, sticky="w")
        self.note = ttk.Label(f, text="", foreground=COLORS["muted"], wraplength=560, justify="left")
        self.note.grid(row=4, column=0, columnspan=2, sticky="w", pady=8)
        ttk.Button(f, text="Import", style="Accent.TButton", command=self.run).grid(row=5, column=1, sticky="e")
        self.update_ui()
        self.grab_set()

    def _kind(self):
        return {v: k for k, v in self.KINDS.items()}[self.kind.get()]

    def update_ui(self):
        k = self._kind()
        self.sheet.config(state="disabled" if k == "bank" else "readonly")
        if k == "employees":
            try:
                heads = importer.header_names(self.path, self.sheet.get())
            except Exception:  # noqa
                heads = []
            nums = [h for h in heads if any(w in h for w in ("salary", "bs", "by", "amount", "pay"))]
            self.salary["values"] = ["(auto)"] + heads
            self.salary.set(nums[-1] if nums else "(auto)")
            self.salary.grid()
            self.sal_lbl.grid()
            self.note.config(text="Existing employees are matched by Iqama No (or ID No) and updated; new ones are "
                                  "added. Rows marked 'Remove' or 'Final exit' get that status.")
        else:
            self.salary.grid_remove()
            self.sal_lbl.grid_remove()
            self.note.config(text={
                "companies": "Companies are matched by name; ID No and Category are filled in when the sheet has them.",
                "bank": "Every dated sheet becomes daily-book entries (expenses + the income lines above the table). "
                        "Sheets already imported are skipped. The opening balance is taken from the first sheet, and "
                        "every day's balance is checked against the old file.",
                "statement": "Creates the client (if new) and adds each invoice row with its received amount.",
            }[k])

    def run(self):
        k = self._kind()
        self.config(cursor="watch")
        self.update()
        try:
            if k == "employees":
                sal = self.salary.get()
                msg = importer.import_employees(self.app.db, self.path, self.sheet.get(),
                                                None if sal in ("", "(auto)") else sal)
            elif k == "companies":
                msg = importer.import_companies(self.app.db, self.path, self.sheet.get())
            elif k == "bank":
                msg = importer.import_bank_workbook(self.app.db, self.path)
            else:
                msg = importer.import_statement(self.app.db, self.path, self.sheet.get())
        except Exception as e:  # noqa
            self.config(cursor="")
            messagebox.showerror("Import failed", str(e), parent=self)
            return
        self.destroy()
        self.app.refresh_all()
        top = tk.Toplevel(self.app)
        top.title("Import finished")
        top.geometry("640x360")
        t = tk.Text(top, wrap="word", font=FONT)
        t.pack(fill="both", expand=True, padx=10, pady=10)
        t.insert("1.0", msg)
        t.config(state="disabled")


# =================================================================== main window
class App(tk.Tk):
    PAGES = [("dashboard", "Dashboard", DashboardPage), ("daily", "Daily Book", DailyPage),
             ("payroll", "Salaries", PayrollPage), ("employees", "Employees", EmployeesPage),
             ("clients", "Clients & Invoices", ClientsPage), ("reports", "Reports", ReportsPage),
             ("settings", "Settings & Import", SettingsPage)]

    def __init__(self):
        super().__init__()
        self.db = Database()
        self.title(f"{APP_NAME} - {self.db.get('company_name')}")
        self.geometry("1360x800")
        self.minsize(1100, 680)
        try:
            self.state("zoomed")
        except tk.TclError:
            pass
        setup_style(self)
        side = tk.Frame(self, bg=COLORS["side"], width=200)
        side.pack(side="left", fill="y")
        side.pack_propagate(False)
        tk.Label(side, text="Office\nManager", bg=COLORS["side"], fg="white", font=("Segoe UI", 15, "bold"),
                 justify="left").pack(anchor="w", padx=18, pady=(20, 4))
        self.comp_lbl = tk.Label(side, text=self.db.get("company_name"), bg=COLORS["side"], fg="#AFC3DB",
                                 font=("Segoe UI", 8), wraplength=170, justify="left")
        self.comp_lbl.pack(anchor="w", padx=18, pady=(0, 18))
        self.nav = {}
        for key, label, _ in self.PAGES:
            b = tk.Label(side, text="   " + label, bg=COLORS["side"], fg="white", anchor="w",
                         font=("Segoe UI", 11), padx=10, pady=10, cursor="hand2")
            b.pack(fill="x")
            b.bind("<Button-1>", lambda e, k=key: self.show(k))
            self.nav[key] = b
        tk.Label(side, text=date.today().strftime("%A\n%d %B %Y"), bg=COLORS["side"], fg="#AFC3DB",
                 font=("Segoe UI", 9), justify="left").pack(side="bottom", anchor="w", padx=18, pady=16)

        right = ttk.Frame(self)
        right.pack(side="left", fill="both", expand=True)
        self.container = ttk.Frame(right)
        self.container.pack(fill="both", expand=True)
        self.status_var = tk.StringVar(value="Ready")
        tk.Label(right, textvariable=self.status_var, anchor="w", bg="#E8EEF6", fg=COLORS["muted"],
                 font=("Segoe UI", 9), padx=10).pack(fill="x", side="bottom")
        self.pages = {}
        for key, _, cls in self.PAGES:
            p = cls(self.container, self)
            p.place(relwidth=1, relheight=1)
            self.pages[key] = p
        self.current = None
        self.bind_all("<Control-s>", lambda e: self.current == "daily" and self.pages["daily"].save())
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        first_run = not self.db.scalar("SELECT COUNT(*) FROM employees") and not self.db.scalar(
            "SELECT COUNT(*) FROM transactions")
        self.show("settings" if first_run else "dashboard")
        if first_run:
            self.after(400, lambda: messagebox.showinfo(
                "Welcome", "Welcome! Start by importing your existing Excel files:\n\n"
                           "Settings & Import → 'Import Excel File...'\n\n"
                           "Import the Nominal Roll, the Companies list and the Bank Expense workbook - "
                           "then the dashboard fills itself."))

    def show(self, key):
        for k, b in self.nav.items():
            b.config(bg=COLORS["side_hi"] if k == key else COLORS["side"])
        self.current = key
        page = self.pages[key]
        page.tkraise()
        try:
            page.refresh()
        except Exception as e:  # noqa
            messagebox.showerror("Error", f"Could not load {key}: {e}")

    def refresh_all(self):
        self.comp_lbl.config(text=self.db.get("company_name"))
        self.show(self.current or "dashboard")

    def status(self, text):
        self.status_var.set(f"{datetime.now():%H:%M}  {text}")

    def on_close(self):
        try:
            self.db.backup()
            files = sorted(f for f in os.listdir(BACKUP_DIR) if f.startswith("office_"))
            for f in files[:-30]:                       # keep the last 30 backups
                os.remove(os.path.join(BACKUP_DIR, f))
        except Exception:  # noqa
            pass
        self.destroy()


if __name__ == "__main__":
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:  # noqa
            pass
    App().mainloop()
