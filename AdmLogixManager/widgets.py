"""Reusable Tkinter widgets: date picker, scrollable frame, table, form builder."""
import calendar
import os
import subprocess
import sys
import tkinter as tk
from datetime import date
from tkinter import messagebox, ttk

from db import fmt_date, parse_date

COLORS = {
    "bg": "#F4F6FA", "card": "#FFFFFF", "side": "#1F3A5F", "side_hi": "#2E5A8C",
    "accent": "#2E6DB4", "text": "#1E2633", "muted": "#6B7685", "green": "#1E8E5A",
    "red": "#C0392B", "amber": "#C77C02", "line": "#DDE3EC",
}
FONT = ("Segoe UI", 10)
FONT_B = ("Segoe UI", 10, "bold")


def open_file(path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)                      # noqa
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        messagebox.showinfo("Saved", f"File saved to:\n{path}")


def setup_style(root):
    st = ttk.Style(root)
    st.theme_use("clam")
    root.configure(bg=COLORS["bg"])
    root.option_add("*Font", FONT)
    st.configure(".", font=FONT, background=COLORS["bg"], foreground=COLORS["text"])
    st.configure("TFrame", background=COLORS["bg"])
    st.configure("Card.TFrame", background=COLORS["card"], relief="flat")
    st.configure("TLabel", background=COLORS["bg"])
    st.configure("Card.TLabel", background=COLORS["card"])
    st.configure("Muted.TLabel", background=COLORS["card"], foreground=COLORS["muted"])
    st.configure("H1.TLabel", font=("Segoe UI", 16, "bold"), background=COLORS["bg"])
    st.configure("H2.TLabel", font=("Segoe UI", 11, "bold"), background=COLORS["card"])
    st.configure("TLabelframe", background=COLORS["card"], bordercolor=COLORS["line"])
    st.configure("TLabelframe.Label", background=COLORS["card"], font=FONT_B, foreground=COLORS["accent"])
    st.configure("TCheckbutton", background=COLORS["card"])
    st.configure("TRadiobutton", background=COLORS["card"])
    st.configure("TButton", padding=(10, 5))
    st.configure("Accent.TButton", background=COLORS["accent"], foreground="white", font=FONT_B)
    st.map("Accent.TButton", background=[("active", "#24589A")])
    st.configure("Danger.TButton", foreground=COLORS["red"])
    st.configure("Treeview", rowheight=26, font=FONT, fieldbackground="white")
    st.configure("Treeview.Heading", font=FONT_B, background="#E8EEF6")
    st.map("Treeview", background=[("selected", "#CFE0F5")], foreground=[("selected", "black")])


class DateEntry(ttk.Frame):
    """Text box showing 10-Oct-2026 with a small calendar popup."""

    def __init__(self, master, value=None, width=12, on_change=None, **kw):
        super().__init__(master, **kw)
        self.on_change = on_change
        self.var = tk.StringVar(value=fmt_date(value or date.today()))
        self.entry = ttk.Entry(self, textvariable=self.var, width=width)
        self.entry.pack(side="left")
        ttk.Button(self, text="📅", width=3, command=self._popup).pack(side="left", padx=(2, 0))
        self.entry.bind("<Return>", lambda e: self._changed())
        self.entry.bind("<FocusOut>", lambda e: self._normalise())

    def _normalise(self):
        d = parse_date(self.var.get())
        if d:
            self.var.set(fmt_date(d))

    def _changed(self):
        self._normalise()
        if self.on_change:
            self.on_change()

    def get(self):
        """ISO date string or None."""
        d = parse_date(self.var.get())
        return d.isoformat() if d else None

    def get_date(self):
        return parse_date(self.var.get())

    def set(self, value):
        self.var.set(fmt_date(value) if value else "")

    def _popup(self):
        cur = self.get_date() or date.today()
        top = tk.Toplevel(self)
        top.title("Pick date")
        top.transient(self.winfo_toplevel())
        top.resizable(False, False)
        top.geometry(f"+{self.winfo_rootx()}+{self.winfo_rooty() + 30}")
        state = {"y": cur.year, "m": cur.month}
        head = ttk.Frame(top)
        head.pack(fill="x", padx=6, pady=4)
        lbl = ttk.Label(head, font=FONT_B, anchor="center")
        grid = ttk.Frame(top)
        grid.pack(padx=6, pady=(0, 6))

        def draw():
            for w in grid.winfo_children():
                w.destroy()
            lbl.config(text=f"{calendar.month_name[state['m']]} {state['y']}")
            for i, dname in enumerate(["Mo", "Tu", "We", "Th", "Fr", "Sa", "Su"]):
                ttk.Label(grid, text=dname, foreground=COLORS["muted"]).grid(row=0, column=i)
            for r, week in enumerate(calendar.monthcalendar(state["y"], state["m"]), 1):
                for c, day in enumerate(week):
                    if day:
                        b = tk.Button(grid, text=str(day), width=3, relief="flat", bg="white",
                                      command=lambda d=day: pick(d))
                        if date(state["y"], state["m"], day) == date.today():
                            b.config(bg="#CFE0F5")
                        if date(state["y"], state["m"], day) == cur:
                            b.config(bg=COLORS["accent"], fg="white")
                        b.grid(row=r, column=c, padx=1, pady=1)

        def move(step):
            m = state["m"] + step
            state["y"] += (m - 1) // 12
            state["m"] = (m - 1) % 12 + 1
            draw()

        def pick(day):
            self.set(date(state["y"], state["m"], day))
            top.destroy()
            if self.on_change:
                self.on_change()

        ttk.Button(head, text="◀", width=3, command=lambda: move(-1)).pack(side="left")
        lbl.pack(side="left", expand=True, fill="x")
        ttk.Button(head, text="▶", width=3, command=lambda: move(1)).pack(side="right")
        ttk.Button(top, text="Today", command=lambda: (state.update(y=date.today().year, m=date.today().month),
                                                        pick(date.today().day))).pack(pady=(0, 6))
        draw()
        top.grab_set()


class ScrollFrame(ttk.Frame):
    """Vertical scrolling container. Put children in .inner"""

    def __init__(self, master, bg=COLORS["card"], **kw):
        super().__init__(master, **kw)
        self.canvas = tk.Canvas(self, bg=bg, highlightthickness=0)
        sb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas, style="Card.TFrame")
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(win, width=e.width))
        self.canvas.configure(yscrollcommand=sb.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.canvas.bind("<Enter>", lambda e: self.canvas.bind_all("<MouseWheel>", self._wheel))
        self.canvas.bind("<Leave>", lambda e: self.canvas.unbind_all("<MouseWheel>"))

    def _wheel(self, e):
        self.canvas.yview_scroll(int(-e.delta / 120), "units")


class Table(ttk.Frame):
    """Treeview + scrollbars with simple column spec [(key, title, width, anchor)]."""

    def __init__(self, master, columns, height=12, **kw):
        super().__init__(master, **kw)
        self.tree = ttk.Treeview(self, show="headings", height=height, selectmode="browse")
        ys = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        xs = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.tree.tag_configure("odd", background="#F7F9FC")
        self.tree.tag_configure("income", foreground=COLORS["green"])
        self.tree.tag_configure("muted", foreground=COLORS["muted"])
        self.tree.tag_configure("warn", foreground=COLORS["red"])
        self.set_columns(columns)

    def set_columns(self, columns):
        self.columns = columns
        self.tree["columns"] = [c[0] for c in columns]
        for key, title, width, anchor in columns:
            self.tree.heading(key, text=title)
            self.tree.column(key, width=width, anchor=anchor, stretch=anchor == "w")

    def load(self, rows):
        """rows: list of (iid, values, tags)"""
        self.tree.delete(*self.tree.get_children())
        for i, (iid, values, tags) in enumerate(rows):
            tags = tuple(tags or ()) + (("odd",) if i % 2 else ())
            self.tree.insert("", "end", iid=str(iid), values=values, tags=tags)

    def selected(self):
        sel = self.tree.selection()
        return int(sel[0]) if sel else None

    def bind_select(self, fn):
        self.tree.bind("<<TreeviewSelect>>", lambda e: fn())

    def bind_double(self, fn):
        self.tree.bind("<Double-1>", lambda e: fn())


class Form:
    """Builds a two-column label/field grid. Field spec: (key, label, kind, options)
    kind: entry | combo | readonly | date | text"""

    def __init__(self, parent, spec, columns=1, label_width=14):
        self.widgets = {}
        self.kinds = {}
        self.parent = parent
        r = c = 0
        for key, label, kind, opts in spec:
            ttk.Label(parent, text=label, style="Card.TLabel", width=label_width).grid(
                row=r, column=c * 2, sticky="w", padx=(8, 4), pady=3)
            if kind == "date":
                w = DateEntry(parent, value=None, style="Card.TFrame")
                w.set("")
            elif kind in ("combo", "readonly"):
                w = ttk.Combobox(parent, values=opts or [], state="readonly" if kind == "readonly" else "normal")
            elif kind == "text":
                w = tk.Text(parent, height=3, width=28, font=FONT, relief="solid", bd=1)
            else:
                w = ttk.Entry(parent)
            w.grid(row=r, column=c * 2 + 1, sticky="ew", padx=(0, 10), pady=3)
            self.widgets[key] = w
            self.kinds[key] = kind
            c += 1
            if c >= columns:
                c = 0
                r += 1
        for i in range(columns):
            parent.columnconfigure(i * 2 + 1, weight=1)
        self.next_row = r + (1 if c else 0)

    def get(self, key):
        w, k = self.widgets[key], self.kinds[key]
        if k == "date":
            return w.get()
        if k == "text":
            return w.get("1.0", "end").strip()
        return w.get().strip()

    def set(self, key, value):
        w, k = self.widgets[key], self.kinds[key]
        value = "" if value is None else value
        if k == "date":
            w.set(value)
        elif k == "text":
            w.delete("1.0", "end")
            w.insert("1.0", value)
        elif k == "readonly":
            w.set(value)
        else:
            w.delete(0, "end")
            w.insert(0, str(value))

    def values(self):
        return {k: self.get(k) for k in self.widgets}

    def clear(self):
        for k in self.widgets:
            self.set(k, "")


def ask_string(parent, title, prompt, initial=""):
    from tkinter import simpledialog
    return simpledialog.askstring(title, prompt, initialvalue=initial, parent=parent)


class AutoCombo(ttk.Combobox):
    """Combobox that filters its list while typing (for descriptions)."""

    def __init__(self, master, values=(), **kw):
        super().__init__(master, **kw)
        self.all_values = list(values)
        self["values"] = self.all_values[:200]
        self.bind("<KeyRelease>", self._filter)

    def set_values(self, values):
        self.all_values = list(values)
        self["values"] = self.all_values[:200]

    def _filter(self, e):
        if e.keysym in ("Up", "Down", "Return", "Escape", "Tab"):
            return
        t = self.get().lower()
        self["values"] = [v for v in self.all_values if t in v.lower()][:200] if t else self.all_values[:200]
