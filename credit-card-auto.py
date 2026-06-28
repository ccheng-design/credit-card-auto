#!/usr/bin/env python3
"""
Capital One Statement Tools - GUI
Run with: python3 capital_one_gui.py
Requires: pip install pdfplumber customtkinter
"""

import csv
import io
import os
import re
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Add user site-packages to path so double-click works even when pip installed with --user
import site
for _sp in site.getusersitepackages() if isinstance(site.getusersitepackages(), list) else [site.getusersitepackages()]:
    if _sp not in sys.path:
        sys.path.insert(0, _sp)

# Also add common Windows user pip paths explicitly
import os as _os
_appdata = _os.environ.get("APPDATA", "")
_localappdata = _os.environ.get("LOCALAPPDATA", "")
for _base in [_appdata, _localappdata]:
    for _ver in ["Python313", "Python312", "Python311", "Python310", "Python39"]:
        _p = _os.path.join(_base, "Python", _ver, "site-packages")
        if _os.path.isdir(_p) and _p not in sys.path:
            sys.path.insert(0, _p)

try:
    import pdfplumber
except ImportError:
    import tkinter as _tk
    from tkinter import messagebox as _mb
    _r = _tk.Tk(); _r.withdraw()
    _mb.showerror("Missing dependency",
        "pdfplumber is not installed.\n\n"
        "Open a terminal and run:\n\n"
        "    pip install pdfplumber\n\n"
        "Then try again.")
    sys.exit(1)

# ── Category map ─────────────────────────────────────────────────────────────
CATEGORY_MAP = {
    "Car Rental":      "Travel Fund",
    "Dining":          "Food",
    "Entertainment":   "Savings",
    "Health Care":     "Savings",
    "Insurance":       "Medical Insurance",
    "Internet":        "Utilities",
    "Merchandise":     "Remainder",
    "Other Travel":    "Travel Fund",
    "Phone/Cable":     "Utilities",
    "Lodging":         "Travel Fund",
    "Gas/Automotive":  "Utilities",
    "Other Services":  "Travel Fund",
    "Payment/Credit":        "Savings",
    "Professional Services": "Utilities",
    "Fee/Interest Charge": "Savings",
}

CHASE_CATEGORY_MAP = {
    "Bills & Utilities": "Utilities",
    "Food & Drink":      "Food",
    "Gas":               "Utilities",
    "Groceries":         "Food",
    "Home":              "Savings",
    "Shopping":          "Remainder",
    "Travel":            "Travel Fund",
    "Health & Wellness": "Remainder",
}

COLORS = {
    "bg":       "#0f1117",
    "panel":    "#1a1d27",
    "border":   "#2a2d3a",
    "accent":   "#4f8ef7",
    "accent2":  "#7c5cbf",
    "success":  "#3ecf8e",
    "warning":  "#f59e0b",
    "danger":   "#ef4444",
    "text":     "#e8eaf0",
    "muted":    "#6b7280",
    "tag_isk":  "#1e3a5f",
    "tag_eur":  "#1e3d2f",
    "tag_dkk":  "#3d2a1e",
    "tag_czk":  "#2d1e3d",
    "tag_nzd":  "#1e3d3d",
    "tag_other":"#2d2d1e",
}

CURRENCY_COLORS = {
    "ISK": "#60a5fa",
    "EUR": "#34d399",
    "DKK": "#fb923c",
    "CZK": "#c084fc",
    "NZD": "#22d3ee",
    "CHW": "#f9a8d4",
    "CHF": "#f9a8d4",
    "TRY": "#fbbf24",
    "USD": "#a78bfa",
}

# ── PDF Parser ────────────────────────────────────────────────────────────────
# Capital One patterns
C1_DATE_RE  = re.compile(r"^((?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2})\s+((?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2})\s+(.+?)\s+(-?\s*\$[\d,]+\.\d{2})\s*$")
C1_AMOUNT_RE= re.compile(r"^\$?([\d,]+\.\d{2})\s*$")
C1_CUR_RE   = re.compile(r"^([A-Z]{3})\s*$")
C1_EXCH_RE  = re.compile(r"^([\d.]+)\s+Exchange Rate\s*$", re.IGNORECASE)

# Chase patterns
# Transaction line:  "MM/DD  MERCHANT NAME  amount"
CH_TX_RE    = re.compile(r"^(\d{2}/\d{2})\s+(.+?)\s+(-?[\d,]+\.\d{2})\s*$")
# Foreign sub-line:  "MM/DD  NEW ZEALAND DOLLAR" or "MM/DD  EURO" etc.
CH_CUR_RE   = re.compile(r"^\d{2}/\d{2}\s+([A-Z][A-Z\s]+)\s*$")
# Exchange line:     "amount X rate (EXCHG RATE)"
CH_EXCH_RE  = re.compile(r"^([\d,]+\.\d+)\s+X\s+([\d.]+)\s+\(EXCHG RATE\)\s*$", re.IGNORECASE)

CURRENCY_NAMES = {
    "NEW ZEALAND DOLLAR": "NZD",
    "EURO":               "EUR",
    "POUND STERLING":     "GBP",
    "CANADIAN DOLLAR":    "CAD",
    "AUSTRALIAN DOLLAR":  "AUD",
    "JAPANESE YEN":       "JPY",
    "SWISS FRANC":        "CHF",
    "SWEDISH KRONA":      "SEK",
    "NORWEGIAN KRONE":    "NOK",
    "DANISH KRONE":       "DKK",
    "MEXICAN PESO":       "MXN",
    "TURKISH LIRA":       "TRY",
    "ICELANDIC KRONA":    "ISK",
    "CZECH KORUNA":       "CZK",
    "HUNGARIAN FORINT":   "HUF",
    "POLISH ZLOTY":       "PLN",
    "SINGAPORE DOLLAR":   "SGD",
    "HONG KONG DOLLAR":   "HKD",
    "SOUTH KOREAN WON":   "KRW",
    "THAI BAHT":          "THB",
    "INDIAN RUPEE":       "INR",
    "CHINESE YUAN":       "CNY",
    "SOUTH AFRICAN RAND": "ZAR",
    "BRAZILIAN REAL":     "BRL",
}

def _parse_chase_pdf(path):
    """Parse a Chase PDF statement for foreign currency transactions."""
    rows = []
    with pdfplumber.open(path) as pdf:
        all_lines = []
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                all_lines.extend([ln.rstrip() for ln in text.splitlines()])
    return _parse_chase(all_lines)

def _detect_bank(lines):
    """Return 'chase' or 'capitalone' based on PDF content."""
    joined = " ".join(lines[:60]).upper()
    if "EXCHG RATE" in joined:
        return "chase"
    return "capitalone"

def parse_pdf(path):
    rows = []
    with pdfplumber.open(path) as pdf:
        all_lines = []
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                all_lines.extend([ln.rstrip() for ln in text.splitlines()])

        bank = _detect_bank(all_lines)

        if bank == "chase":
            rows = _parse_chase(all_lines)
        else:
            rows = _parse_capitalone(all_lines)
    return rows

def _parse_capitalone(lines):
    rows = []
    i = 0
    while i < len(lines):
        m = C1_DATE_RE.match(lines[i])
        if m:
            trans_date, post_date, desc, usd_raw = m.groups()
            usd_str  = usd_raw.replace(" ", "").lstrip("-$").replace(",", "")
            is_credit = "-" in usd_raw
            fa = cur = rate = None
            j = i + 1
            while j < len(lines) and j < i + 6:
                ln = lines[j].strip()
                if not ln or ("Exchange Rate" not in ln and not C1_AMOUNT_RE.match(ln) and not C1_CUR_RE.match(ln)):
                    if fa is None:
                        break
                if fa is None:
                    ma = C1_AMOUNT_RE.match(ln)
                    if ma:
                        fa = ma.group(1).replace(",", "")
                elif cur is None:
                    mc = C1_CUR_RE.match(ln)
                    if mc:
                        cur = mc.group(1)
                elif rate is None:
                    me = C1_EXCH_RE.match(ln)
                    if me:
                        rate = me.group(1)
                        i = j
                        break
                j += 1
            if cur and cur != "USD":
                rows.append({
                    "Transaction Date": trans_date,
                    "Posted Date":      post_date,
                    "Description":      desc.strip(),
                    "USD Amount":       ("-" if is_credit else "") + usd_str,
                    "Foreign Amount":   fa or "",
                    "Currency":         cur,
                    "Exchange Rate":    rate or "",
                    "Type":             "Credit" if is_credit else "Charge",
                    "Bank":             "Capital One",
                })
        i += 1
    return rows

# Lines to skip when scanning across Chase page boundaries
_CHASE_SKIP_RE = re.compile(
    r"^(x\s*$|0000001\s+FIS|Manage\s+your\s+account|MMaann"
    r"|www\.chase\.com|ACCOUNT ACTIVITY|Date of"
    r"|Transaction\s+Merchant|CLARK CHENG|Page \d+ of \d+"
    r"|Customer Service|Download|Chase Mobile|\d{10,})",
    re.IGNORECASE
)

# Section headers that match the tx regex but aren't real transactions
_CHASE_SKIP_DESC_RE = re.compile(
    r"^(PAYMENTS AND OTHER CREDITS|PURCHASE|YOUR ACCOUNT MESSAGES"
    r"|AUTOMATIC PAYMENT - THANK YOU|ACCOUNT ACTIVITY)",
    re.IGNORECASE
)

def _parse_chase(lines):
    rows = []
    i = 0
    while i < len(lines):
        m = CH_TX_RE.match(lines[i].strip())
        if m:
            tx_date, desc, amt_raw = m.groups()
            is_credit = amt_raw.startswith("-")
            usd_str   = amt_raw.lstrip("-").replace(",", "")

            # Skip section header lines that happen to match the tx regex
            if _CHASE_SKIP_DESC_RE.match(desc.strip()):
                i += 1
                continue

            # Scan ahead up to 20 lines, skipping page header/footer noise,
            # to find the optional currency name + exchange rate block.
            j = i + 1
            cur_name = cur = fa = rate = None
            while j < len(lines) and j < i + 20:
                ln = lines[j].strip()
                if not ln or _CHASE_SKIP_RE.match(ln):
                    j += 1
                    continue
                if cur_name is None:
                    mc = CH_CUR_RE.match(ln)
                    if mc:
                        cur_name = mc.group(1).strip()
                        cur = CURRENCY_NAMES.get(cur_name, cur_name[:3].upper())
                        j += 1
                        continue
                    else:
                        # Next real content is not a currency line → USD transaction
                        break
                else:
                    me = CH_EXCH_RE.match(ln)
                    if me:
                        fa   = me.group(1).replace(",", "")
                        rate = me.group(2)
                        i = j
                    break
                j += 1

            if cur and cur != "USD":
                # Foreign currency transaction
                rows.append({
                    "Transaction Date": tx_date,
                    "Posted Date":      "-",
                    "Description":      desc.strip(),
                    "USD Amount":       ("-" if is_credit else "") + usd_str,
                    "Foreign Amount":   fa or "-",
                    "Currency":         cur,
                    "Exchange Rate":    rate or "-",
                    "Type":             "Credit" if is_credit else "Charge",
                    "Bank":             "Chase",
                })
            else:
                # USD transaction — no foreign block, fill with dashes, color purple
                rows.append({
                    "Transaction Date": tx_date,
                    "Posted Date":      "-",
                    "Description":      desc.strip(),
                    "USD Amount":       ("-" if is_credit else "") + usd_str,
                    "Foreign Amount":   "-",
                    "Currency":         "USD",
                    "Exchange Rate":    "-",
                    "Type":             "Credit" if is_credit else "Charge",
                    "Bank":             "Chase",
                })
        i += 1
    return rows

# ── Remap ─────────────────────────────────────────────────────────────────────
def remap_csv(path):
    rows = []
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        fieldnames = list(reader.fieldnames or [])
        cat_col = next((c for c in fieldnames if c.strip().lower() == "category"), None)
        for row in reader:
            orig = row.get(cat_col, "").strip() if cat_col else ""
            row["My Category"] = CATEGORY_MAP.get(orig, orig or "Uncategorized")
            rows.append(row)
    out_fields = fieldnames + ["My Category"]
    return rows, out_fields

def remap_chase_csv(path, live_map=None):
    mapping = live_map or CHASE_CATEGORY_MAP
    rows = []
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        if reader.fieldnames is None:
            return [], []
        fieldnames = list(reader.fieldnames)
        cat_col = next((c for c in fieldnames if c.strip().lower() == "category"), None)
        for row in reader:
            orig = row.get(cat_col, "").strip() if cat_col else ""
            row["My Category"] = mapping.get(orig, orig or "Uncategorized")
            rows.append(row)
    out_fields = fieldnames + ["My Category"]
    return rows, out_fields


def _sort_rows_by_master(pdf_rows, master_path, date_col, amount_col):
    """Re-order pdf_rows to match the row order in master_path CSV.
    Matching by (MM/DD date, abs amount). Unmatched rows appended at end."""
    from difflib import SequenceMatcher
    from collections import defaultdict

    def norm_date(d):
        d = str(d).strip()
        if '/' in d:
            parts = d.split('/')
            return f"{parts[0].zfill(2)}/{parts[1].zfill(2)}"
        months = {"Jan":"01","Feb":"02","Mar":"03","Apr":"04","May":"05","Jun":"06",
                  "Jul":"07","Aug":"08","Sep":"09","Oct":"10","Nov":"11","Dec":"12"}
        for name, num in months.items():
            if d.startswith(name):
                return f"{num}/{d.split()[-1].zfill(2)}"
        return d[:5]

    def norm_amount(a):
        try:
            return round(abs(float(str(a).replace(',', '').replace('$', ''))), 2)
        except Exception:
            return 0.0

    def norm_desc(d):
        return str(d).strip().upper()[:12]

    try:
        import csv as _csv
        with open(master_path, newline='', encoding='utf-8-sig') as f:
            master_rows = list(_csv.DictReader(f))
    except Exception:
        return pdf_rows

    # Detect amount column if not found exactly
    if master_rows and amount_col not in master_rows[0]:
        for candidate in ['Amount', 'Debit', 'amount', 'debit']:
            if candidate in master_rows[0]:
                amount_col = candidate
                break

    lookup = defaultdict(list)
    for idx, row in enumerate(pdf_rows):
        key = (norm_date(row.get("Transaction Date", "")),
               norm_amount(row.get("USD Amount", "0")))
        lookup[key].append(idx)

    used = set()
    sorted_rows = []

    for mrow in master_rows:
        md = norm_date(mrow.get(date_col, ""))
        ma = norm_amount(mrow.get(amount_col, "0"))
        key = (md, ma)
        candidates = [i for i in lookup.get(key, []) if i not in used]
        if not candidates:
            continue
        if len(candidates) == 1:
            chosen = candidates[0]
        else:
            mdesc = norm_desc(mrow.get("Description", ""))
            best_idx, best_score = candidates[0], 0
            for ci in candidates:
                score = SequenceMatcher(
                    None, mdesc,
                    norm_desc(pdf_rows[ci].get("Description", ""))
                ).ratio()
                if score > best_score:
                    best_score = score
                    best_idx = ci
            chosen = best_idx
        used.add(chosen)
        sorted_rows.append(pdf_rows[chosen])

    # Append any unmatched PDF rows at end
    for idx in range(len(pdf_rows)):
        if idx not in used:
            sorted_rows.append(pdf_rows[idx])

    return sorted_rows

# ── Tooltip helper ────────────────────────────────────────────────────────────
class Tooltip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip = None
        widget.bind("<Enter>", self.show)
        widget.bind("<Leave>", self.hide)

    def show(self, _=None):
        x, y, _, _ = self.widget.bbox("insert") if hasattr(self.widget,"bbox") else (0,0,0,0)
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        self.tip = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tk.Label(tw, text=self.text, bg="#1a1d27", fg="#e8eaf0",
                 relief="flat", bd=1, padx=8, pady=4,
                 font=("SF Pro Text", 11)).pack()

    def hide(self, _=None):
        if self.tip:
            self.tip.destroy()
            self.tip = None

# ── Main App ──────────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Capital One Statement Tools")
        self.geometry("1100x720")
        self.minsize(900, 600)
        self.configure(bg=COLORS["bg"])
        self._pdf_rows = []
        self._csv_rows = []
        self._csv_fields = []
        self._build_ui()

    # ── Layout ────────────────────────────────────────────────────────────────
    def _build_ui(self):
        # Top bar
        bar = tk.Frame(self, bg=COLORS["panel"], height=52)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        tk.Label(bar, text="Capital One Tools", bg=COLORS["panel"],
                 fg=COLORS["text"], font=("SF Pro Display", 15, "bold")).pack(side="left", padx=20)

        # Tab row
        self._tab_var = tk.StringVar(value="pdf")
        tab_frame = tk.Frame(self, bg=COLORS["bg"])
        tab_frame.pack(fill="x", padx=20, pady=(12, 0))
        for label, val in [("Capital One Statement → Foreign CSV", "pdf"), ("Chase → Foreign CSV", "chase"), ("Capital One Remap Categories", "remap"), ("Chase Remap Categories", "chase_remap")]:
            btn = tk.Button(tab_frame, text=label, bg=COLORS["bg"], fg=COLORS["muted"],
                            font=("SF Pro Text", 12), relief="flat", bd=0,
                            activebackground=COLORS["bg"], activeforeground=COLORS["text"],
                            cursor="hand2",
                            command=lambda v=val: self._switch_tab(v))
            btn.pack(side="left", padx=(0, 4))
            setattr(self, f"_tab_btn_{val}", btn)

        # Tab underline
        self._tab_underline = tk.Frame(self, bg=COLORS["accent"], height=2)
        self._tab_underline.place(x=20, y=60, width=160)

        # Content area
        self._content = tk.Frame(self, bg=COLORS["bg"])
        self._content.pack(fill="both", expand=True, padx=20, pady=12)

        self._build_pdf_tab()
        self._build_chase_tab()
        self._build_remap_tab()
        self._build_chase_remap_tab()
        self._switch_tab("pdf")

    def _switch_tab(self, val):
        self._tab_var.set(val)
        all_tabs = ["pdf", "chase", "remap", "chase_remap"]
        for v in all_tabs:
            btn = getattr(self, f"_tab_btn_{v}", None)
            if btn:
                btn.config(fg=COLORS["text"] if v == val else COLORS["muted"],
                           font=("SF Pro Text", 12, "bold" if v == val else "normal"))
        # Measure actual button position so underline always aligns
        self.update_idletasks()
        active_btn = getattr(self, f"_tab_btn_{val}", None)
        if active_btn:
            bx = active_btn.winfo_rootx() - self.winfo_rootx()
            bw = active_btn.winfo_width()
            by = active_btn.winfo_rooty() - self.winfo_rooty() + active_btn.winfo_height()
            self._tab_underline.place(x=bx, y=by, width=bw)
        for v in all_tabs:
            frame = getattr(self, f"_{v}_frame", None)
            if frame:
                frame.pack_forget()
        target = getattr(self, f"_{val}_frame", None)
        if target:
            target.pack(fill="both", expand=True)

    # ── PDF Tab ───────────────────────────────────────────────────────────────
    def _build_pdf_tab(self):
        f = self._pdf_frame = tk.Frame(self._content, bg=COLORS["bg"])

        # Upload row
        upload_row = tk.Frame(f, bg=COLORS["bg"])
        upload_row.pack(fill="x", pady=(0, 10))

        self._pdf_path_var = tk.StringVar(value="No file selected")
        tk.Button(upload_row, text="  Upload Capital One PDF Statement  ",
                  bg=COLORS["accent"], fg="white", font=("SF Pro Text", 12, "bold"),
                  relief="flat", bd=0, padx=14, pady=8, cursor="hand2",
                  activebackground="#3a7de8", activeforeground="white",
                  command=self._pick_pdf).pack(side="left")

        self._pdf_label = tk.Label(upload_row, textvariable=self._pdf_path_var,
                                   bg=COLORS["bg"], fg=COLORS["muted"],
                                   font=("SF Pro Text", 11))
        self._pdf_label.pack(side="left", padx=12)

        # Stats bar
        self._pdf_stats = tk.Frame(f, bg=COLORS["panel"], pady=8)
        self._pdf_stats.pack(fill="x", pady=(0, 10))
        self._stat_labels = {}
        for key in ["Total", "ISK", "EUR", "DKK", "CZK", "NZD", "TRY", "Other"]:
            cell = tk.Frame(self._pdf_stats, bg=COLORS["panel"])
            cell.pack(side="left", padx=16)
            tk.Label(cell, text=key, bg=COLORS["panel"], fg=COLORS["muted"],
                     font=("SF Pro Text", 10)).pack()
            lbl = tk.Label(cell, text="—", bg=COLORS["panel"], fg=COLORS["text"],
                           font=("SF Pro Text", 13, "bold"))
            lbl.pack()
            self._stat_labels[key] = lbl

        # Filter bar
        filter_row = tk.Frame(f, bg=COLORS["bg"])
        filter_row.pack(fill="x", pady=(0, 6))
        tk.Label(filter_row, text="Filter:", bg=COLORS["bg"], fg=COLORS["muted"],
                 font=("SF Pro Text", 11)).pack(side="left")
        self._pdf_filter = tk.StringVar()
        self._pdf_filter.trace_add("write", lambda *_: self._apply_pdf_filter())
        entry = tk.Entry(filter_row, textvariable=self._pdf_filter,
                         bg=COLORS["panel"], fg=COLORS["text"], insertbackground=COLORS["text"],
                         font=("SF Pro Text", 11), relief="flat", bd=0, width=28)
        entry.pack(side="left", padx=(6,16), ipady=5, ipadx=6)
        Tooltip(entry, "Filter by description, currency, or date")

        self._cur_filter_var = tk.StringVar(value="All")
        cur_menu = tk.OptionMenu(filter_row, self._cur_filter_var,
                                  "All", "ISK", "EUR", "DKK", "CZK", "NZD", "TRY", "CHW",
                                  command=lambda _: self._apply_pdf_filter())
        cur_menu.config(bg=COLORS["panel"], fg=COLORS["text"], font=("SF Pro Text", 11),
                        relief="flat", bd=0, activebackground=COLORS["border"],
                        highlightthickness=0)
        cur_menu["menu"].config(bg=COLORS["panel"], fg=COLORS["text"])
        cur_menu.pack(side="left")

        # Table
        cols = ("Trans Date","Post Date","Description","USD Amount","Foreign Amount","Currency","Exchange Rate","Type","Bank")
        col_widths = (80, 80, 260, 90, 110, 70, 130, 70, 90)

        frame_tree = tk.Frame(f, bg=COLORS["border"], bd=1)
        frame_tree.pack(fill="both", expand=True, pady=(0, 10))

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Custom.Treeview",
                         background=COLORS["panel"],
                         foreground=COLORS["text"],
                         fieldbackground=COLORS["panel"],
                         rowheight=28,
                         font=("SF Pro Text", 11),
                         borderwidth=0)
        style.configure("Custom.Treeview.Heading",
                         background=COLORS["border"],
                         foreground=COLORS["muted"],
                         font=("SF Pro Text", 10, "bold"),
                         relief="flat", borderwidth=0)
        style.map("Custom.Treeview",
                  background=[("selected", COLORS["accent2"])],
                  foreground=[("selected", "white")])

        self._pdf_tree = ttk.Treeview(frame_tree, columns=cols, show="headings",
                                       style="Custom.Treeview")
        for col, w in zip(cols, col_widths):
            self._pdf_tree.heading(col, text=col, anchor="w",
                                   command=lambda c=col: self._sort_pdf(c))
            self._pdf_tree.column(col, width=w, minwidth=40, anchor="w")

        vsb = ttk.Scrollbar(frame_tree, orient="vertical", command=self._pdf_tree.yview)
        hsb = ttk.Scrollbar(frame_tree, orient="horizontal", command=self._pdf_tree.xview)
        self._pdf_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self._pdf_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame_tree.grid_rowconfigure(0, weight=1)
        frame_tree.grid_columnconfigure(0, weight=1)

        # Tag colors per currency
        for cur, color in CURRENCY_COLORS.items():
            self._pdf_tree.tag_configure(cur, foreground=color)
        self._pdf_tree.tag_configure("Credit", foreground=COLORS["success"])

        # Bottom bar
        bot = tk.Frame(f, bg=COLORS["bg"])
        bot.pack(fill="x")
        self._pdf_count_lbl = tk.Label(bot, text="0 transactions", bg=COLORS["bg"],
                                        fg=COLORS["muted"], font=("SF Pro Text", 11))
        self._pdf_count_lbl.pack(side="left")
        tk.Button(bot, text="  Export CSV  ",
                  bg=COLORS["success"], fg="#0a1a10", font=("SF Pro Text", 12, "bold"),
                  relief="flat", bd=0, padx=14, pady=7, cursor="hand2",
                  activebackground="#2eb87a",
                  command=self._export_pdf_csv).pack(side="right")
        tk.Button(bot, text="  Reset  ",
                  bg=COLORS["danger"], fg="white", font=("SF Pro Text", 12, "bold"),
                  relief="flat", bd=0, padx=14, pady=7, cursor="hand2",
                  activebackground="#cc2222",
                  command=self._reset_pdf).pack(side="right", padx=(0, 8))
        tk.Button(bot, text="  Sort by Master CSV  ",
                  bg=COLORS["warning"], fg="#1a1000", font=("SF Pro Text", 12, "bold"),
                  relief="flat", bd=0, padx=14, pady=7, cursor="hand2",
                  activebackground="#d48a00",
                  command=self._sort_pdf_by_master).pack(side="right", padx=(0, 8))

        self._sort_pdf_col = None
        self._sort_pdf_rev = False

    def _sort_pdf_by_master(self):
        if not self._pdf_rows:
            messagebox.showwarning("No data", "Parse a Capital One PDF first.")
            return
        path = filedialog.askopenfilename(
            title="Select Master Capital One CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if not path:
            return
        import csv as _csv
        try:
            with open(path, newline='', encoding='utf-8-sig') as f:
                sample = list(_csv.DictReader(f))
            amount_col = next((c for c in ["Debit","Amount","amount","debit"]
                               if sample and c in sample[0]), "Amount")
        except Exception:
            amount_col = "Amount"
        sorted_rows = _sort_rows_by_master(
            self._pdf_rows, path,
            date_col="Transaction Date", amount_col=amount_col)
        self._pdf_rows = sorted_rows
        self._apply_pdf_filter()
        self._pdf_count_lbl.config(
            text=f"{len(sorted_rows)} rows (sorted to master order)",
            fg=COLORS["success"])

    def _reset_pdf(self):
        self._pdf_rows = []
        self._pdf_path_var.set("No file selected")
        self._pdf_label.config(fg=COLORS["muted"])
        self._pdf_filter.set("")
        self._cur_filter_var.set("All")
        self._pdf_tree.delete(*self._pdf_tree.get_children())
        self._pdf_count_lbl.config(text="0 transactions", fg=COLORS["muted"])
        for key in self._stat_labels:
            self._stat_labels[key].config(text="—")

    def _pick_pdf(self):
        path = filedialog.askopenfilename(
            title="Select Capital One PDF Statement",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if not path:
            return
        self._pdf_path_var.set(os.path.basename(path))
        self._pdf_label.config(fg=COLORS["text"])
        self._pdf_count_lbl.config(text="Parsing…", fg=COLORS["warning"])
        self.update()

        def do_parse():
            try:
                rows = parse_pdf(path)
                self.after(0, lambda: self._load_pdf_rows(rows))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Parse Error", str(e)))
                self.after(0, lambda: self._pdf_count_lbl.config(text="Error", fg=COLORS["danger"]))

        threading.Thread(target=do_parse, daemon=True).start()

    def _load_pdf_rows(self, rows):
        self._pdf_rows = rows
        self._apply_pdf_filter()
        self._update_pdf_stats(rows)

    def _update_pdf_stats(self, rows):
        from collections import defaultdict
        CURRENCY_SYMBOLS = {
            "ISK": "kr", "EUR": "€", "DKK": "kr", "CZK": "Kč",
            "NZD": "NZ$", "TRY": "₺", "CHW": "Fr", "CHF": "Fr",
        }
        totals_usd = defaultdict(float)
        totals_foreign = defaultdict(float)
        for r in rows:
            if r["Type"] == "Credit":
                continue
            cur = r["Currency"]
            try:
                usd_val = float(r["USD Amount"].lstrip("-") or 0)
            except Exception:
                usd_val = 0.0
            totals_usd[cur] += usd_val
            if cur not in ("-", "USD") and r.get("Foreign Amount", "-") not in ("-", ""):
                try:
                    totals_foreign[cur] += float(r["Foreign Amount"])
                except Exception:
                    pass
        total_all = sum(totals_usd.values())
        self._stat_labels["Total"].config(text=f"${total_all:,.2f}")
        known = {"ISK","EUR","DKK","CZK","NZD","TRY"}
        for k in known:
            sym = CURRENCY_SYMBOLS.get(k, k)
            v = totals_foreign.get(k, 0)
            self._stat_labels[k].config(text=f"{sym}{v:,.2f}" if v else "—")
        other_usd = sum(v for k, v in totals_usd.items() if k not in known)
        self._stat_labels["Other"].config(text=f"${other_usd:,.2f}" if other_usd else "—")

    def _apply_pdf_filter(self):
        query = self._pdf_filter.get().lower()
        cur_f  = self._cur_filter_var.get()
        filtered = [r for r in self._pdf_rows
                    if (cur_f == "All" or r["Currency"] == cur_f)
                    and (not query or query in r["Description"].lower()
                         or query in r["Transaction Date"].lower()
                         or query in r["Posted Date"].lower()
                         or query in r["Currency"].lower())]
        self._pdf_count_lbl.config(
            text=f"{len(filtered)} transaction{'s' if len(filtered)!=1 else ''}",
            fg=COLORS["muted"])
        self._pdf_tree.delete(*self._pdf_tree.get_children())
        for r in filtered:
            tags = [r["Currency"]]
            if r["Type"] == "Credit":
                tags.append("Credit")
            self._pdf_tree.insert("", "end", values=(
                r["Transaction Date"], r["Posted Date"], r["Description"],
                r["USD Amount"], r["Foreign Amount"], r["Currency"],
                r["Exchange Rate"], r["Type"]
            ), tags=tags)

    def _sort_pdf(self, col):
        if self._sort_pdf_col == col:
            self._sort_pdf_rev = not self._sort_pdf_rev
        else:
            self._sort_pdf_col = col
            self._sort_pdf_rev = False
        self._pdf_rows.sort(key=lambda r: r.get(col,""), reverse=self._sort_pdf_rev)
        self._apply_pdf_filter()

    def _export_pdf_csv(self):
        if not self._pdf_rows:
            messagebox.showwarning("Nothing to export", "Parse a PDF statement first.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files","*.csv")],
            initialfile="foreign_transactions.csv")
        if not path:
            return
        fields = ["Transaction Date","Posted Date","Description",
                  "USD Amount","Foreign Amount","Currency","Exchange Rate","Type","Bank"]
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
            w.writeheader()
            w.writerows(self._pdf_rows)
        messagebox.showinfo("Exported", f"Saved {len(self._pdf_rows)} rows to:\n{path}")

    # ── Chase Tab ─────────────────────────────────────────────────────────────
    def _build_chase_tab(self):
        f = self._chase_frame = tk.Frame(self._content, bg=COLORS["bg"])

        # Upload row
        upload_row = tk.Frame(f, bg=COLORS["bg"])
        upload_row.pack(fill="x", pady=(0, 10))
        tk.Button(upload_row, text="  Upload Chase PDF Statement  ",
                  bg="#1a6fa8", fg="white", font=("SF Pro Text", 12, "bold"),
                  relief="flat", bd=0, padx=14, pady=8, cursor="hand2",
                  activebackground="#155d8e",
                  command=self._pick_chase).pack(side="left")
        self._chase_path_var = tk.StringVar(value="No file selected")
        self._chase_label = tk.Label(upload_row, textvariable=self._chase_path_var,
                                     bg=COLORS["bg"], fg=COLORS["muted"],
                                     font=("SF Pro Text", 11))
        self._chase_label.pack(side="left", padx=12)

        # Stats bar
        self._chase_stats = tk.Frame(f, bg=COLORS["panel"], pady=8)
        self._chase_stats.pack(fill="x", pady=(0, 10))
        self._chase_stat_labels = {}
        for key in ["Total", "NZD", "EUR", "GBP", "CAD", "AUD", "JPY", "Other"]:
            cell = tk.Frame(self._chase_stats, bg=COLORS["panel"])
            cell.pack(side="left", padx=16)
            tk.Label(cell, text=key, bg=COLORS["panel"], fg=COLORS["muted"],
                     font=("SF Pro Text", 10)).pack()
            lbl = tk.Label(cell, text="—", bg=COLORS["panel"], fg=COLORS["text"],
                           font=("SF Pro Text", 13, "bold"))
            lbl.pack()
            self._chase_stat_labels[key] = lbl

        # Filter bar
        filter_row = tk.Frame(f, bg=COLORS["bg"])
        filter_row.pack(fill="x", pady=(0, 6))
        tk.Label(filter_row, text="Filter:", bg=COLORS["bg"], fg=COLORS["muted"],
                 font=("SF Pro Text", 11)).pack(side="left")
        self._chase_filter = tk.StringVar()
        self._chase_filter.trace_add("write", lambda *_: self._apply_chase_filter())
        tk.Entry(filter_row, textvariable=self._chase_filter,
                 bg=COLORS["panel"], fg=COLORS["text"], insertbackground=COLORS["text"],
                 font=("SF Pro Text", 11), relief="flat", bd=0, width=28
                 ).pack(side="left", padx=(6,16), ipady=5, ipadx=6)

        self._chase_cur_filter = tk.StringVar(value="All")
        cur_menu = tk.OptionMenu(filter_row, self._chase_cur_filter,
                                 "All", "NZD", "EUR", "GBP", "CAD", "AUD", "JPY",
                                 "SGD", "HKD", "THB", "MXN", "TRY",
                                 command=lambda _: self._apply_chase_filter())
        cur_menu.config(bg=COLORS["panel"], fg=COLORS["text"], font=("SF Pro Text", 11),
                        relief="flat", bd=0, activebackground=COLORS["border"],
                        highlightthickness=0)
        cur_menu["menu"].config(bg=COLORS["panel"], fg=COLORS["text"])
        cur_menu.pack(side="left")

        # Table
        cols = ("Trans Date", "Description", "USD Amount", "Foreign Amount", "Currency", "Exchange Rate", "Type")
        col_widths = (80, 320, 100, 110, 70, 130, 70)

        frame_tree = tk.Frame(f, bg=COLORS["border"], bd=1)
        frame_tree.pack(fill="both", expand=True, pady=(0, 10))

        self._chase_tree = ttk.Treeview(frame_tree, columns=cols, show="headings",
                                         style="Custom.Treeview")
        for col, w in zip(cols, col_widths):
            self._chase_tree.heading(col, text=col, anchor="w",
                                     command=lambda c=col: self._sort_chase(c))
            self._chase_tree.column(col, width=w, minwidth=40, anchor="w")

        vsb = ttk.Scrollbar(frame_tree, orient="vertical", command=self._chase_tree.yview)
        hsb = ttk.Scrollbar(frame_tree, orient="horizontal", command=self._chase_tree.xview)
        self._chase_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self._chase_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame_tree.grid_rowconfigure(0, weight=1)
        frame_tree.grid_columnconfigure(0, weight=1)

        for cur, color in CURRENCY_COLORS.items():
            self._chase_tree.tag_configure(cur, foreground=color)
        self._chase_tree.tag_configure("USD", foreground="#a78bfa")
        self._chase_tree.tag_configure("Credit", foreground=COLORS["success"])

        # Bottom bar
        bot = tk.Frame(f, bg=COLORS["bg"])
        bot.pack(fill="x")
        self._chase_count_lbl = tk.Label(bot, text="0 transactions", bg=COLORS["bg"],
                                          fg=COLORS["muted"], font=("SF Pro Text", 11))
        self._chase_count_lbl.pack(side="left")
        tk.Button(bot, text="  Export CSV  ",
                  bg=COLORS["success"], fg="#0a1a10", font=("SF Pro Text", 12, "bold"),
                  relief="flat", bd=0, padx=14, pady=7, cursor="hand2",
                  activebackground="#2eb87a",
                  command=self._export_chase_csv).pack(side="right")
        tk.Button(bot, text="  Reset  ",
                  bg=COLORS["danger"], fg="white", font=("SF Pro Text", 12, "bold"),
                  relief="flat", bd=0, padx=14, pady=7, cursor="hand2",
                  activebackground="#cc2222",
                  command=self._reset_chase).pack(side="right", padx=(0, 8))
        tk.Button(bot, text="  Sort by Master CSV  ",
                  bg=COLORS["warning"], fg="#1a1000", font=("SF Pro Text", 12, "bold"),
                  relief="flat", bd=0, padx=14, pady=7, cursor="hand2",
                  activebackground="#d48a00",
                  command=self._sort_chase_by_master).pack(side="right", padx=(0, 8))

        self._chase_rows = []
        self._sort_chase_col = None
        self._sort_chase_rev = False

    def _sort_chase_by_master(self):
        if not self._chase_rows:
            messagebox.showwarning("No data", "Parse a Chase PDF first.")
            return
        path = filedialog.askopenfilename(
            title="Select Master Chase CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if not path:
            return
        sorted_rows = _sort_rows_by_master(
            self._chase_rows, path,
            date_col="Transaction Date", amount_col="Amount")
        self._chase_rows = sorted_rows
        self._apply_chase_filter()
        self._chase_count_lbl.config(
            text=f"{len(sorted_rows)} rows (sorted to master order)",
            fg=COLORS["success"])

    def _reset_chase(self):
        self._chase_rows = []
        self._chase_path_var.set("No file selected")
        self._chase_label.config(fg=COLORS["muted"])
        self._chase_filter.set("")
        self._chase_cur_filter.set("All")
        self._chase_tree.delete(*self._chase_tree.get_children())
        self._chase_count_lbl.config(text="0 transactions", fg=COLORS["muted"])
        for key in self._chase_stat_labels:
            self._chase_stat_labels[key].config(text="—")

    def _pick_chase(self):
        path = filedialog.askopenfilename(
            title="Select Chase PDF Statement",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if not path:
            return
        self._chase_path_var.set(os.path.basename(path))
        self._chase_label.config(fg=COLORS["text"])
        self._chase_count_lbl.config(text="Parsing…", fg=COLORS["warning"])
        self.update()

        def do_parse():
            try:
                rows = _parse_chase_pdf(path)
                self.after(0, lambda: self._load_chase_rows(rows))
            except Exception as e:
                import traceback
                self.after(0, lambda: messagebox.showerror("Parse Error", traceback.format_exc()))
                self.after(0, lambda: self._chase_count_lbl.config(text="Error", fg=COLORS["danger"]))

        threading.Thread(target=do_parse, daemon=True).start()

    def _load_chase_rows(self, rows):
        self._chase_rows = rows
        self._apply_chase_filter()
        self._update_chase_stats(rows)

    def _update_chase_stats(self, rows):
        from collections import defaultdict
        # Use Foreign Amount for foreign currencies, USD Amount for USD rows
        CURRENCY_SYMBOLS = {
            "NZD": "NZ$", "EUR": "€", "GBP": "£", "CAD": "CA$",
            "AUD": "A$",  "JPY": "¥", "USD": "$",
        }
        totals_usd = defaultdict(float)   # USD equiv per currency (for Total)
        totals_foreign = defaultdict(float) # native amount per currency
        for r in rows:
            if r["Type"] == "Credit":
                continue
            cur = r["Currency"]
            try:
                usd_val = float(r["USD Amount"].lstrip("-") or 0)
            except Exception:
                usd_val = 0.0
            totals_usd[cur] += usd_val
            if cur != "USD" and r.get("Foreign Amount", "-") not in ("-", ""):
                try:
                    totals_foreign[cur] += float(r["Foreign Amount"])
                except Exception:
                    pass
        total_all = sum(totals_usd.values())
        self._chase_stat_labels["Total"].config(text=f"${total_all:,.2f}")
        known = {"NZD", "EUR", "GBP", "CAD", "AUD", "JPY"}
        for k in known:
            sym = CURRENCY_SYMBOLS.get(k, k)
            v = totals_foreign.get(k, 0)
            self._chase_stat_labels[k].config(text=f"{sym}{v:,.2f}" if v else "—")
        other_usd = sum(v for k, v in totals_usd.items() if k not in known and k != "USD")
        self._chase_stat_labels["Other"].config(text=f"${other_usd:,.2f}" if other_usd else "—")

    def _apply_chase_filter(self):
        query  = self._chase_filter.get().lower()
        cur_f  = self._chase_cur_filter.get()
        filtered = [r for r in self._chase_rows
                    if (cur_f == "All" or r["Currency"] == cur_f)
                    and (not query or query in r["Description"].lower()
                         or query in r["Transaction Date"].lower()
                         or query in r["Currency"].lower())]
        self._chase_count_lbl.config(
            text=f"{len(filtered)} transaction{'s' if len(filtered)!=1 else ''}",
            fg=COLORS["muted"])
        self._chase_tree.delete(*self._chase_tree.get_children())
        for r in filtered:
            tags = [r["Currency"]]
            if r["Type"] == "Credit":
                tags.append("Credit")
            self._chase_tree.insert("", "end", values=(
                r["Transaction Date"], r["Description"],
                r["USD Amount"], r["Foreign Amount"],
                r["Currency"], r["Exchange Rate"], r["Type"]
            ), tags=tags)

    def _sort_chase(self, col):
        if self._sort_chase_col == col:
            self._sort_chase_rev = not self._sort_chase_rev
        else:
            self._sort_chase_col = col
            self._sort_chase_rev = False
        self._chase_rows.sort(key=lambda r: r.get(col, ""), reverse=self._sort_chase_rev)
        self._apply_chase_filter()

    def _export_chase_csv(self):
        if not self._chase_rows:
            messagebox.showwarning("Nothing to export", "Parse a Chase PDF statement first.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialfile="chase_foreign_transactions.csv")
        if not path:
            return
        fields = ["Transaction Date", "Description", "USD Amount",
                  "Foreign Amount", "Currency", "Exchange Rate", "Type"]
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
            w.writeheader()
            w.writerows(self._chase_rows)
        messagebox.showinfo("Exported", f"Saved {len(self._chase_rows)} rows to:\n{path}")

    # ── Remap Tab ─────────────────────────────────────────────────────────────
    def _build_remap_tab(self):
        f = self._remap_frame = tk.Frame(self._content, bg=COLORS["bg"])

        # Upload row
        upload_row = tk.Frame(f, bg=COLORS["bg"])
        upload_row.pack(fill="x", pady=(0, 10))
        tk.Button(upload_row, text="  Upload Capital One CSV  ",
                  bg=COLORS["accent2"], fg="white", font=("SF Pro Text", 12, "bold"),
                  relief="flat", bd=0, padx=14, pady=8, cursor="hand2",
                  activebackground="#6a4daa",
                  command=self._pick_csv).pack(side="left")
        self._csv_path_var = tk.StringVar(value="No file selected")
        self._csv_label = tk.Label(upload_row, textvariable=self._csv_path_var,
                                    bg=COLORS["bg"], fg=COLORS["muted"],
                                    font=("SF Pro Text", 11))
        self._csv_label.pack(side="left", padx=12)

        # Editable mapping panel
        map_outer = tk.Frame(f, bg=COLORS["panel"], pady=8, padx=14)
        map_outer.pack(fill="x", pady=(0, 8))

        header_row = tk.Frame(map_outer, bg=COLORS["panel"])
        header_row.pack(fill="x", pady=(0, 6))
        tk.Label(header_row, text="Category Mappings  (edit right column, then click Recalculate)",
                 bg=COLORS["panel"], fg=COLORS["muted"],
                 font=("SF Pro Text", 10, "bold")).pack(side="left")
        tk.Button(header_row, text="  Recalculate  ",
                  bg=COLORS["accent"], fg="white", font=("SF Pro Text", 10, "bold"),
                  relief="flat", bd=0, padx=10, pady=4, cursor="hand2",
                  activebackground="#3a7de8",
                  command=self._recalculate_mappings).pack(side="right")

        # Scrollable grid of editable entries
        map_canvas = tk.Canvas(map_outer, bg=COLORS["panel"], highlightthickness=0, height=160)
        map_scroll = ttk.Scrollbar(map_outer, orient="vertical", command=map_canvas.yview)
        map_canvas.configure(yscrollcommand=map_scroll.set)
        map_canvas.pack(side="left", fill="both", expand=True)
        map_scroll.pack(side="right", fill="y")

        map_frame = tk.Frame(map_canvas, bg=COLORS["panel"])
        map_canvas.create_window((0, 0), window=map_frame, anchor="nw")

        # Header labels
        tk.Label(map_frame, text="Capital One Category", bg=COLORS["panel"],
                 fg=COLORS["muted"], font=("SF Pro Text", 9, "bold"),
                 width=22, anchor="w").grid(row=0, column=0, padx=(0,4), pady=(0,4))
        tk.Label(map_frame, text="Your Label", bg=COLORS["panel"],
                 fg=COLORS["muted"], font=("SF Pro Text", 9, "bold"),
                 width=22, anchor="w").grid(row=0, column=1, padx=(0,4), pady=(0,4))
        tk.Label(map_frame, text="Capital One Category", bg=COLORS["panel"],
                 fg=COLORS["muted"], font=("SF Pro Text", 9, "bold"),
                 width=22, anchor="w").grid(row=0, column=2, padx=(16,4), pady=(0,4))
        tk.Label(map_frame, text="Your Label", bg=COLORS["panel"],
                 fg=COLORS["muted"], font=("SF Pro Text", 9, "bold"),
                 width=22, anchor="w").grid(row=0, column=3, padx=(0,4), pady=(0,4))

        self._mapping_entries = {}
        items = list(CATEGORY_MAP.items())
        half = (len(items) + 1) // 2
        for i, (src, dst) in enumerate(items):
            col_offset = 0 if i < half else 2
            row = 1 + (i if i < half else i - half)
            pad_left = 0 if col_offset == 0 else 16
            tk.Label(map_frame, text=src, bg=COLORS["panel"], fg=COLORS["text"],
                     font=("SF Pro Text", 10), width=22, anchor="w").grid(
                row=row, column=col_offset, padx=(pad_left, 4), pady=2, sticky="w")
            var = tk.StringVar(value=dst)
            entry = tk.Entry(map_frame, textvariable=var,
                             bg=COLORS["border"], fg=COLORS["text"],
                             insertbackground=COLORS["text"],
                             font=("SF Pro Text", 10), relief="flat", bd=0,
                             width=20)
            entry.grid(row=row, column=col_offset+1, padx=(0,4), pady=2, ipady=3, sticky="w")
            self._mapping_entries[src] = var

        map_frame.update_idletasks()
        map_canvas.configure(scrollregion=map_canvas.bbox("all"))

        # Filter
        filter_row = tk.Frame(f, bg=COLORS["bg"])
        filter_row.pack(fill="x", pady=(0, 6))
        tk.Label(filter_row, text="Filter:", bg=COLORS["bg"], fg=COLORS["muted"],
                 font=("SF Pro Text", 11)).pack(side="left")
        self._csv_filter = tk.StringVar()
        self._csv_filter.trace_add("write", lambda *_: self._apply_csv_filter())
        tk.Entry(filter_row, textvariable=self._csv_filter,
                 bg=COLORS["panel"], fg=COLORS["text"], insertbackground=COLORS["text"],
                 font=("SF Pro Text", 11), relief="flat", bd=0, width=28
                 ).pack(side="left", padx=(6,0), ipady=5, ipadx=6)

        # Table
        frame_tree = tk.Frame(f, bg=COLORS["border"], bd=1)
        frame_tree.pack(fill="both", expand=True, pady=(0, 10))

        self._csv_tree = ttk.Treeview(frame_tree, show="headings", style="Custom.Treeview")
        vsb = ttk.Scrollbar(frame_tree, orient="vertical", command=self._csv_tree.yview)
        hsb = ttk.Scrollbar(frame_tree, orient="horizontal", command=self._csv_tree.xview)
        self._csv_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self._csv_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame_tree.grid_rowconfigure(0, weight=1)
        frame_tree.grid_columnconfigure(0, weight=1)

        self._csv_tree.tag_configure("remapped", foreground=COLORS["success"])
        self._csv_tree.tag_configure("unmapped", foreground=COLORS["warning"])

        # Bottom bar
        bot = tk.Frame(f, bg=COLORS["bg"])
        bot.pack(fill="x")
        self._csv_count_lbl = tk.Label(bot, text="0 rows", bg=COLORS["bg"],
                                        fg=COLORS["muted"], font=("SF Pro Text", 11))
        self._csv_count_lbl.pack(side="left")
        tk.Button(bot, text="  Export Remapped CSV  ",
                  bg=COLORS["success"], fg="#0a1a10", font=("SF Pro Text", 12, "bold"),
                  relief="flat", bd=0, padx=14, pady=7, cursor="hand2",
                  activebackground="#2eb87a",
                  command=self._export_csv).pack(side="right")
        tk.Button(bot, text="  Reset  ",
                  bg=COLORS["danger"], fg="white", font=("SF Pro Text", 12, "bold"),
                  relief="flat", bd=0, padx=14, pady=7, cursor="hand2",
                  activebackground="#cc2222",
                  command=self._reset_csv).pack(side="right", padx=(0, 8))

    def _reset_csv(self):
        self._csv_rows = []
        self._csv_fields = []
        self._csv_path_var.set("No file selected")
        self._csv_label.config(fg=COLORS["muted"])
        self._csv_filter.set("")
        self._csv_tree.delete(*self._csv_tree.get_children())
        self._csv_tree.configure(columns=[])
        self._csv_count_lbl.config(text="0 rows", fg=COLORS["muted"])

    def _get_live_map(self):
        """Return current mapping from editable entries, falling back to CATEGORY_MAP."""
        if hasattr(self, "_mapping_entries"):
            return {k: v.get().strip() or k for k, v in self._mapping_entries.items()}
        return CATEGORY_MAP

    def _recalculate_mappings(self):
        """Re-apply current mapping entries to loaded CSV rows and refresh table."""
        if not self._csv_rows:
            messagebox.showinfo("No data", "Load a CSV file first.")
            return
        live_map = self._get_live_map()
        cat_col = next((c for c in self._csv_fields if c.strip().lower() == "category"), None)
        for row in self._csv_rows:
            orig = row.get(cat_col, "").strip() if cat_col else ""
            row["My Category"] = live_map.get(orig, orig or "Uncategorized")
        self._apply_csv_filter()
        self._csv_count_lbl.config(
            text=f"{len(self._csv_rows)} rows (recalculated)", fg=COLORS["success"])

    def _pick_csv(self):
        path = filedialog.askopenfilename(
            title="Select Capital One CSV Export",
            filetypes=[("CSV files","*.csv"),("All files","*.*")])
        if not path:
            return
        self._csv_path_var.set(os.path.basename(path))
        self._csv_label.config(fg=COLORS["text"])
        try:
            rows, fields = remap_csv(path)
        except Exception as e:
            messagebox.showerror("CSV Error", str(e))
            return
        self._csv_rows  = rows
        self._csv_fields = fields
        # Apply live mapping immediately on load
        live_map = self._get_live_map() if hasattr(self, "_mapping_entries") else CATEGORY_MAP
        cat_col = next((c for c in fields if c.strip().lower() == "category"), None)
        for row in rows:
            orig = row.get(cat_col, "").strip() if cat_col else ""
            row["My Category"] = live_map.get(orig, orig or "Uncategorized")

        # Configure columns dynamically
        self._csv_tree.configure(columns=fields)
        widths = {"Description": 220, "My Category": 140}
        for col in fields:
            w = widths.get(col, 110)
            self._csv_tree.heading(col, text=col, anchor="w")
            self._csv_tree.column(col, width=w, minwidth=40, anchor="w")

        self._apply_csv_filter()

    def _apply_csv_filter(self):
        query = self._csv_filter.get().lower()
        filtered = [r for r in self._csv_rows
                    if not query or any(query in str(v).lower() for v in r.values())]
        self._csv_count_lbl.config(text=f"{len(filtered)} rows", fg=COLORS["muted"])
        self._csv_tree.delete(*self._csv_tree.get_children())
        live_map = self._get_live_map() if hasattr(self, "_mapping_entries") else CATEGORY_MAP
        for r in filtered:
            orig = r.get("Category","").strip()
            tag = "remapped" if orig in live_map else "unmapped"
            self._csv_tree.insert("", "end",
                values=[r.get(f,"") for f in self._csv_fields],
                tags=(tag,))

    def _export_csv(self):
        if not self._csv_rows:
            messagebox.showwarning("Nothing to export", "Load a CSV file first.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files","*.csv")],
            initialfile="remapped_transactions.csv")
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=self._csv_fields, extrasaction="ignore")
            w.writeheader()
            w.writerows(self._csv_rows)
        messagebox.showinfo("Exported", f"Saved {len(self._csv_rows)} rows to:\n{path}")


    # ── Chase Remap Tab ───────────────────────────────────────────────────────
    def _build_chase_remap_tab(self):
        f = self._chase_remap_frame = tk.Frame(self._content, bg=COLORS["bg"])

        upload_row = tk.Frame(f, bg=COLORS["bg"])
        upload_row.pack(fill="x", pady=(0, 10))
        tk.Button(upload_row, text="  Upload Chase CSV  ",
                  bg="#1a6fa8", fg="white", font=("SF Pro Text", 12, "bold"),
                  relief="flat", bd=0, padx=14, pady=8, cursor="hand2",
                  activebackground="#155d8e",
                  command=self._pick_chase_csv).pack(side="left")
        self._chase_csv_path_var = tk.StringVar(value="No file selected")
        self._chase_csv_label = tk.Label(upload_row, textvariable=self._chase_csv_path_var,
                                          bg=COLORS["bg"], fg=COLORS["muted"],
                                          font=("SF Pro Text", 11))
        self._chase_csv_label.pack(side="left", padx=12)

        map_outer = tk.Frame(f, bg=COLORS["panel"], pady=8, padx=14)
        map_outer.pack(fill="x", pady=(0, 8))

        header_row = tk.Frame(map_outer, bg=COLORS["panel"])
        header_row.pack(fill="x", pady=(0, 6))
        tk.Label(header_row, text="Category Mappings  (edit right column, then click Recalculate)",
                 bg=COLORS["panel"], fg=COLORS["muted"],
                 font=("SF Pro Text", 10, "bold")).pack(side="left")
        tk.Button(header_row, text="  Recalculate  ",
                  bg=COLORS["accent"], fg="white", font=("SF Pro Text", 10, "bold"),
                  relief="flat", bd=0, padx=10, pady=4, cursor="hand2",
                  activebackground="#3a7de8",
                  command=self._recalculate_chase_mappings).pack(side="right")

        map_canvas = tk.Canvas(map_outer, bg=COLORS["panel"], highlightthickness=0, height=160)
        map_scroll = ttk.Scrollbar(map_outer, orient="vertical", command=map_canvas.yview)
        map_canvas.configure(yscrollcommand=map_scroll.set)
        map_canvas.pack(side="left", fill="both", expand=True)
        map_scroll.pack(side="right", fill="y")

        map_frame = tk.Frame(map_canvas, bg=COLORS["panel"])
        map_canvas.create_window((0, 0), window=map_frame, anchor="nw")

        tk.Label(map_frame, text="Chase Category", bg=COLORS["panel"],
                 fg=COLORS["muted"], font=("SF Pro Text", 9, "bold"),
                 width=22, anchor="w").grid(row=0, column=0, padx=(0,4), pady=(0,4))
        tk.Label(map_frame, text="Your Label", bg=COLORS["panel"],
                 fg=COLORS["muted"], font=("SF Pro Text", 9, "bold"),
                 width=22, anchor="w").grid(row=0, column=1, padx=(0,4), pady=(0,4))
        tk.Label(map_frame, text="Chase Category", bg=COLORS["panel"],
                 fg=COLORS["muted"], font=("SF Pro Text", 9, "bold"),
                 width=22, anchor="w").grid(row=0, column=2, padx=(16,4), pady=(0,4))
        tk.Label(map_frame, text="Your Label", bg=COLORS["panel"],
                 fg=COLORS["muted"], font=("SF Pro Text", 9, "bold"),
                 width=22, anchor="w").grid(row=0, column=3, padx=(0,4), pady=(0,4))

        self._chase_mapping_entries = {}
        items = list(CHASE_CATEGORY_MAP.items())
        half = (len(items) + 1) // 2
        for i, (src, dst) in enumerate(items):
            col_offset = 0 if i < half else 2
            row = 1 + (i if i < half else i - half)
            pad_left = 0 if col_offset == 0 else 16
            tk.Label(map_frame, text=src, bg=COLORS["panel"], fg=COLORS["text"],
                     font=("SF Pro Text", 10), width=22, anchor="w").grid(
                row=row, column=col_offset, padx=(pad_left, 4), pady=2, sticky="w")
            var = tk.StringVar(value=dst)
            entry = tk.Entry(map_frame, textvariable=var,
                             bg=COLORS["border"], fg=COLORS["text"],
                             insertbackground=COLORS["text"],
                             font=("SF Pro Text", 10), relief="flat", bd=0, width=20)
            entry.grid(row=row, column=col_offset+1, padx=(0,4), pady=2, ipady=3, sticky="w")
            self._chase_mapping_entries[src] = var

        map_frame.update_idletasks()
        map_canvas.configure(scrollregion=map_canvas.bbox("all"))

        filter_row = tk.Frame(f, bg=COLORS["bg"])
        filter_row.pack(fill="x", pady=(0, 6))
        tk.Label(filter_row, text="Filter:", bg=COLORS["bg"], fg=COLORS["muted"],
                 font=("SF Pro Text", 11)).pack(side="left")
        self._chase_csv_filter = tk.StringVar()
        self._chase_csv_filter.trace_add("write", lambda *_: self._apply_chase_csv_filter())
        tk.Entry(filter_row, textvariable=self._chase_csv_filter,
                 bg=COLORS["panel"], fg=COLORS["text"], insertbackground=COLORS["text"],
                 font=("SF Pro Text", 11), relief="flat", bd=0, width=28
                 ).pack(side="left", padx=(6,0), ipady=5, ipadx=6)

        frame_tree = tk.Frame(f, bg=COLORS["border"], bd=1)
        frame_tree.pack(fill="both", expand=True, pady=(0, 10))

        self._chase_csv_tree = ttk.Treeview(frame_tree, show="headings",
                                             style="Custom.Treeview")
        vsb = ttk.Scrollbar(frame_tree, orient="vertical", command=self._chase_csv_tree.yview)
        hsb = ttk.Scrollbar(frame_tree, orient="horizontal", command=self._chase_csv_tree.xview)
        self._chase_csv_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self._chase_csv_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame_tree.grid_rowconfigure(0, weight=1)
        frame_tree.grid_columnconfigure(0, weight=1)

        self._chase_csv_tree.tag_configure("remapped", foreground=COLORS["success"])
        self._chase_csv_tree.tag_configure("unmapped", foreground=COLORS["warning"])

        bot = tk.Frame(f, bg=COLORS["bg"])
        bot.pack(fill="x")
        self._chase_csv_count_lbl = tk.Label(bot, text="0 rows", bg=COLORS["bg"],
                                              fg=COLORS["muted"], font=("SF Pro Text", 11))
        self._chase_csv_count_lbl.pack(side="left")
        tk.Button(bot, text="  Export Remapped CSV  ",
                  bg=COLORS["success"], fg="#0a1a10", font=("SF Pro Text", 12, "bold"),
                  relief="flat", bd=0, padx=14, pady=7, cursor="hand2",
                  activebackground="#2eb87a",
                  command=self._export_chase_remap_csv).pack(side="right")
        tk.Button(bot, text="  Reset  ",
                  bg=COLORS["danger"], fg="white", font=("SF Pro Text", 12, "bold"),
                  relief="flat", bd=0, padx=14, pady=7, cursor="hand2",
                  activebackground="#cc2222",
                  command=self._reset_chase_csv).pack(side="right", padx=(0, 8))

        self._chase_csv_rows = []
        self._chase_csv_fields = []

    def _get_chase_live_map(self):
        if hasattr(self, "_chase_mapping_entries"):
            return {k: v.get().strip() or k for k, v in self._chase_mapping_entries.items()}
        return CHASE_CATEGORY_MAP

    def _recalculate_chase_mappings(self):
        if not self._chase_csv_rows:
            messagebox.showinfo("No data", "Load a Chase CSV file first.")
            return
        live_map = self._get_chase_live_map()
        cat_col = next((c for c in self._chase_csv_fields if c.strip().lower() == "category"), None)
        for row in self._chase_csv_rows:
            orig = row.get(cat_col, "").strip() if cat_col else ""
            row["My Category"] = live_map.get(orig, orig or "Uncategorized")
        self._apply_chase_csv_filter()
        self._chase_csv_count_lbl.config(
            text=f"{len(self._chase_csv_rows)} rows (recalculated)", fg=COLORS["success"])

    def _pick_chase_csv(self):
        path = filedialog.askopenfilename(
            title="Select Chase CSV Export",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if not path:
            return
        self._chase_csv_path_var.set(os.path.basename(path))
        self._chase_csv_label.config(fg=COLORS["text"])
        try:
            live_map = self._get_chase_live_map()
            rows, fields = remap_chase_csv(path, live_map)
        except Exception as e:
            messagebox.showerror("CSV Error", str(e))
            return
        self._chase_csv_rows = rows
        self._chase_csv_fields = fields
        self._chase_csv_tree.configure(columns=fields)
        widths = {"Description": 220, "My Category": 140, "Memo": 20}
        for col in fields:
            self._chase_csv_tree.heading(col, text=col, anchor="w")
            self._chase_csv_tree.column(col, width=widths.get(col, 110), minwidth=40, anchor="w")
        self._apply_chase_csv_filter()

    def _apply_chase_csv_filter(self):
        query = self._chase_csv_filter.get().lower()
        filtered = [r for r in self._chase_csv_rows
                    if not query or any(query in str(v).lower() for v in r.values())]
        self._chase_csv_count_lbl.config(text=f"{len(filtered)} rows", fg=COLORS["muted"])
        self._chase_csv_tree.delete(*self._chase_csv_tree.get_children())
        live_map = self._get_chase_live_map()
        for r in filtered:
            orig = r.get("Category", "").strip()
            tag = "remapped" if orig in live_map else "unmapped"
            self._chase_csv_tree.insert("", "end",
                values=[r.get(f, "") for f in self._chase_csv_fields],
                tags=(tag,))

    def _reset_chase_csv(self):
        self._chase_csv_rows = []
        self._chase_csv_fields = []
        self._chase_csv_path_var.set("No file selected")
        self._chase_csv_label.config(fg=COLORS["muted"])
        self._chase_csv_filter.set("")
        self._chase_csv_tree.delete(*self._chase_csv_tree.get_children())
        self._chase_csv_tree.configure(columns=[])
        self._chase_csv_count_lbl.config(text="0 rows", fg=COLORS["muted"])

    def _export_chase_remap_csv(self):
        if not self._chase_csv_rows:
            messagebox.showwarning("Nothing to export", "Load a Chase CSV file first.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialfile="chase_remapped.csv")
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=self._chase_csv_fields, extrasaction="ignore")
            w.writeheader()
            w.writerows(self._chase_csv_rows)
        messagebox.showinfo("Exported", f"Saved {len(self._chase_csv_rows)} rows to:\n{path}")

# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    try:
        app = App()
        app.mainloop()
    except Exception as _e:
        import traceback
        try:
            import tkinter as _tk
            from tkinter import messagebox as _mb
            _r = _tk.Tk(); _r.withdraw()
            _mb.showerror("Startup Error",
                f"The app crashed on startup:\n\n{traceback.format_exc()}")
        except Exception:
            pass
