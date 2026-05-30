#!/usr/bin/env python3
"""Gurudev Motors — Best Dealer In Town Sales Score Calculator."""

import tkinter as tk
from tkinter import font as tkfont, ttk, messagebox
import json, os, sys, datetime, base64, io

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    _HAS_XL = True
except ImportError:
    _HAS_XL = False

# ── Embedded Škoda logo (80×80 PNG, base64) ──────────────────────────────────
_LOGO_B64 = """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"""

# ── Palette ───────────────────────────────────────────────────────────────────
BG       = "#1a1a2e"
PANEL    = "#16213e"
HDR_BG   = "#0f3460"
ACCENT   = "#4caf50"
ACCENT2  = "#e94560"
TEXT     = "#e0e0e0"
DIM      = "#9e9e9e"
INP_BG   = "#0d2137"
GOLD     = "#ffd700"
RED      = "#e94560"
SEP      = "#2a3a5c"

SLAB_COLOR = {1.0: ACCENT, 0.8: GOLD, 0.0: RED}
SLAB_LBL   = {1.0: "Slab 1 — 100%", 0.8: "Slab 2 — 80%", 0.0: "Slab 3 — 0%"}

# ── KPI definitions ───────────────────────────────────────────────────────────
KPIS = [
    {"name": "Sales Certified Manpower",           "weight": 100, "unit": "%",
     "hint": "● ≥80%  ◑ 55–79%  ○ <55%",
     "slabs": [(1.0, lambda v: v >= 80,            "≥ 80%"),
               (0.8, lambda v: 55 <= v <= 79,      "55% – 79%"),
               (0.0, lambda v: v < 55,             "< 55%")]},

    {"name": "Central Web-In Lead to Retail Ratio","weight": 100, "unit": "%",
     "hint": "● ≥2.25%  ◑ 1.40–2.24%  ○ <1.40%",
     "slabs": [(1.0, lambda v: v >= 2.25,          "≥ 2.25%"),
               (0.8, lambda v: 1.40 <= v < 2.25,   "1.40% – 2.24%"),
               (0.0, lambda v: v < 1.40,           "< 1.40%")]},

    {"name": "Mystery Shopping",                   "weight": 50,  "unit": "%",
     "hint": "● ≥80%  ◑ 60–79%  ○ <60%",
     "slabs": [(1.0, lambda v: v >= 80,            "≥ 80%"),
               (0.8, lambda v: 60 <= v < 80,       "60% – 79%"),
               (0.0, lambda v: v < 60,             "< 60%")]},

    {"name": "Escalation Index Sales",             "weight": 50,  "unit": "",
     "hint": "● ≤0.20  ◑ 0.21–0.94  ○ >0.94",
     "slabs": [(1.0, lambda v: v <= 0.20,          "≤ 0.20"),
               (0.8, lambda v: 0.21 <= v <= 0.94,  "0.21 – 0.94"),
               (0.0, lambda v: v > 0.94,           "> 0.94")]},

    {"name": "PSI Sales",                          "weight": 100, "unit": "",
     "hint": "● ≥4.85  ◑ 4.60–4.84  ○ <4.60",
     "slabs": [(1.0, lambda v: v >= 4.85,          "≥ 4.85"),
               (0.8, lambda v: 4.60 <= v < 4.85,   "4.60 – 4.84"),
               (0.0, lambda v: v < 4.60,           "< 4.60")]},

    {"name": "CX Sales",                           "weight": 100, "unit": "",
     "hint": "● ≥4.85  ◑ 4.60–4.84  ○ <4.60",
     "slabs": [(1.0, lambda v: v >= 4.85,          "≥ 4.85"),
               (0.8, lambda v: 4.60 <= v < 4.85,   "4.60 – 4.84"),
               (0.0, lambda v: v < 4.60,           "< 4.60")]},
]

TOTAL_WT = sum(k["weight"] for k in KPIS)   # 500

MONTHS = ["January","February","March","April","May","June",
          "July","August","September","October","November","December"]

# Q2 business rule: April excluded — only May & Jun count toward Q2 average
QUARTERS = {
    "Q1": {"months": ["January","February","March"],
           "counted": ["January","February","March"],
           "label": "Q1 Average  (Jan – Mar)"},
    "Q2": {"months": ["April","May","June"],
           "counted": ["May","June"],
           "label": "Q2 Average  (May – Jun)"},
    "Q3": {"months": ["July","August","September"],
           "counted": ["July","August","September"],
           "label": "Q3 Average  (Jul – Sep)"},
    "Q4": {"months": ["October","November","December"],
           "counted": ["October","November","December"],
           "label": "Q4 Average  (Oct – Dec)"},
}


def get_quarter(month):
    for qk, qv in QUARTERS.items():
        if month in qv["months"]:
            return qk, qv
    return None, None


def compute_slab(kpi, value):
    for pct, cond, lbl in kpi["slabs"]:
        if cond(value):
            return pct, SLAB_LBL[pct], int(round(pct * kpi["weight"]))
    return 0.0, SLAB_LBL[0.0], 0


# ── Data persistence ──────────────────────────────────────────────────────────
if getattr(sys, "frozen", False):
    _data_dir = os.path.expanduser("~/.skoda_bdit")
else:
    _data_dir = os.path.dirname(os.path.abspath(__file__))

os.makedirs(_data_dir, exist_ok=True)
DATA_FILE  = os.path.join(_data_dir, "skoda_scores.json")
EXCEL_FILE = os.path.join(_data_dir, "skoda_scores.xlsx")


def _load():
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE) as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def _save(data):
    try:
        with open(DATA_FILE, "w") as f:
            json.dump(data, f, indent=2)
    except Exception as e:
        print("Save error:", e)


def _save_excel(db):
    """Rewrite the full Excel workbook from the in-memory db."""
    if not _HAS_XL:
        return
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "BDIT Scores"

        # ── Styles ────────────────────────────────────────────────────────────
        hdr_fill = PatternFill("solid", fgColor="0F3460")
        hdr_font = Font(bold=True, color="FFFFFF", size=10)
        hdr_aln  = Alignment(horizontal="center", vertical="center",
                             wrap_text=True)
        val_aln  = Alignment(horizontal="center", vertical="center")
        thin     = Side(style="thin", color="2A3A5C")
        border   = Border(left=thin, right=thin, top=thin, bottom=thin)

        slab_fill = {
            1.0: PatternFill("solid", fgColor="1B5E20"),   # green
            0.8: PatternFill("solid", fgColor="7B5E00"),   # amber
            0.0: PatternFill("solid", fgColor="7B1A2C"),   # red
        }

        # ── Header row ────────────────────────────────────────────────────────
        headers = ["User Name", "Year", "Month"]
        for kpi in KPIS:
            n = kpi["name"]
            headers += [f"{n}\nValue", f"{n}\nSlab", f"{n}\nScore"]
        headers.append("Total\nScore")
        ws.append(headers)

        for cell in ws[1]:
            cell.font   = hdr_font
            cell.fill   = hdr_fill
            cell.alignment = hdr_aln
            cell.border = border

        ws.row_dimensions[1].height = 42

        # ── Data rows ─────────────────────────────────────────────────────────
        month_order = {m: i for i, m in enumerate(MONTHS)}

        for user in sorted(db):
            for year in sorted(db[user]):
                month_data = db[user][year]
                for month in sorted(month_data, key=lambda m: month_order.get(m, 99)):
                    kpis = month_data[month]
                    row  = [user, int(year), month]
                    total = 0
                    pcts  = []
                    for kpi in KPIS:
                        d = kpis.get(kpi["name"], {})
                        row  += [d.get("value", ""), d.get("label", ""), d.get("score", "")]
                        total += d.get("score", 0)
                        pcts.append(d.get("pct", None))
                    row.append(total)
                    ws.append(row)

                    # Colour score cells by slab
                    r = ws.max_row
                    for col_i, pct in enumerate(pcts):
                        if pct is not None and pct in slab_fill:
                            score_col = 4 + col_i * 3  # every 3rd column starting at col 6
                            ws.cell(r, score_col).fill = slab_fill[pct]

                    for cell in ws[r]:
                        cell.alignment = val_aln
                        cell.border    = border

        # ── Column widths ─────────────────────────────────────────────────────
        ws.column_dimensions["A"].width = 18  # User
        ws.column_dimensions["B"].width = 6   # Year
        ws.column_dimensions["C"].width = 11  # Month
        col_letter = [chr(c) for c in range(ord("D"), ord("D") + len(KPIS) * 3 + 1)]
        for i, ltr in enumerate(col_letter):
            ws.column_dimensions[ltr].width = 8 if (i % 3 == 2) else 10
        # Total score column
        last = openpyxl.utils.get_column_letter(4 + len(KPIS) * 3)
        ws.column_dimensions[last].width = 8

        # Freeze header row
        ws.freeze_panes = "A2"

        wb.save(EXCEL_FILE)
    except Exception as e:
        print("Excel save error:", e)


# ── Shared font cache (populated once Tk root exists) ─────────────────────────
F = {}


def _init_fonts():
    F["hdr"]   = tkfont.Font(family="Helvetica", size=18, weight="bold")
    F["b14"]   = tkfont.Font(family="Helvetica", size=14, weight="bold")
    F["b12"]   = tkfont.Font(family="Helvetica", size=12, weight="bold")
    F["b11"]   = tkfont.Font(family="Helvetica", size=11, weight="bold")
    F["r11"]   = tkfont.Font(family="Helvetica", size=11)
    F["r9"]    = tkfont.Font(family="Helvetica", size=9)
    F["score"] = tkfont.Font(family="Helvetica", size=16, weight="bold")


def _style_combobox():
    s = ttk.Style()
    s.theme_use("default")
    s.configure("Dark.TCombobox",
                fieldbackground=INP_BG, background=INP_BG,
                foreground=TEXT, selectbackground=ACCENT,
                selectforeground=TEXT, arrowcolor=ACCENT,
                borderwidth=1, relief="flat")
    s.map("Dark.TCombobox",
          fieldbackground=[("readonly", INP_BG)],
          foreground=[("readonly", TEXT)])


def _btn(parent, text, cmd, color, font_key="b12", px=24, py=9):
    """Label-based button — macOS tk.Button ignores bg/fg in the Aqua theme."""
    lbl = tk.Label(parent, text=text, bg=color, fg="white",
                   font=F[font_key], padx=px, pady=py, cursor="hand2")
    lbl.bind("<Button-1>", lambda e: cmd())
    return lbl


# ── App ───────────────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Skoda Best Dealer in Town App")
        self.configure(bg=BG)
        self.resizable(True, True)

        _init_fonts()
        _style_combobox()

        self.geometry("860x660")
        self.minsize(820, 520)
        self._db = _load()

        container = tk.Frame(self, bg=BG)
        container.pack(fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self._pages = {}
        for Cls in (InputPage, ResultsPage):
            pg = Cls(container, self)
            self._pages[Cls.__name__] = pg
            pg.grid(row=0, column=0, sticky="nsew")

        self.show("InputPage")

    def show(self, name, **kw):
        pg = self._pages[name]
        if hasattr(pg, "on_show"):
            pg.on_show(**kw)
        pg.tkraise()

    def submit(self, username, month, year, raw_values):
        """Validate → compute scores → persist → show results."""
        scores = {}
        for i, kpi in enumerate(KPIS):
            pct, lbl, sc = compute_slab(kpi, raw_values[i])
            scores[kpi["name"]] = {
                "value": raw_values[i],
                "pct": pct,
                "label": lbl,
                "score": sc,
            }
        yr = str(year)
        self._db.setdefault(username, {}).setdefault(yr, {})[month] = scores
        _save(self._db)
        _save_excel(self._db)
        self.show("ResultsPage", username=username, month=month, year=yr, scores=scores)

    def get_user_year_data(self, username, year):
        return self._db.get(username, {}).get(str(year), {})


# ── Shared header widget ──────────────────────────────────────────────────────
_LOGO_PHOTO = None   # cached PhotoImage (must stay alive)

def make_header(parent):
    global _LOGO_PHOTO

    hdr = tk.Frame(parent, bg=HDR_BG, pady=10)
    hdr.pack(fill="x")

    # ── Decode & cache the logo image once ───────────────────────────────────
    if _LOGO_PHOTO is None:
        try:
            from PIL import Image, ImageTk
            raw = base64.b64decode(_LOGO_B64)
            pil_img = Image.open(io.BytesIO(raw)).convert("RGBA")
            pil_img = pil_img.resize((72, 72), Image.LANCZOS)
            _LOGO_PHOTO = ImageTk.PhotoImage(pil_img)
        except Exception:
            _LOGO_PHOTO = False   # PIL not available; skip image

    # ── Logo + centred text ───────────────────────────────────────────────────
    inner = tk.Frame(hdr, bg=HDR_BG)
    inner.pack(anchor="center")

    if _LOGO_PHOTO:
        tk.Label(inner, image=_LOGO_PHOTO, bg=HDR_BG).pack(side="left", padx=(0, 16))

    text_col = tk.Frame(inner, bg=HDR_BG)
    text_col.pack(side="left")

    tk.Label(text_col, text="Gurudev Motors Pvt. Ltd",
             bg=HDR_BG, fg=TEXT,
             font=tkfont.Font(family="Helvetica", size=20, weight="bold")).pack(anchor="center")
    tk.Label(text_col, text="Best Dealer in Town Dashboard",
             bg=HDR_BG, fg=ACCENT,
             font=tkfont.Font(family="Helvetica", size=13, weight="bold")).pack(anchor="center")
    tk.Label(text_col, text="Designed by Sabestina Kennet",
             bg=HDR_BG, fg="#6a7fa8",
             font=tkfont.Font(family="Helvetica", size=9, slant="italic")).pack(anchor="center")

    return hdr


# ── Page 1 — Input ────────────────────────────────────────────────────────────
class InputPage(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=BG)
        self._app = app
        self._kpi_vars = []      # StringVar per KPI
        self._slab_widgets = []  # (slab_var, score_var, slab_lbl, score_lbl)
        self._build()

    def _build(self):
        make_header(self)

        # ── User / month / year bar ───────────────────────────────────────────
        info = tk.Frame(self, bg=PANEL, pady=10)
        info.pack(fill="x", padx=0)

        def lbl(parent, text):
            return tk.Label(parent, text=text, bg=PANEL, fg=TEXT, font=F["b11"])

        def entry(parent, var, w=18):
            return tk.Entry(parent, textvariable=var,
                            bg=INP_BG, fg=TEXT, insertbackground=TEXT,
                            font=F["b11"], width=w, relief="flat",
                            highlightthickness=1,
                            highlightbackground=ACCENT, highlightcolor=ACCENT)

        row = tk.Frame(info, bg=PANEL)
        row.pack(padx=18)

        lbl(row, "User Name:").grid(row=0, column=0, padx=(0, 6), sticky="w")
        self._user_var = tk.StringVar()
        entry(row, self._user_var, 22).grid(row=0, column=1, padx=(0, 24), sticky="w")

        lbl(row, "Month:").grid(row=0, column=2, padx=(0, 6), sticky="w")
        self._month_var = tk.StringVar(value=MONTHS[datetime.date.today().month - 1])
        cb = ttk.Combobox(row, textvariable=self._month_var, values=MONTHS,
                          width=13, state="readonly", style="Dark.TCombobox",
                          font=F["r11"])
        # On macOS the custom style can suppress native click events — force open
        cb.bind("<Button-1>", lambda e: cb.after(10, cb.event_generate, "<Down>"))
        cb.grid(row=0, column=3, padx=(0, 24), sticky="w")

        lbl(row, "Year:").grid(row=0, column=4, padx=(0, 6), sticky="w")
        self._year_var = tk.StringVar(value=str(datetime.date.today().year))
        entry(row, self._year_var, 6).grid(row=0, column=5, sticky="w")

        # ── Column headers ────────────────────────────────────────────────────
        # Widths (chars): KPI=30, Wt=5, Value=13, SlabHit=16, Score=7
        COL_W = [30, 5, 13, 16, 7]
        cols = tk.Frame(self, bg=SEP, pady=5)
        cols.pack(fill="x", padx=18, pady=(10, 0))
        for text, w, anc in zip(["KPI", "Wt.", "Your Value", "Slab Hit", "Score"],
                                 COL_W, ["w", "w", "w", "w", "e"]):
            tk.Label(cols, text=text, bg=SEP, fg=DIM,
                     font=F["b11"], width=w, anchor=anc, padx=5).pack(side="left")

        # ── KPI rows ──────────────────────────────────────────────────────────
        for idx, kpi in enumerate(KPIS):
            bg = PANEL if idx % 2 == 0 else INP_BG
            card = tk.Frame(self, bg=bg, pady=2, padx=4)
            card.pack(fill="x", padx=18, pady=1)
            card.rowconfigure(0, weight=1)   # allow vertical centering
            card.columnconfigure(0, minsize=COL_W[0] * 7)  # lock column 0 width

            # KPI name + hint — column 0, height determined by children
            nf = tk.Frame(card, bg=bg)
            nf.grid(row=0, column=0, padx=5, pady=8, sticky="nsew")
            tk.Label(nf, text=kpi["name"], bg=bg, fg=TEXT, font=F["b11"],
                     anchor="w", wraplength=195, justify="left").pack(anchor="w")
            tk.Label(nf, text=kpi["hint"], bg=bg, fg="#b0bfd4", font=F["r9"],
                     anchor="w").pack(anchor="w", pady=(2, 0))

            # Weight — vertically centred
            tk.Label(card, text=str(kpi["weight"]), bg=bg,
                     fg=DIM, font=F["r11"], width=COL_W[1], anchor="w").grid(
                row=0, column=1, padx=4, sticky="ns")

            # Entry + unit — vertically centred
            sv = tk.StringVar()
            sv.trace_add("write", lambda *a, i=idx, v=sv: self._on_kpi_change(i, v))
            self._kpi_vars.append(sv)

            ef = tk.Frame(card, bg=bg)
            ef.grid(row=0, column=2, padx=5, sticky="ns")
            tk.Entry(ef, textvariable=sv, bg=INP_BG, fg=TEXT, insertbackground=TEXT,
                     font=F["b11"], width=9, relief="flat",
                     highlightthickness=1, highlightbackground=ACCENT,
                     highlightcolor=ACCENT).pack(side="left", pady=6)
            # Always include unit label (width=2) so column 2 is consistent
            tk.Label(ef, text=kpi["unit"], bg=bg, fg=DIM,
                     font=F["r11"], width=2).pack(side="left", padx=(2, 0), pady=6)

            # Slab label — vertically centred
            sl_var = tk.StringVar(value="—")
            sl_lbl = tk.Label(card, textvariable=sl_var, bg=bg, fg=DIM,
                              font=F["b11"], width=COL_W[3], anchor="w")
            sl_lbl.grid(row=0, column=3, padx=5, sticky="ns")

            # Score label — vertically centred, right-aligned
            sc_var = tk.StringVar(value="—")
            sc_lbl = tk.Label(card, textvariable=sc_var, bg=bg, fg=DIM,
                              font=F["b11"], width=COL_W[4], anchor="e")
            sc_lbl.grid(row=0, column=4, padx=5, sticky="ns")

            self._slab_widgets.append((sl_var, sc_var, sl_lbl, sc_lbl))

        # ── Total bar ─────────────────────────────────────────────────────────
        tk.Frame(self, bg=ACCENT, height=2).pack(fill="x", padx=18, pady=(8, 0))
        tot = tk.Frame(self, bg=HDR_BG, pady=9)
        tot.pack(fill="x", padx=18, pady=(0, 4))

        tk.Label(tot, text="TOTAL SCORE", bg=HDR_BG, fg=TEXT,
                 font=F["b12"]).pack(side="left", padx=16)
        tk.Label(tot, text=f"/ {TOTAL_WT}", bg=HDR_BG, fg=DIM,
                 font=F["r11"]).pack(side="right", padx=16)

        self._total_var = tk.StringVar(value="—")
        tk.Label(tot, textvariable=self._total_var, bg=HDR_BG, fg=GOLD,
                 font=F["score"]).pack(side="right", padx=4)
        self._pct_var = tk.StringVar(value="")
        tk.Label(tot, textvariable=self._pct_var, bg=HDR_BG, fg=ACCENT,
                 font=F["b12"]).pack(side="right", padx=8)

        # ── Buttons ───────────────────────────────────────────────────────────
        btns = tk.Frame(self, bg=BG, pady=8)
        btns.pack(fill="x", padx=18)

        _btn(btns, "Submit", self._on_submit, ACCENT).pack(side="right", padx=(6, 0))
        _btn(btns, "Reset", self._reset, ACCENT2).pack(side="right")

        tk.Label(self, text="Slab 1 = 100% of weight  |  Slab 2 = 80%  |  Slab 3 = 0%",
                 bg=BG, fg=DIM, font=F["r9"]).pack(pady=(0, 8))

    # ── Handlers ──────────────────────────────────────────────────────────────
    def _on_kpi_change(self, idx, var):
        raw = var.get().strip()
        sl_var, sc_var, sl_lbl, sc_lbl = self._slab_widgets[idx]

        if not raw:
            sl_var.set("—"); sl_lbl.config(fg=DIM)
            sc_var.set("—"); sc_lbl.config(fg=DIM)
            self._refresh_total(); return
        try:
            val = float(raw)
        except ValueError:
            sl_var.set("Invalid"); sl_lbl.config(fg=RED)
            sc_var.set("—"); sc_lbl.config(fg=DIM)
            self._refresh_total(); return

        pct, lbl, score = compute_slab(KPIS[idx], val)
        col = SLAB_COLOR[pct]
        sl_var.set(lbl); sl_lbl.config(fg=col)
        sc_var.set(str(score)); sc_lbl.config(fg=col)
        self._refresh_total()

    def _refresh_total(self):
        total, all_ok = 0, True
        for i, sv in enumerate(self._kpi_vars):
            raw = sv.get().strip()
            if not raw:
                all_ok = False; continue
            try:
                total += compute_slab(KPIS[i], float(raw))[2]
            except ValueError:
                all_ok = False
        if all_ok:
            self._total_var.set(str(total))
            self._pct_var.set(f"{total / TOTAL_WT * 100:.1f}%")
        elif total:
            self._total_var.set(f"{total} (partial)")
            self._pct_var.set("")
        else:
            self._total_var.set("—"); self._pct_var.set("")

    def _on_submit(self):
        username = self._user_var.get().strip()
        if not username:
            messagebox.showerror("Missing", "Please enter a user name.", parent=self)
            return

        month = self._month_var.get()

        try:
            year = int(self._year_var.get().strip())
        except ValueError:
            messagebox.showerror("Invalid", "Year must be a number.", parent=self)
            return

        raw_values = []
        for i, sv in enumerate(self._kpi_vars):
            raw = sv.get().strip()
            if not raw:
                messagebox.showerror(
                    "Missing value",
                    f"Please enter a value for:\n{KPIS[i]['name']}",
                    parent=self)
                return
            try:
                raw_values.append(float(raw))
            except ValueError:
                messagebox.showerror(
                    "Invalid value",
                    f"Invalid number for:\n{KPIS[i]['name']}",
                    parent=self)
                return

        self._app.submit(username, month, year, raw_values)

    def _reset(self):
        for sv in self._kpi_vars:
            sv.set("")
        for sl_var, sc_var, sl_lbl, sc_lbl in self._slab_widgets:
            sl_var.set("—"); sl_lbl.config(fg=DIM)
            sc_var.set("—"); sc_lbl.config(fg=DIM)
        self._total_var.set("—"); self._pct_var.set("")

    def on_show(self):
        pass  # nothing to refresh; keep user's last entries intact


# ── Page 2 — Results ──────────────────────────────────────────────────────────
class ResultsPage(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=BG)
        self._app = app
        self._content = None   # rebuilt on each show

    def on_show(self, username="", month="", year="", scores=None):
        if self._content:
            self._content.destroy()

        self._content = tk.Frame(self, bg=BG)
        self._content.pack(fill="both", expand=True)

        make_header(self._content)

        # ── Scrollable body ───────────────────────────────────────────────────
        canvas = tk.Canvas(self._content, bg=BG, highlightthickness=0)
        sb = ttk.Scrollbar(self._content, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        body = tk.Frame(canvas, bg=BG)
        win_id = canvas.create_window((0, 0), window=body, anchor="nw")

        def _on_canvas_resize(e):
            canvas.itemconfig(win_id, width=e.width)
        canvas.bind("<Configure>", _on_canvas_resize)
        body.bind("<Configure>", lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")))

        # Poll until the canvas has real width, then set body width once
        def _init_body_width():
            w = canvas.winfo_width()
            if w > 1:
                canvas.itemconfig(win_id, width=w)
            else:
                canvas.after(30, _init_body_width)
        canvas.after(10, _init_body_width)

        # Mouse-wheel scrolling
        def _scroll(e):
            canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _scroll)

        self._build_body(body, username, month, year, scores or {})

        # ── Bottom nav bar ────────────────────────────────────────────────────
        nav = tk.Frame(self._content, bg=HDR_BG, pady=8)
        nav.pack(fill="x", side="bottom")

        _btn(nav, "← Enter Another Month",
             lambda: self._app.show("InputPage"),
             ACCENT, font_key="b11", px=16, py=6).pack(side="left", padx=16)
        _btn(nav, "New User",
             lambda: self._go_new_user(),
             ACCENT2, font_key="b11", px=16, py=6).pack(side="left")

    def _go_new_user(self):
        pg = self._app._pages["InputPage"]
        pg._user_var.set("")
        pg._reset()
        self._app.show("InputPage")

    def _build_body(self, body, username, month, year, scores):
        # ── Identity strip ────────────────────────────────────────────────────
        id_fr = tk.Frame(body, bg=HDR_BG, pady=10)
        id_fr.pack(padx=24, fill="x", pady=(0, 0))
        tk.Label(id_fr, text=f"{username}   |   {month} {year}",
                 bg=HDR_BG, fg=TEXT, font=F["b14"]).pack()

        # ── Determine quarterly availability up-front ─────────────────────────
        qk, qinfo = get_quarter(month)
        has_quarterly = False
        counted_present = []
        yd = {}
        avg_scores = {}

        if qk:
            yd = self._app.get_user_year_data(username, year)
            counted_present = [m for m in qinfo["counted"] if m in yd]
            if len(counted_present) >= 2:
                has_quarterly = True
                avg_scores = self._compute_avg(yd, counted_present)

        if has_quarterly:
            # ── Full-width quarterly avg table only ───────────────────────────
            self._section(body, qinfo["label"])
            self._avg_table(body, avg_scores,
                            qinfo["counted"], counted_present, yd)
            avg_total = sum(v for v in avg_scores.values())
            self._total_row(body, avg_total, label="Quarterly Avg Total")
        else:
            # ── No quarterly data yet — show this month's score ───────────────
            self._section(body, f"Score — {month} {year}")
            self._score_table(body, scores)
            total = sum(v["score"] for v in scores.values())
            self._total_row(body, total)

        tk.Frame(body, bg=BG, height=20).pack()  # bottom padding

    def _section(self, parent, title, px=24):
        tk.Frame(parent, bg=ACCENT, height=2).pack(fill="x", padx=px, pady=(14, 0))
        tk.Label(parent, text=title, bg=BG, fg=ACCENT,
                 font=F["b12"]).pack(anchor="w", padx=px + 4, pady=(4, 2))

    @staticmethod
    def _row_lbl(parent, text, bg, fg, font, w_chars, anchor="w", wrap=0):
        """Label sized in characters — reliable cross-platform table cell."""
        kw = dict(text=text, bg=bg, fg=fg, font=font, width=w_chars,
                  anchor=anchor, padx=5, pady=7, justify="left")
        if wrap:
            kw["wraplength"] = wrap
        tk.Label(parent, **kw).pack(side="left")

    def _score_table(self, parent, scores, px=24, compact=False):
        # Column widths in characters: KPI, Value, Slab Hit, Score
        CW = [19, 7, 12, 5] if compact else [28, 10, 17, 7]
        wrap = 140 if compact else 190
        tbl = tk.Frame(parent, bg=PANEL)
        tbl.pack(fill="x", padx=px, pady=(2, 0))

        hdr = tk.Frame(tbl, bg=SEP); hdr.pack(fill="x")
        for text, w, anc in zip(["KPI", "Value", "Slab Hit", "Score"],
                                 CW, ["w", "w", "w", "e"]):
            self._row_lbl(hdr, text, SEP, DIM, F["b11"], w, anc)

        for r, (kname, kdata) in enumerate(scores.items()):
            bg = INP_BG if r % 2 == 0 else PANEL
            sc_col = SLAB_COLOR[kdata["pct"]]
            kpi_obj = next(k for k in KPIS if k["name"] == kname)
            val_str = f"{kdata['value']}{kpi_obj['unit']}"
            row = tk.Frame(tbl, bg=bg); row.pack(fill="x")
            self._row_lbl(row, kname,          bg, TEXT,   F["r11"], CW[0], "w", wrap)
            self._row_lbl(row, val_str,         bg, DIM,    F["r11"], CW[1])
            self._row_lbl(row, kdata["label"],  bg, sc_col, F["r11"], CW[2])
            self._row_lbl(row, str(kdata["score"]), bg, sc_col, F["b11"], CW[3], "e")

    def _avg_table(self, parent, avg_scores, all_months, entered_months, yd, px=24):
        # Grid-based table — fills full content width (~812 px).
        # KPI col: weight=3  |  each month + Avg col: weight=2
        n = len(all_months)
        # KPI wraplength ≈ 3/(3+(n+1)*2) × 812px, minus cell padding
        wrap = int(3 / (3 + (n + 1) * 2) * 812) - 14
        mhdr = lambda m: m        # always show full month name

        tbl = tk.Frame(parent, bg=PANEL)
        tbl.pack(fill="x", padx=px, pady=(2, 0))

        tbl.columnconfigure(0, weight=3)          # KPI — largest share
        for ci in range(1, n + 2):
            tbl.columnconfigure(ci, weight=2)     # months + avg — wider share

        def _hc(text, col, fg):
            tk.Label(tbl, text=text, bg=SEP, fg=fg, font=F["b11"],
                     anchor="w" if col == 0 else "e",
                     padx=6, pady=5).grid(row=0, column=col, sticky="ew")

        def _dc(text, row, col, bg, fg, bold=False, wl=0):
            kw = dict(text=text, bg=bg, fg=fg,
                      font=F["b11"] if bold else F["r11"],
                      anchor="w" if col == 0 else "e",
                      padx=6, pady=7)
            if wl:
                kw["wraplength"] = wl
                kw["justify"] = "left"
            tk.Label(tbl, **kw).grid(row=row, column=col, sticky="ew")

        # Header
        _hc("KPI", 0, DIM)
        for i, m in enumerate(all_months):
            _hc(mhdr(m), i + 1, TEXT)
        _hc("Avg", n + 1, GOLD)

        # Data rows
        for r, kpi in enumerate(KPIS):
            kname   = kpi["name"]
            bg      = INP_BG if r % 2 == 0 else PANEL
            avg_sc  = avg_scores[kname]
            avg_col = SLAB_COLOR[_pct_from_score(kpi, avg_sc)]

            _dc(kname, r + 1, 0, bg, TEXT, wl=wrap)
            for i, m in enumerate(all_months):
                if m in entered_months:
                    sc = yd[m][kname]["score"]
                    mc = SLAB_COLOR[yd[m][kname]["pct"]]
                    _dc(str(sc), r + 1, i + 1, bg, mc)
                else:
                    _dc("—", r + 1, i + 1, bg, DIM)
            _dc(f"{avg_sc:.1f}", r + 1, n + 1, bg, avg_col, bold=True)

    def _total_row(self, parent, total, label="Total Score", px=24):
        tot = tk.Frame(parent, bg=HDR_BG, pady=7)
        tot.pack(fill="x", padx=px, pady=(1, 0))
        tot.columnconfigure(1, weight=1)
        pct = total / TOTAL_WT * 100
        score_txt = str(int(round(total))) if total == int(total) else f"{total:.1f}"
        tk.Label(tot, text=label, bg=HDR_BG, fg=TEXT,
                 font=F["b12"]).grid(row=0, column=0, padx=14, sticky="w")
        tk.Label(tot, text=f"{score_txt}  ({pct:.1f}%)  / {TOTAL_WT}",
                 bg=HDR_BG, fg=GOLD, font=F["score"]).grid(
            row=0, column=2, padx=14, sticky="e")

    def _compute_avg(self, yd, months):
        avg = {}
        for kpi in KPIS:
            kname = kpi["name"]
            vals = [yd[m][kname]["score"] for m in months if kname in yd.get(m, {})]
            avg[kname] = sum(vals) / len(vals) if vals else 0.0
        return avg


def _pct_from_score(kpi, avg_score):
    """Map an average score back to a slab colour bucket."""
    full = kpi["weight"]
    if avg_score >= full * 0.9:
        return 1.0
    if avg_score >= full * 0.4:
        return 0.8
    return 0.0


if __name__ == "__main__":
    app = App()
    app.mainloop()
