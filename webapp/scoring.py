"""Shared KPI/scoring rules — ported from sales_score_calculator.py (no Tkinter)."""

ACCENT = "#1b8a3a"   # Slab 1 — green
GOLD   = "#b8860b"   # Slab 2 — amber
RED    = "#c0392b"   # Slab 3 — red
SLAB_COLOR = {1.0: ACCENT, 0.8: GOLD, 0.0: RED}
SLAB_LBL   = {1.0: "Slab 1 — 100%", 0.8: "Slab 2 — 80%", 0.0: "Slab 3 — 0%"}

KPIS = [
    {"name": "Sales Certified Manpower",            "weight": 100, "unit": "%",
     "hint": "≥80% full · 55–79% partial · <55% none",
     "slabs": [(1.0, lambda v: v >= 80,            "≥ 80%"),
               (0.8, lambda v: 55 <= v <= 79,      "55% – 79%"),
               (0.0, lambda v: v < 55,             "< 55%")]},

    {"name": "Central Web-In Lead to Retail Ratio", "weight": 100, "unit": "%",
     "hint": "≥2.25% full · 1.40–2.24% partial · <1.40% none",
     "slabs": [(1.0, lambda v: v >= 2.25,          "≥ 2.25%"),
               (0.8, lambda v: 1.40 <= v < 2.25,   "1.40% – 2.24%"),
               (0.0, lambda v: v < 1.40,           "< 1.40%")]},

    {"name": "Mystery Shopping",                    "weight": 50,  "unit": "%",
     "hint": "≥80% full · 60–79% partial · <60% none",
     "slabs": [(1.0, lambda v: v >= 80,            "≥ 80%"),
               (0.8, lambda v: 60 <= v < 80,       "60% – 79%"),
               (0.0, lambda v: v < 60,             "< 60%")]},

    {"name": "Escalation Index Sales",              "weight": 50,  "unit": "",
     "hint": "≤0.20 full · 0.21–0.94 partial · >0.94 none",
     "slabs": [(1.0, lambda v: v <= 0.20,          "≤ 0.20"),
               (0.8, lambda v: 0.21 <= v <= 0.94,  "0.21 – 0.94"),
               (0.0, lambda v: v > 0.94,           "> 0.94")]},

    {"name": "PSI Sales",                           "weight": 100, "unit": "",
     "hint": "≥4.85 full · 4.60–4.84 partial · <4.60 none",
     "slabs": [(1.0, lambda v: v >= 4.85,          "≥ 4.85"),
               (0.8, lambda v: 4.60 <= v < 4.85,   "4.60 – 4.84"),
               (0.0, lambda v: v < 4.60,           "< 4.60")]},

    {"name": "CX Sales",                            "weight": 100, "unit": "",
     "hint": "≥4.85 full · 4.60–4.84 partial · <4.60 none",
     "slabs": [(1.0, lambda v: v >= 4.85,          "≥ 4.85"),
               (0.8, lambda v: 4.60 <= v < 4.85,   "4.60 – 4.84"),
               (0.0, lambda v: v < 4.60,           "< 4.60")]},
]

TOTAL_WT = sum(k["weight"] for k in KPIS)  # 500

MONTHS = ["January", "February", "March", "April", "May", "June",
          "July", "August", "September", "October", "November", "December"]

# Q2 business rule: April excluded — only May & Jun count toward Q2 average
QUARTERS = {
    "Q1": {"months": ["January", "February", "March"],
           "counted": ["January", "February", "March"],
           "label": "Q1 Average (Jan – Mar)"},
    "Q2": {"months": ["April", "May", "June"],
           "counted": ["May", "June"],
           "label": "Q2 Average (May – Jun)"},
    "Q3": {"months": ["July", "August", "September"],
           "counted": ["July", "August", "September"],
           "label": "Q3 Average (Jul – Sep)"},
    "Q4": {"months": ["October", "November", "December"],
           "counted": ["October", "November", "December"],
           "label": "Q4 Average (Oct – Dec)"},
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
