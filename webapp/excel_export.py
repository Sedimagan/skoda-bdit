"""Build the same formatted Excel workbook the desktop app produces, in-memory."""

import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from scoring import KPIS, MONTHS

MONTH_ORDER = {m: i for i, m in enumerate(MONTHS)}


def build_workbook(db) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BDIT Scores"

    hdr_fill = PatternFill("solid", fgColor="0F3460")
    hdr_font = Font(bold=True, color="FFFFFF", size=10)
    hdr_aln = Alignment(horizontal="center", vertical="center", wrap_text=True)
    val_aln = Alignment(horizontal="center", vertical="center")
    thin = Side(style="thin", color="2A3A5C")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    slab_fill = {
        1.0: PatternFill("solid", fgColor="1B5E20"),
        0.8: PatternFill("solid", fgColor="7B5E00"),
        0.0: PatternFill("solid", fgColor="7B1A2C"),
    }

    headers = ["User Name", "Year", "Month"]
    for kpi in KPIS:
        n = kpi["name"]
        headers += [f"{n}\nValue", f"{n}\nSlab", f"{n}\nScore"]
    headers.append("Total\nScore")
    ws.append(headers)

    for cell in ws[1]:
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.alignment = hdr_aln
        cell.border = border
    ws.row_dimensions[1].height = 42

    for user in sorted(db):
        for year in sorted(db[user]):
            month_data = db[user][year]
            for month in sorted(month_data, key=lambda m: MONTH_ORDER.get(m, 99)):
                kpis = month_data[month]
                row = [user, int(year), month]
                total = 0
                pcts = []
                for kpi in KPIS:
                    d = kpis.get(kpi["name"], {})
                    row += [d.get("value", ""), d.get("label", ""), d.get("score", "")]
                    total += d.get("score", 0) or 0
                    pcts.append(d.get("pct", None))
                row.append(total)
                ws.append(row)

                r = ws.max_row
                for col_i, pct in enumerate(pcts):
                    if pct is not None and pct in slab_fill:
                        score_col = 4 + col_i * 3
                        ws.cell(r, score_col).fill = slab_fill[pct]
                for cell in ws[r]:
                    cell.alignment = val_aln
                    cell.border = border

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 6
    ws.column_dimensions["C"].width = 11
    col_letters = [chr(c) for c in range(ord("D"), ord("D") + len(KPIS) * 3 + 1)]
    for i, ltr in enumerate(col_letters):
        ws.column_dimensions[ltr].width = 8 if (i % 3 == 2) else 10
    last = openpyxl.utils.get_column_letter(4 + len(KPIS) * 3)
    ws.column_dimensions[last].width = 8
    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
