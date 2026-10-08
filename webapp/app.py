"""Gurudev Motors — Best Dealer In Town Sales Score Calculator (Gradio web edition)."""

import base64
import datetime
import os

import gradio as gr
import pandas as pd

try:
    import spaces  # Hugging Face's ZeroGPU helper — a no-op outside that runtime
except ImportError:
    spaces = None

from scoring import KPIS, TOTAL_WT, MONTHS, QUARTERS, SLAB_COLOR, compute_slab
from storage import load_db, save_entry
from excel_export import build_workbook

if spaces is not None:
    @spaces.GPU
    def _zerogpu_startup_check():
        """This app never uses a GPU, but Spaces on ZeroGPU hardware require at
        least one @spaces.GPU-decorated function to exist or the Space fails
        to start. Never called."""
        return True

ASSET_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
LOGO_PATH = os.path.join(ASSET_DIR, "skoda_logo.png")

with open(LOGO_PATH, "rb") as f:
    _LOGO_DATA_URI = "data:image/png;base64," + base64.b64encode(f.read()).decode()

MONTH_ORDER = {m: i for i, m in enumerate(MONTHS)}
TODAY = datetime.date.today()

CUSTOM_CSS = """
.app-header {
    display: flex; align-items: center; gap: 16px;
    padding: 14px 22px; margin-bottom: 1rem;
    background: linear-gradient(135deg, #0f3460, #16213e);
    border-bottom: 3px solid #4caf50; border-radius: 8px;
}
.app-header img { height: 48px; width: 48px; }
.app-header .titles { display: flex; flex-direction: column; justify-content: center; }
.app-header .main-title { font-size: 1.4rem; font-weight: 700; color: #f2f2f2; line-height: 1.2; }
.app-header .sub-title { font-size: 0.9rem; font-weight: 600; color: #4caf50; line-height: 1.3; }

.kpi-name { font-size: 1rem; font-weight: 600; margin-bottom: 2px; }
.kpi-hint { font-size: 0.78rem; color: #b8c2d6; margin-bottom: 6px; }
.slab-badge {
    display: inline-block; padding: 6px 12px; border-radius: 6px;
    font-weight: 600; font-size: 0.9rem;
}
.total-score { text-align: right; }
"""

HEADER_HTML = f"""
<div class="app-header">
    <img src="{_LOGO_DATA_URI}" />
    <div class="titles">
        <span class="main-title">Gurudev Motors Pvt. Ltd</span>
        <span class="sub-title">Best Dealer in Town — Sales Score Calculator</span>
    </div>
</div>
"""


def _badge(kpi, val):
    if val is None:
        return "—"
    pct, lbl, score = compute_slab(kpi, val)
    color = SLAB_COLOR[pct]
    return (f"<span class='slab-badge' style='background:{color}20;color:{color};'>"
            f"{lbl} · {score} pts</span>")


def _recompute(*vals):
    badges = [_badge(kpi, v) for kpi, v in zip(KPIS, vals)]
    if any(v is None for v in vals):
        total_md = "### Total Score: —"
    else:
        total = sum(compute_slab(kpi, v)[2] for kpi, v in zip(KPIS, vals))
        total_md = f"### Total Score: {total} / {TOTAL_WT}  ({total / TOTAL_WT * 100:.1f}%)"
    return badges + [total_md]


def _on_submit(username, month, year, *vals):
    if not username or not username.strip():
        return "❌ Please enter a user name."
    if any(v is None for v in vals):
        return "❌ Please enter a value for every KPI."
    scores = {}
    for kpi, v in zip(KPIS, vals):
        pct, lbl, score = compute_slab(kpi, v)
        scores[kpi["name"]] = {"value": v, "pct": pct, "label": lbl, "score": score}
    save_entry(username.strip(), int(year), month, scores)
    total = sum(s["score"] for s in scores.values())
    return f"✅ Saved {username.strip()} — {month} {int(year)}: {total} / {TOTAL_WT}"


def _refresh_hist_users():
    db = load_db()
    users = sorted(db.keys())
    return gr.update(choices=users, value=users[0] if users else None)


def _refresh_hist_years(user):
    db = load_db()
    years = sorted(db.get(user, {}).keys()) if user else []
    return gr.update(choices=years, value=years[-1] if years else None)


def _show_history(user, year):
    db = load_db()
    if not user or not year or user not in db or year not in db.get(user, {}):
        return pd.DataFrame(), "No entries for this year yet."

    yd = db[user][year]
    entered_months = sorted(yd.keys(), key=lambda m: MONTH_ORDER.get(m, 99))

    rows = []
    for m in entered_months:
        row = {"Month": m}
        tot = 0
        for kpi in KPIS:
            d = yd[m].get(kpi["name"], {})
            row[kpi["name"]] = d.get("score", "")
            tot += d.get("score", 0) or 0
        row["Total"] = tot
        rows.append(row)
    df = pd.DataFrame(rows)

    quarter_lines = []
    for qk, qinfo in QUARTERS.items():
        counted_present = [m for m in qinfo["counted"] if m in yd]
        if len(counted_present) >= 2:
            avg = {}
            for kpi in KPIS:
                vals = [yd[m][kpi["name"]]["score"] for m in counted_present
                         if kpi["name"] in yd.get(m, {})]
                avg[kpi["name"]] = sum(vals) / len(vals) if vals else 0.0
            avg_line = " · ".join(f"{k}: {v:.1f}" for k, v in avg.items())
            quarter_lines.append(
                f"**{qinfo['label']}** — {avg_line}  \n"
                f"Quarterly Avg Total: **{sum(avg.values()):.1f} / {TOTAL_WT}**"
            )
    return df, ("\n\n".join(quarter_lines) if quarter_lines else "")


def _refresh_team_years():
    db = load_db()
    years = sorted({y for u in db.values() for y in u.keys()})
    return gr.update(choices=years, value=years[-1] if years else None)


def _show_leaderboard(month, year):
    db = load_db()
    leaderboard = []
    for user, years in db.items():
        yd = years.get(str(year), {})
        md = yd.get(month)
        if md:
            tot = sum(v.get("score", 0) or 0 for v in md.values())
            leaderboard.append({"User": user, "Total Score": tot,
                                 "%": round(tot / TOTAL_WT * 100, 1)})
    if not leaderboard:
        return pd.DataFrame(columns=["User", "Total Score", "%"])
    return pd.DataFrame(leaderboard).sort_values("Total Score", ascending=False)


def _export_excel():
    db = load_db()
    xlsx_bytes = build_workbook(db)
    path = os.path.join("/tmp", "skoda_scores.xlsx")
    with open(path, "wb") as f:
        f.write(xlsx_bytes)
    return path


# Dark palette matching the original desktop app. Every token is set for
# BOTH its light and _dark variant to the same value, so the UI stays dark
# regardless of the visitor's browser/OS color-scheme preference.
_BG = "#1a1a2e"      # page background
_PANEL = "#16213e"   # block/card background
_INPUT = "#0d2137"   # input fields, table stripes
_BORDER = "#2a3a5c"
_TEXT = "#f2f2f2"    # primary text — high contrast on dark
_DIM = "#b8c2d6"     # secondary text — still readable, not washed out
_ACCENT = "#4caf50"
_ACCENT_HOVER = "#5fc563"
_RED = "#e94560"

theme = gr.themes.Default(primary_hue="green", neutral_hue="slate").set(
    body_background_fill=_BG, body_background_fill_dark=_BG,
    body_text_color=_TEXT, body_text_color_dark=_TEXT,
    body_text_color_subdued=_DIM, body_text_color_subdued_dark=_DIM,

    background_fill_primary=_PANEL, background_fill_primary_dark=_PANEL,
    background_fill_secondary=_INPUT, background_fill_secondary_dark=_INPUT,

    border_color_primary=_BORDER, border_color_primary_dark=_BORDER,
    border_color_accent=_ACCENT, border_color_accent_dark=_ACCENT,
    border_color_accent_subdued=_ACCENT, border_color_accent_subdued_dark=_ACCENT,

    link_text_color=_ACCENT, link_text_color_dark=_ACCENT,
    link_text_color_hover=_ACCENT_HOVER, link_text_color_hover_dark=_ACCENT_HOVER,
    link_text_color_active=_ACCENT, link_text_color_active_dark=_ACCENT,
    link_text_color_visited=_ACCENT, link_text_color_visited_dark=_ACCENT,

    code_background_fill=_INPUT, code_background_fill_dark=_INPUT,

    block_background_fill=_PANEL, block_background_fill_dark=_PANEL,
    block_border_color=_BORDER, block_border_color_dark=_BORDER,
    block_info_text_color=_DIM, block_info_text_color_dark=_DIM,
    block_label_background_fill=_PANEL, block_label_background_fill_dark=_PANEL,
    block_label_border_color=_BORDER, block_label_border_color_dark=_BORDER,
    block_label_text_color=_TEXT, block_label_text_color_dark=_TEXT,
    block_title_background_fill=_PANEL, block_title_background_fill_dark=_PANEL,
    block_title_border_color=_BORDER, block_title_border_color_dark=_BORDER,
    block_title_text_color=_TEXT, block_title_text_color_dark=_TEXT,

    panel_background_fill=_PANEL, panel_background_fill_dark=_PANEL,
    panel_border_color=_BORDER, panel_border_color_dark=_BORDER,

    accordion_text_color=_TEXT, accordion_text_color_dark=_TEXT,

    table_text_color=_TEXT, table_text_color_dark=_TEXT,
    table_border_color=_BORDER, table_border_color_dark=_BORDER,
    table_even_background_fill=_PANEL, table_even_background_fill_dark=_PANEL,
    table_odd_background_fill=_INPUT, table_odd_background_fill_dark=_INPUT,

    checkbox_background_color=_INPUT, checkbox_background_color_dark=_INPUT,
    checkbox_background_color_selected=_ACCENT, checkbox_background_color_selected_dark=_ACCENT,
    checkbox_border_color=_BORDER, checkbox_border_color_dark=_BORDER,
    checkbox_border_color_selected=_ACCENT, checkbox_border_color_selected_dark=_ACCENT,
    checkbox_label_background_fill=_PANEL, checkbox_label_background_fill_dark=_PANEL,
    checkbox_label_background_fill_selected=_ACCENT, checkbox_label_background_fill_selected_dark=_ACCENT,
    checkbox_label_border_color=_BORDER, checkbox_label_border_color_dark=_BORDER,
    checkbox_label_text_color=_TEXT, checkbox_label_text_color_dark=_TEXT,
    checkbox_label_text_color_selected="#ffffff", checkbox_label_text_color_selected_dark="#ffffff",

    error_background_fill="#3a1620", error_background_fill_dark="#3a1620",
    error_border_color=_RED, error_border_color_dark=_RED,
    error_text_color="#ff93a6", error_text_color_dark="#ff93a6",

    input_background_fill=_INPUT, input_background_fill_dark=_INPUT,
    input_background_fill_focus=_INPUT, input_background_fill_focus_dark=_INPUT,
    input_background_fill_hover=_INPUT, input_background_fill_hover_dark=_INPUT,
    input_border_color=_BORDER, input_border_color_dark=_BORDER,
    input_border_color_focus=_ACCENT, input_border_color_focus_dark=_ACCENT,
    input_border_color_hover=_ACCENT, input_border_color_hover_dark=_ACCENT,

    stat_background_fill=_INPUT, stat_background_fill_dark=_INPUT,

    button_primary_background_fill=_ACCENT, button_primary_background_fill_dark=_ACCENT,
    button_primary_background_fill_hover=_ACCENT_HOVER, button_primary_background_fill_hover_dark=_ACCENT_HOVER,
    button_primary_border_color=_ACCENT, button_primary_border_color_dark=_ACCENT,
    button_primary_text_color="#ffffff", button_primary_text_color_dark="#ffffff",
    button_primary_text_color_hover="#ffffff", button_primary_text_color_hover_dark="#ffffff",

    button_secondary_background_fill=_INPUT, button_secondary_background_fill_dark=_INPUT,
    button_secondary_background_fill_hover=_PANEL, button_secondary_background_fill_hover_dark=_PANEL,
    button_secondary_border_color=_BORDER, button_secondary_border_color_dark=_BORDER,
    button_secondary_text_color=_TEXT, button_secondary_text_color_dark=_TEXT,
    button_secondary_text_color_hover=_TEXT, button_secondary_text_color_hover_dark=_TEXT,

    button_cancel_background_fill=_RED, button_cancel_background_fill_dark=_RED,
    button_cancel_background_fill_hover="#ff6b85", button_cancel_background_fill_hover_dark="#ff6b85",
    button_cancel_border_color=_RED, button_cancel_border_color_dark=_RED,
    button_cancel_text_color="#ffffff", button_cancel_text_color_dark="#ffffff",
)

with gr.Blocks(title="BDIT Sales Score") as demo:
    gr.HTML(HEADER_HTML)

    with gr.Tabs():
        # ── Enter Score ───────────────────────────────────────────────────
        with gr.Tab("✍️ Enter Score"):
            with gr.Row():
                username_in = gr.Textbox(label="User Name")
                month_in = gr.Dropdown(MONTHS, value=MONTHS[TODAY.month - 1], label="Month")
                year_in = gr.Number(label="Year", value=TODAY.year, precision=0)

            kpi_components = {}  # kpi name -> (number input, badge HTML)
            with gr.Row():
                for col_idx in range(2):
                    with gr.Column():
                        for kpi in KPIS[col_idx::2]:
                            with gr.Group():
                                gr.Markdown(f"<div class='kpi-name'>{kpi['name']}</div>"
                                            f"<div class='kpi-hint'>{kpi['hint']}</div>")
                                with gr.Row():
                                    num = gr.Number(show_label=False, value=None, scale=1)
                                    badge = gr.HTML("—", scale=1)
                            kpi_components[kpi["name"]] = (num, badge)

            kpi_inputs = [kpi_components[kpi["name"]][0] for kpi in KPIS]
            kpi_badges = [kpi_components[kpi["name"]][1] for kpi in KPIS]

            total_display = gr.Markdown("### Total Score: —", elem_classes="total-score")
            submit_btn = gr.Button("Submit", variant="primary")
            submit_status = gr.Markdown()

            for inp in kpi_inputs:
                inp.change(_recompute, inputs=kpi_inputs, outputs=kpi_badges + [total_display])

            submit_btn.click(
                _on_submit,
                inputs=[username_in, month_in, year_in] + kpi_inputs,
                outputs=submit_status,
            )

        # ── My History ───────────────────────────────────────────────────
        with gr.Tab("📈 My History"):
            hist_refresh_btn = gr.Button("🔄 Refresh")
            with gr.Row():
                hist_user = gr.Dropdown(choices=[], label="User Name")
                hist_year = gr.Dropdown(choices=[], label="Year")
            hist_table = gr.Dataframe(label="Monthly Scores")
            hist_quarter_md = gr.Markdown()

            hist_refresh_btn.click(_refresh_hist_users, outputs=hist_user).then(
                _refresh_hist_years, inputs=hist_user, outputs=hist_year
            ).then(_show_history, inputs=[hist_user, hist_year],
                   outputs=[hist_table, hist_quarter_md])
            hist_user.change(_refresh_hist_years, inputs=hist_user, outputs=hist_year).then(
                _show_history, inputs=[hist_user, hist_year], outputs=[hist_table, hist_quarter_md]
            )
            hist_year.change(_show_history, inputs=[hist_user, hist_year],
                              outputs=[hist_table, hist_quarter_md])

        # ── Team Overview ────────────────────────────────────────────────
        with gr.Tab("🏆 Team Overview"):
            with gr.Row():
                team_month = gr.Dropdown(MONTHS, value=MONTHS[TODAY.month - 1], label="Month")
                team_year = gr.Dropdown(choices=[], label="Year")
            team_refresh_btn = gr.Button("🔄 Refresh")
            leaderboard_table = gr.Dataframe(label="Leaderboard")
            download_btn = gr.DownloadButton("⬇️ Download full Excel export")

            team_refresh_btn.click(_refresh_team_years, outputs=team_year).then(
                _show_leaderboard, inputs=[team_month, team_year], outputs=leaderboard_table
            )
            team_month.change(_show_leaderboard, inputs=[team_month, team_year],
                               outputs=leaderboard_table)
            team_year.change(_show_leaderboard, inputs=[team_month, team_year],
                              outputs=leaderboard_table)
            download_btn.click(_export_excel, outputs=download_btn)

    demo.load(_refresh_hist_users, outputs=hist_user).then(
        _refresh_hist_years, inputs=hist_user, outputs=hist_year
    ).then(_show_history, inputs=[hist_user, hist_year], outputs=[hist_table, hist_quarter_md])

    demo.load(_refresh_team_years, outputs=team_year).then(
        _show_leaderboard, inputs=[team_month, team_year], outputs=leaderboard_table
    )

if __name__ == "__main__":
    demo.launch(theme=theme, css=CUSTOM_CSS, favicon_path=LOGO_PATH)
