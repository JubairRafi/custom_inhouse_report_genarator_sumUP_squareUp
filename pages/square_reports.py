"""
pages/square_reports.py
=======================
Square POS Report Generator — Streamlit page.
Supports multiple outlets, date range, category & product filters.
Reports show both Quantity and Net Sales £.
"""
from __future__ import annotations

import datetime
import pathlib
import sys

# Ensure project root is importable (pages/ is one level below root)
sys.path.insert(0, str(pathlib.Path(__file__).parent.parent))

import pandas as pd
import streamlit as st

from square_loader import (
    load_square_csv,
    get_outlets,
    get_categories_sq,
    get_items_sq,
)
from square_report_engine import (
    top10_sq,
    worst5_sq,
    all_items_sq,
    category_totals_sq,
    weekly_by_day_sq,
)
from square_excel_exporter import build_square_excel

# ─────────────────────────────────────────────────────────────────────────────
# Page config
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Square POS Reports",
    page_icon="🟣",
    layout="wide",
    initial_sidebar_state="collapsed",
)

import auth
auth.require_login()

# ─────────────────────────────────────────────────────────────────────────────
# CSS — premium dark theme (purple accent)
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

.stApp { background: linear-gradient(135deg, #0d0a1a 0%, #1a0f2e 50%, #0f1a2e 100%); color: #e0eaf5; }

.settings-card {
    background: rgba(30, 15, 50, 0.92);
    border: 1px solid #4a1472;
    border-radius: 12px;
    padding: 18px 24px 14px 24px;
    margin-bottom: 18px;
    box-shadow: 0 2px 18px rgba(107,33,168,0.2);
}
.settings-card label, .settings-card p { color: #d8b4fe !important; }

div[data-testid="stButton"] > button {
    border-radius: 8px; font-weight: 600; font-size: 0.95rem;
    padding: 10px 26px; transition: all 0.2s ease;
    border: 2px solid transparent; cursor: pointer;
}
div[data-testid="stButton"] > button[kind="primary"] {
    background: linear-gradient(90deg, #6b21a8, #9333ea);
    color: #fff; border-color: #9333ea;
    box-shadow: 0 3px 14px rgba(147,51,234,0.45);
}
div[data-testid="stButton"] > button[kind="secondary"] {
    background: rgba(40, 15, 65, 0.8); color: #a78bfa; border-color: #4a1472;
}
div[data-testid="stButton"] > button[kind="secondary"]:hover {
    background: rgba(107,33,168,0.2); border-color: #9333ea; color: #e9d5ff;
}

div[data-testid="stDownloadButton"] > button {
    background: linear-gradient(90deg, #1a7a48, #22a05a) !important;
    color: #fff !important; font-weight: 700 !important;
    font-size: 1rem !important; border-radius: 8px !important;
    padding: 12px 28px !important; border: none !important;
    box-shadow: 0 3px 14px rgba(34,160,90,0.38) !important;
    transition: all 0.18s ease !important; width: 100% !important;
}
div[data-testid="stDownloadButton"] > button:hover {
    background: linear-gradient(90deg, #166040, #1a8a4a) !important;
    box-shadow: 0 5px 20px rgba(34,160,90,0.5) !important;
    transform: translateY(-1px) !important;
}

.stTabs [data-baseweb="tab-list"] {
    background: rgba(10, 5, 20, 0.7); border-radius: 8px; padding: 4px; gap: 4px;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 6px; color: #a78bfa; font-weight: 600; padding: 8px 20px;
}
.stTabs [aria-selected="true"] { background: #6b21a8 !important; color: #fff !important; }

[data-testid="stMetric"] {
    background: rgba(30, 15, 50, 0.8); border: 1px solid #4a1472;
    border-radius: 8px; padding: 12px 16px;
}
hr { border-color: #2d1050; }

.section-label {
    font-size: 0.78rem; font-weight: 600; letter-spacing: 0.08em;
    color: #7c3aed; text-transform: uppercase; margin-bottom: 6px;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# Session state
# ─────────────────────────────────────────────────────────────────────────────
_DEFAULTS: dict = {
    "sq_df":            None,
    "sq_active_report": None,
    "sq_reports":       {},
    "sq_excel_bytes":   None,
    "_sq_last_key":     None,
}
for _k, _v in _DEFAULTS.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v

# ─────────────────────────────────────────────────────────────────────────────
# Header
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<div style="
    background: linear-gradient(90deg, #4a1472, #7b2cbf);
    border-radius: 12px; padding: 20px 28px 16px;
    margin-bottom: 18px;
    box-shadow: 0 4px 24px rgba(107,33,168,0.45);">
  <h1 style="color:#fff; margin:0; font-size:1.75rem; font-weight:700;">
    🟣 Square POS Report Generator
  </h1>
  <p style="color:#d8b4fe; margin:4px 0 0; font-size:0.9rem;">
    Upload your Square CSV export · filter by outlet, date & category · generate reports with Qty <em>and</em> Net Sales.
  </p>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# SETTINGS CARD
# ─────────────────────────────────────────────────────────────────────────────
with st.container():
    st.markdown('<div class="settings-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-label">⚙️ Settings</div>', unsafe_allow_html=True)

    # ── File uploader ─────────────────────────────────────────────────────────
    sq_file = st.file_uploader(
        "📁 Square POS Export (.csv)",
        type=["csv"],
        key="sq_upload",
    )
    if sq_file is not None:
        try:
            st.session_state.sq_df = load_square_csv(sq_file)
        except Exception as exc:
            st.error(f"Error loading file: {exc}")
            st.session_state.sq_df = None

    sq_df = st.session_state.sq_df
    all_outlets = get_outlets(sq_df) if sq_df is not None else []

    # ── Row 1 — dates + outlet filter ────────────────────────────────────────
    rc1, rc2, rc3 = st.columns([1, 1, 2])
    today = datetime.date.today()
    with rc1:
        start_date = st.date_input("📅 From", value=today.replace(day=1), key="sq_start")
    with rc2:
        end_date = st.date_input("📅 To", value=today, key="sq_end")
    with rc3:
        if all_outlets:
            selected_outlets = st.multiselect(
                "🏪 Outlets (leave empty = all)",
                options=all_outlets,
                default=[],
                key="sq_outlets",
                placeholder="All outlets…",
            )
        else:
            selected_outlets = []
            if sq_df is None:
                st.info("Upload a file above to see available outlets.", icon="ℹ️")

    # ── Row 2 — category + exclude products ──────────────────────────────────
    eff_outlets = selected_outlets or None
    all_cats  = get_categories_sq(sq_df, eff_outlets) if sq_df is not None else []
    rc4, rc5, rc6 = st.columns([1, 1, 1])
    with rc4:
        if all_cats:
            selected_cats = st.multiselect(
                "🏷️ Category (leave empty = all)",
                options=all_cats,
                default=[],
                key="sq_cats",
                placeholder="All categories…",
            )
        else:
            selected_cats = []
    with rc5:
        all_items = (
            get_items_sq(sq_df, selected_cats or None, eff_outlets)
            if sq_df is not None else []
        )
        if all_items:
            excluded_items = st.multiselect(
                "🚫 Exclude Products",
                options=all_items,
                default=[],
                key="sq_excl",
                placeholder="Select products to exclude…",
            )
        else:
            excluded_items = []
    with rc6:
        selected_cols = st.multiselect(
            "👁️ Visible Columns",
            options=["#SL", "Category", "Qty", "Net Sales £"],
            default=["#SL", "Category", "Qty", "Net Sales £"],
            key="sq_vis_cols",
            help="Uncheck to hide these metrics in both the on-screen preview AND the downloaded Excel file."
        )

    st.markdown("</div>", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# Status badges
# ─────────────────────────────────────────────────────────────────────────────
file_ready = sq_df is not None
b1, b2, b3, b4 = st.columns(4)

with b1:
    if file_ready:
        st.success(f"✅ {len(sq_df):,} transaction rows loaded")
    else:
        st.warning("⚠️ No file loaded")

with b2:
    if file_ready:
        shown = selected_outlets if selected_outlets else all_outlets
        st.info(f"🏪 {len(shown)} outlet(s) active")

with b3:
    if selected_cats:
        st.info(f"🏷️ {len(selected_cats)} category filter(s)")
    elif all_cats:
        st.info("🏷️ All categories included")

with b4:
    if excluded_items:
        st.warning(f"🚫 {len(excluded_items)} product(s) excluded")
    elif file_ready:
        st.info("📦 No exclusions")

st.markdown("---")

# ─────────────────────────────────────────────────────────────────────────────
# REPORT BUTTONS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-label">📋 Select Report</div>', unsafe_allow_html=True)
active = st.session_state.sq_active_report

bc1, bc2, bc3, bc4, bc5 = st.columns(5)
_REPORTS = {
    "top10":           ("🏆  Top 10", "sq_btn_top10"),
    "worst5":          ("📉  Worst 5", "sq_btn_worst5"),
    "all_items":       ("📋  All Items", "sq_btn_all"),
    "category_totals": ("📊  Category Totals", "sq_btn_cat"),
    "weekly_by_day":   ("📅  Weekly by Day", "sq_btn_wk"),
}
for col, (rtype, (label, key)) in zip([bc1, bc2, bc3, bc4, bc5], _REPORTS.items()):
    with col:
        btype = "primary" if active == rtype else "secondary"
        if st.button(label, type=btype, use_container_width=True,
                     key=key, disabled=not file_ready):
            st.session_state.sq_active_report = rtype
            st.rerun()

if not file_ready:
    st.caption("Upload a Square CSV file above to enable report generation.")

# ─────────────────────────────────────────────────────────────────────────────
# REPORT GENERATION
# ─────────────────────────────────────────────────────────────────────────────
_FN_MAP = {
    "top10":           top10_sq,
    "worst5":          worst5_sq,
    "all_items":       all_items_sq,
    "category_totals": category_totals_sq,
    "weekly_by_day":   weekly_by_day_sq,
}


def _generate(report_type: str) -> None:
    if start_date > end_date:
        st.error("'From' date must be on or before the 'To' date.")
        return

    cats  = selected_cats  or None
    excl  = excluded_items or None
    fn    = _FN_MAP[report_type]

    # Determine which outlet tabs to build
    act_outlets = selected_outlets if selected_outlets else all_outlets

    def _apply_visibility(df: pd.DataFrame) -> pd.DataFrame:
        df_out = df.copy()
        if "#SL" not in selected_cols and "#SL" in df_out.columns:
            df_out = df_out.drop(columns=["#SL"])
        if "Category" not in selected_cols and "Category" in df_out.columns:
            df_out = df_out.drop(columns=["Category"])
        if "Qty" not in selected_cols:
            cols_to_drop = [c for c in df_out.columns if str(c).endswith("Qty")]
            df_out = df_out.drop(columns=cols_to_drop)
        if "Net Sales £" not in selected_cols:
            cols_to_drop = [c for c in df_out.columns if "Net" in str(c) or "£" in str(c)]
            df_out = df_out.drop(columns=cols_to_drop)
        return df_out

    with st.spinner("Generating report…"):
        try:
            reports: dict[str, tuple] = {}

            # Combined (all selected outlets merged)
            if report_type != "weekly_by_day":
                df_c, title_c = fn(
                    sq_df, "Combined", start_date, end_date,
                    outlets=selected_outlets or None,
                    categories=cats, exclude_items=excl,
                )
                df_c = _apply_visibility(df_c)
                reports["📊 Combined"] = (df_c, title_c)

            # One tab per outlet
            outlet_icons = ["🟣", "🔵", "🟠", "🟢", "🔴", "🟡"]
            for i, outlet in enumerate(act_outlets):
                icon = outlet_icons[i % len(outlet_icons)]
                df_o, title_o = fn(
                    sq_df, outlet, start_date, end_date,
                    outlets=[outlet],
                    categories=cats, exclude_items=excl,
                )
                df_o = _apply_visibility(df_o)
                reports[f"{icon} {outlet}"] = (df_o, title_o)

        except Exception as exc:
            st.error(f"Report error: {exc}")
            return

        # All results empty?
        if all(len(df) <= 1 for df, _ in reports.values()):
            st.warning("No data found for the selected filters and date range.")
            return

        try:
            excel_bytes = build_square_excel(reports)
        except Exception as exc:
            st.error(f"Excel export error: {exc}")
            return

        st.session_state.sq_reports     = reports
        st.session_state.sq_excel_bytes = excel_bytes


# Trigger only when inputs change
if st.session_state.sq_active_report and file_ready:
    run_key = (
        st.session_state.sq_active_report,
        str(start_date), str(end_date),
        tuple(sorted(selected_outlets)),
        tuple(sorted(selected_cats)),
        tuple(sorted(excluded_items)),
        tuple(sorted(selected_cols)),
        id(sq_df),
    )
    if run_key != st.session_state["_sq_last_key"]:
        _generate(st.session_state.sq_active_report)
        st.session_state["_sq_last_key"] = run_key

# ─────────────────────────────────────────────────────────────────────────────
# REPORT OUTPUT
# ─────────────────────────────────────────────────────────────────────────────
if st.session_state.sq_excel_bytes and st.session_state.sq_reports:
    reports    = st.session_state.sq_reports
    first_title = next(iter(reports.values()))[1]
    safe_name  = (
        first_title
        .replace(" – ", "_").replace(" → ", "_to_")
        .replace(" ", "_").replace("/", "-")
    )
    fname = f"Square_{safe_name}.xlsx"

    st.markdown("---")
    st.success(f"Report ready: **{first_title}**")

    # ── Preview helper ────────────────────────────────────────────────────────
    def _preview(df: pd.DataFrame, title: str) -> None:
        st.caption(title)
        n         = len(df) - 1
        data_rows = df.iloc[:n].copy()
        total_row = df.iloc[[n]].copy()

        if "#SL" in data_rows.columns:
            data_rows["#SL"] = pd.to_numeric(
                data_rows["#SL"], errors="coerce"
            ).astype("Int64")

        if "Net Sales £" in data_rows.columns:
            data_rows["Net Sales £"] = pd.to_numeric(
                data_rows["Net Sales £"], errors="coerce"
            ).map("£{:,.2f}".format)
            total_row["Net Sales £"] = pd.to_numeric(
                total_row["Net Sales £"], errors="coerce"
            ).map("£{:,.2f}".format)

        st.dataframe(data_rows, use_container_width=True, hide_index=True)
        st.dataframe(total_row, use_container_width=True, hide_index=True)

    # ── Dynamic tabs (Combined + one per outlet) ──────────────────────────────
    tab_keys   = list(reports.keys())
    tab_labels = []
    for k in tab_keys:
        label = k if len(k) <= 28 else k[:25] + "…"
        tab_labels.append(label)

    tabs = st.tabs(tab_labels)
    for tab, key in zip(tabs, tab_keys):
        with tab:
            df, title = reports[key]
            _preview(df, title)

    # ── Download ──────────────────────────────────────────────────────────────
    st.markdown("---")
    dl_col, _ = st.columns([2, 3])
    with dl_col:
        st.download_button(
            label="⬇️  Download Excel Report (.xlsx)",
            data=st.session_state.sq_excel_bytes,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="sq_dl_btn",
        )

elif not file_ready:
    st.markdown("""
<div style="
    background: rgba(74,20,114,0.25);
    border: 1px solid #6b21a8;
    border-radius: 12px;
    padding: 28px 32px;
    margin-top: 8px;">
  <h3 style="color:#e9d5ff; margin-top:0;">🚀 Getting started</h3>
  <ol style="color:#c4b5fd; line-height:2.2;">
    <li>Upload your <strong style="color:#e9d5ff;">Square POS CSV export</strong> above.</li>
    <li>Select the <strong style="color:#e9d5ff;">outlets</strong> you want (or leave empty for all).</li>
    <li>Set your <strong style="color:#e9d5ff;">date range</strong> and optional <strong style="color:#e9d5ff;">category</strong> filter.</li>
    <li>Click a <strong style="color:#e9d5ff;">report button</strong> to generate.</li>
    <li>Preview per-outlet tabs, then <strong style="color:#e9d5ff;">download</strong> the Excel report.</li>
  </ol>
  <p style="color:#7c3aed; margin:0; font-size:0.85rem;">
    Reports include both <strong>Quantity</strong> and <strong>Net Sales £</strong>.
    Each outlet gets its own tab in the preview and its own sheet in the Excel file.
  </p>
</div>
""", unsafe_allow_html=True)
