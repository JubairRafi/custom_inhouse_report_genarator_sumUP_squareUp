"""
app.py
======
St George Sales Report Manager — Streamlit UI (v2)

Layout
------
  [Header]
  [Settings bar: Kings file | St Martin file | Date From→To | Category filter]
  [Report buttons: 🏆 Top 10  |  📉 Worst 5  |  (future…)]
  [Report output: tabs (Combined / Kings Road / St Martin) + Download]
"""

from __future__ import annotations

import datetime
import io

import pandas as pd
import streamlit as st

from data_loader import load_outlet_file, get_categories, get_items_for_categories
from excel_exporter import build_excel
from report_engine import (
    top10_combined, top10_single,
    worst5_combined, worst5_single,
    all_items_combined, all_items_single,
    category_totals_combined, category_totals_single,
)

# ─────────────────────────────────────────────────────────────────────────────
# Page config
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="St George Report Manager",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────────────────────────────────────
# CSS — premium dark theme
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

/* Background */
.stApp { background: linear-gradient(135deg, #0d1b2a 0%, #1a2f4a 100%); color: #e0eaf5; }

/* ── Settings card ── */
.settings-card {
    background: rgba(20, 38, 60, 0.92);
    border: 1px solid #2a4568;
    border-radius: 12px;
    padding: 18px 24px 14px 24px;
    margin-bottom: 18px;
    box-shadow: 0 2px 14px rgba(0,0,0,0.35);
}
.settings-card label, .settings-card p { color: #b8d0e8 !important; }

/* ── Report button bar ── */
.report-btn-bar {
    display: flex;
    gap: 12px;
    margin-bottom: 20px;
    flex-wrap: wrap;
}

/* ── Active/inactive report selector buttons ── */
div[data-testid="stButton"] > button {
    border-radius: 8px;
    font-weight: 600;
    font-size: 0.95rem;
    padding: 10px 26px;
    transition: all 0.18s ease;
    border: 2px solid transparent;
    cursor: pointer;
}
div[data-testid="stButton"] > button[kind="primary"] {
    background: linear-gradient(90deg, #1f6fbf, #2e90e0);
    color: #fff;
    border-color: #2e90e0;
    box-shadow: 0 3px 12px rgba(46,144,224,0.38);
}
div[data-testid="stButton"] > button[kind="secondary"] {
    background: rgba(28, 48, 72, 0.8);
    color: #90b8d8;
    border-color: #2a4568;
}
div[data-testid="stButton"] > button[kind="secondary"]:hover {
    background: rgba(46, 144, 224, 0.15);
    border-color: #2e90e0;
    color: #c8e0f5;
}

/* ── Download button ── */
div[data-testid="stDownloadButton"] > button {
    background: linear-gradient(90deg, #1a7a48, #22a05a) !important;
    color: #fff !important;
    font-weight: 700 !important;
    font-size: 1rem !important;
    border-radius: 8px !important;
    padding: 12px 28px !important;
    border: none !important;
    box-shadow: 0 3px 14px rgba(34,160,90,0.38) !important;
    transition: all 0.18s ease !important;
    width: 100% !important;
}
div[data-testid="stDownloadButton"] > button:hover {
    background: linear-gradient(90deg, #166040, #1a8a4a) !important;
    box-shadow: 0 5px 18px rgba(34,160,90,0.5) !important;
    transform: translateY(-1px) !important;
}

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"] {
    background: rgba(13, 27, 42, 0.7);
    border-radius: 8px;
    padding: 4px;
    gap: 4px;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 6px;
    color: #7aaece;
    font-weight: 600;
    padding: 8px 22px;
}
.stTabs [aria-selected="true"] {
    background: #1f6fbf !important;
    color: #fff !important;
}

/* ── Metrics ── */
[data-testid="stMetric"] {
    background: rgba(20, 38, 60, 0.8);
    border: 1px solid #2a4568;
    border-radius: 8px;
    padding: 12px 16px;
}
[data-testid="stMetricLabel"] p  { color: #7aaece !important; font-size: 0.8rem; }
[data-testid="stMetricValue"]    { color: #fff !important; font-weight: 700; }

/* ── Alerts ── */
.stAlert { border-radius: 8px; }

/* ── Dataframe ── */
[data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; }

/* ── Divider ── */
hr { border-color: #1e3550; }

/* ── Section label ── */
.section-label {
    font-size: 0.78rem;
    font-weight: 600;
    letter-spacing: 0.08em;
    color: #5a84aa;
    text-transform: uppercase;
    margin-bottom: 6px;
}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# Session state initialisation
# ─────────────────────────────────────────────────────────────────────────────
_SS_DEFAULTS = {
    "kings_df": None,
    "stmartin_df": None,
    "all_categories": [],      # union of categories from both files
    "active_report": None,     # "top10" | "worst5" | None
    "report_combined": None,
    "report_kings": None,
    "report_stmartin": None,
    "title_combined": "",
    "title_kings": "",
    "title_stmartin": "",
    "excel_bytes": None,
}
for k, v in _SS_DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ─────────────────────────────────────────────────────────────────────────────
# Header
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<div style="
    background: linear-gradient(90deg,#12406e,#1c5fa0);
    border-radius:12px; padding:20px 28px 16px;
    margin-bottom:18px;
    box-shadow:0 4px 20px rgba(0,0,0,0.4);">
  <h1 style="color:#fff;margin:0;font-size:1.75rem;font-weight:700;">
    📊 St George Sales Report Manager
  </h1>
  <p style="color:#9ec4e8;margin:4px 0 0;font-size:0.9rem;">
    Load your outlet files, set a date range, filter by category, then pick a report.
  </p>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# SETTINGS SECTION (top card, always visible)
# ─────────────────────────────────────────────────────────────────────────────
with st.container():
    st.markdown('<div class="settings-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-label">⚙️ Settings</div>', unsafe_allow_html=True)

    # Row 1 — file uploaders
    fc1, fc2 = st.columns(2)
    with fc1:
        kings_file = st.file_uploader(
            "📁 Kings Road Sales File",
            type=["xlsx", "xls"],
            key="kings_upload",
            label_visibility="visible",
        )
    with fc2:
        stmartin_file = st.file_uploader(
            "📁 St Martin Sales File",
            type=["xlsx", "xls"],
            key="stmartin_upload",
            label_visibility="visible",
        )

    # Load DataFrames when files change
    if kings_file is not None:
        try:
            st.session_state.kings_df = load_outlet_file(kings_file)
        except ValueError as e:
            st.error(f"Kings Road file error: {e}")
            st.session_state.kings_df = None

    if stmartin_file is not None:
        try:
            st.session_state.stmartin_df = load_outlet_file(stmartin_file)
        except ValueError as e:
            st.error(f"St Martin file error: {e}")
            st.session_state.stmartin_df = None

    # Compute category list
    cats_kings = get_categories(st.session_state.kings_df) if st.session_state.kings_df is not None else []
    cats_stm   = get_categories(st.session_state.stmartin_df) if st.session_state.stmartin_df is not None else []
    all_cats   = sorted(set(cats_kings) | set(cats_stm))

    # Row A — date range + category  (must render before product picker)
    ra1, ra2, ra3 = st.columns([1, 1, 2])
    today = datetime.date.today()
    with ra1:
        start_date = st.date_input("📅 From", value=today.replace(day=1), key="start_date")
    with ra2:
        end_date = st.date_input("📅 To", value=today, key="end_date")
    with ra3:
        if all_cats:
            selected_cats = st.multiselect(
                "🏷️ Category (leave empty = all)",
                options=all_cats,
                default=[],
                key="cat_filter",
                placeholder="All categories…",
            )
        else:
            selected_cats = []
            st.info("No Category column found.", icon="ℹ️")

    # Now that selected_cats is defined, compute scoped item list
    items_kings = get_items_for_categories(st.session_state.kings_df, selected_cats or None) \
        if st.session_state.kings_df is not None else []
    items_stm   = get_items_for_categories(st.session_state.stmartin_df, selected_cats or None) \
        if st.session_state.stmartin_df is not None else []
    all_items   = sorted(set(items_kings) | set(items_stm))

    # Row B — exclude products (scoped to selected categories)
    _, rb_col = st.columns([0.001, 1])
    with rb_col:
        if all_items:
            excluded_items = st.multiselect(
                "🚫 Exclude Products",
                options=all_items,
                default=[],
                key="item_filter",
                placeholder="Select products to exclude from report…",
            )
        else:
            excluded_items = []

    st.markdown("</div>", unsafe_allow_html=True)



# ─────────────────────────────────────────────────────────────────────────────
# File status badges
# ─────────────────────────────────────────────────────────────────────────────
kings_ready = st.session_state.kings_df is not None
stm_ready   = st.session_state.stmartin_df is not None
files_ready = kings_ready and stm_ready

sb1, sb2, sb3, sb4 = st.columns(4)
with sb1:
    if kings_ready:
        st.success(f"✅ Kings Road — {len(st.session_state.kings_df):,} rows loaded")
    else:
        st.warning("⚠️ Kings Road file not loaded")
with sb2:
    if stm_ready:
        st.success(f"✅ St Martin — {len(st.session_state.stmartin_df):,} rows loaded")
    else:
        st.warning("⚠️ St Martin file not loaded")
with sb3:
    if selected_cats:
        st.info(f"🏷️ Categories: {', '.join(selected_cats)}")
    elif all_cats:
        st.info("🏷️ All categories included")
with sb4:
    if excluded_items:
        st.warning(f"🚫 Excluded: {', '.join(excluded_items)}")
    elif all_items:
        st.info("📦 No products excluded")

st.markdown("---")


# ─────────────────────────────────────────────────────────────────────────────
# REPORT TYPE BUTTONS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-label">📋 Select Report</div>', unsafe_allow_html=True)

# Determine which button is "active" to style it differently
active = st.session_state.active_report

btn_col1, btn_col2, btn_col3, btn_col4 = st.columns([1, 1, 1, 1])

with btn_col1:
    top10_type = "primary" if active == "top10" else "secondary"
    if st.button("🏆  Top 10 Best Sellers", type=top10_type,
                 use_container_width=True, key="btn_top10",
                 disabled=not files_ready):
        st.session_state.active_report = "top10"
        st.rerun()

with btn_col2:
    worst5_type = "primary" if active == "worst5" else "secondary"
    if st.button("📉  Worst 5 Sellers", type=worst5_type,
                 use_container_width=True, key="btn_worst5",
                 disabled=not files_ready):
        st.session_state.active_report = "worst5"
        st.rerun()

with btn_col3:
    allitems_type = "primary" if active == "all_items" else "secondary"
    if st.button("📋  All Items", type=allitems_type,
                 use_container_width=True, key="btn_all_items",
                 disabled=not files_ready):
        st.session_state.active_report = "all_items"
        st.rerun()

with btn_col4:
    cattotals_type = "primary" if active == "category_totals" else "secondary"
    if st.button("📊  Category Totals", type=cattotals_type,
                 use_container_width=True, key="btn_cat_totals",
                 disabled=not files_ready):
        st.session_state.active_report = "category_totals"
        st.rerun()

if not files_ready:
    st.caption("Upload both outlet files above to enable report generation.")


# ─────────────────────────────────────────────────────────────────────────────
# REPORT GENERATION (runs whenever active_report is set)
# ─────────────────────────────────────────────────────────────────────────────
def _generate(report_type: str):
    """Run the chosen report type and cache results in session state."""
    cats  = selected_cats  if selected_cats  else None
    excl  = excluded_items if excluded_items else None
    kdf   = st.session_state.kings_df
    sdf   = st.session_state.stmartin_df

    if start_date > end_date:
        st.error("'From' date must be on or before the 'To' date.")
        return

    with st.spinner("Generating report…"):
        try:
            if report_type == "top10":
                rc, tc = top10_combined(kdf, sdf, start_date, end_date, categories=cats, exclude_items=excl)
                rk, tk = top10_single(kdf, "Kings", start_date, end_date, categories=cats, exclude_items=excl)
                rs, ts = top10_single(sdf, "St Martin", start_date, end_date, categories=cats, exclude_items=excl)
            elif report_type == "worst5":
                rc, tc = worst5_combined(kdf, sdf, start_date, end_date, categories=cats, exclude_items=excl)
                rk, tk = worst5_single(kdf, "Kings", start_date, end_date, categories=cats, exclude_items=excl)
                rs, ts = worst5_single(sdf, "St Martin", start_date, end_date, categories=cats, exclude_items=excl)
            elif report_type == "all_items":
                rc, tc = all_items_combined(kdf, sdf, start_date, end_date, categories=cats, exclude_items=excl)
                rk, tk = all_items_single(kdf, "Kings", start_date, end_date, categories=cats, exclude_items=excl)
                rs, ts = all_items_single(sdf, "St Martin", start_date, end_date, categories=cats, exclude_items=excl)
            elif report_type == "category_totals":
                rc, tc = category_totals_combined(kdf, sdf, start_date, end_date, categories=cats, exclude_items=excl)
                rk, tk = category_totals_single(kdf, "Kings", start_date, end_date, categories=cats, exclude_items=excl)
                rs, ts = category_totals_single(sdf, "St Martin", start_date, end_date, categories=cats, exclude_items=excl)
            else:
                return
        except Exception as e:
            st.error(f"Report error: {e}")
            return

        # Check all empty
        if all(len(df) <= 1 for df in [rc, rk, rs]):
            st.warning("No data found for the selected date range / category filter.")
            return

        # Build Excel
        try:
            excel_bytes = build_excel(rc, tc, rk, tk, rs, ts)
        except Exception as e:
            st.error(f"Excel export error: {e}")
            return

        st.session_state.report_combined  = rc
        st.session_state.report_kings     = rk
        st.session_state.report_stmartin  = rs
        st.session_state.title_combined   = tc
        st.session_state.title_kings      = tk
        st.session_state.title_stmartin   = ts
        st.session_state.excel_bytes      = excel_bytes


# ─────────────────────────────────────────────────────────────────────────────
# REPORT OUTPUT AREA
# ─────────────────────────────────────────────────────────────────────────────

# Only generate when a report button was just clicked (active_report set this rerun)
if "_last_generated" not in st.session_state:
    st.session_state["_last_generated"] = None

if st.session_state.active_report and files_ready:
    # Regenerate only when report type or key inputs have changed
    run_key = (
        st.session_state.active_report,
        str(start_date), str(end_date),
        tuple(sorted(selected_cats)),
        tuple(sorted(excluded_items)),
        id(st.session_state.kings_df),
        id(st.session_state.stmartin_df),
    )
    if run_key != st.session_state["_last_generated"]:
        _generate(st.session_state.active_report)
        st.session_state["_last_generated"] = run_key

if st.session_state.excel_bytes is not None:
    tc = st.session_state.title_combined
    tk = st.session_state.title_kings
    ts = st.session_state.title_stmartin
    rc = st.session_state.report_combined
    rk = st.session_state.report_kings
    rs = st.session_state.report_stmartin

    safe = tc.replace(" \u2013 ", "_").replace(" ", "_").replace("/", "-")
    fname = f"{safe}.xlsx"

    st.markdown("---")
    st.success(f"Report ready: **{tc}**")

    # ── Preview tabs ───────────────────────────────────────────────────────
    def _preview(df: pd.DataFrame, title: str):
        """Show data rows then the total row in a highlighted block."""
        st.caption(title)
        n = len(df) - 1          # all rows except last (total)
        data_rows  = df.iloc[:n].copy()
        total_row  = df.iloc[[n]].copy()

        # Clean up #SL for display
        if "#SL" in data_rows.columns:
            data_rows["#SL"] = pd.to_numeric(
                data_rows["#SL"], errors="coerce"
            ).astype("Int64")

        st.dataframe(data_rows,  use_container_width=True, hide_index=True)
        st.dataframe(total_row,  use_container_width=True, hide_index=True)

    tab1, tab2, tab3 = st.tabs(["📊 Combined", "🔶 Kings Road", "🟢 St Martin"])
    with tab1: _preview(rc, tc)
    with tab2: _preview(rk, tk)
    with tab3: _preview(rs, ts)

    # ── Download ───────────────────────────────────────────────────────────
    st.markdown("---")
    dl_col, _ = st.columns([2, 3])
    with dl_col:
        st.download_button(
            label="\u2b07\ufe0f  Download Excel Report (.xlsx)",
            data=st.session_state.excel_bytes,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="dl_btn",
        )

elif not files_ready:
    # Idle welcome card
    st.markdown("""
<div style="
    background:rgba(20,38,60,0.85);
    border:1px solid #2a4568;
    border-radius:12px;
    padding:28px 32px;
    margin-top:8px;">
  <h3 style="color:#c8dff0;margin-top:0;">🚀 Getting started</h3>
  <ol style="color:#8ab4d4;line-height:2;">
    <li>Upload <strong style="color:#c8dff0;">Kings Road</strong> and
        <strong style="color:#c8dff0;">St Martin</strong> Excel files above.</li>
    <li>Set your <strong style="color:#c8dff0;">date range</strong>.</li>
    <li>Optionally choose <strong style="color:#c8dff0;">categories</strong> to include.</li>
    <li>Click a <strong style="color:#c8dff0;">report button</strong> — Top 10 or Worst 5.</li>
    <li>Preview and <strong style="color:#c8dff0;">download</strong> the Excel report.</li>
  </ol>
  <p style="color:#4a7a9e;margin:0;font-size:0.85rem;">
    Expected columns: <code>Date</code> | <code>Item Name</code> | <code>Qty</code>
    (+ optional <code>Category</code>)
  </p>
</div>
""", unsafe_allow_html=True)
