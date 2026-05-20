"""
square_report_engine.py
=======================
Report generation for Square POS data.
All reports include both Quantity and Net Sales columns.
"""
from __future__ import annotations

import datetime

import pandas as pd

from square_loader import (
    COL_DATE, COL_ITEM, COL_CATEGORY, COL_QTY, COL_NET, COL_OUTLET,
    filter_sq,
)

_FMT = "%d/%m/%y"

def _fmt(d) -> str:
    if isinstance(d, datetime.date):
        return d.strftime(_FMT)
    return str(d)


def _title(report_name: str, label: str, start, end) -> str:
    # Shorten verbose names
    short_name = (
        report_name
        .replace(" Best Sellers", "")
        .replace(" Sellers", "")
        .replace(" Totals", "")
    )
    if label == "Combined":
        label = "Comb."  # keep it even shorter
    return f"{short_name} {label} ({_fmt(start)}-{_fmt(end)})"


def _summarise_items(fdf: pd.DataFrame, breakdown: bool = False) -> pd.DataFrame:
    """Group by Item + Category, sum Qty and Net Sales."""
    if fdf.empty:
        return pd.DataFrame(columns=[COL_ITEM, COL_CATEGORY, "Qty", "Net Sales £"])
    
    totals = (
        fdf.groupby([COL_ITEM, COL_CATEGORY], as_index=False)
           .agg(Qty=(COL_QTY, "sum"), **{"Net Sales £": (COL_NET, "sum")})
    )
    if not breakdown or fdf.empty or COL_OUTLET not in fdf.columns or fdf[COL_OUTLET].nunique() <= 1:
        return totals

    # Outlet breakdown
    pivot_qty = fdf.pivot_table(index=[COL_ITEM, COL_CATEGORY], columns=COL_OUTLET, values=COL_QTY, aggfunc='sum', fill_value=0)
    pivot_net = fdf.pivot_table(index=[COL_ITEM, COL_CATEGORY], columns=COL_OUTLET, values=COL_NET, aggfunc='sum', fill_value=0)
    
    pivot_qty.columns = [f"{c} Qty" for c in pivot_qty.columns]
    pivot_net.columns = [f"{c} Net £" for c in pivot_net.columns]
    
    comb = pd.merge(totals, pivot_qty, on=[COL_ITEM, COL_CATEGORY], how='left')
    comb = pd.merge(comb, pivot_net, on=[COL_ITEM, COL_CATEGORY], how='left')
    
    # Interleave
    out_cols = [COL_ITEM, COL_CATEGORY]
    for o in sorted(fdf[COL_OUTLET].unique()):
        out_cols.append(f"{o} Qty")
        out_cols.append(f"{o} Net £")
    out_cols.extend(["Qty", "Net Sales £"])
    
    return comb[out_cols].fillna(0)


def _summarise_categories(fdf: pd.DataFrame, breakdown: bool = False) -> pd.DataFrame:
    """Group by Category, sum Qty and Net Sales."""
    if fdf.empty:
        return pd.DataFrame(columns=[COL_CATEGORY, "Qty", "Net Sales £"])
    
    totals = (
        fdf.groupby(COL_CATEGORY, as_index=False)
           .agg(Qty=(COL_QTY, "sum"), **{"Net Sales £": (COL_NET, "sum")})
    )
    if not breakdown or fdf.empty or COL_OUTLET not in fdf.columns or fdf[COL_OUTLET].nunique() <= 1:
        return totals

    pivot_qty = fdf.pivot_table(index=COL_CATEGORY, columns=COL_OUTLET, values=COL_QTY, aggfunc='sum', fill_value=0)
    pivot_net = fdf.pivot_table(index=COL_CATEGORY, columns=COL_OUTLET, values=COL_NET, aggfunc='sum', fill_value=0)
    
    pivot_qty.columns = [f"{c} Qty" for c in pivot_qty.columns]
    pivot_net.columns = [f"{c} Net £" for c in pivot_net.columns]
    
    comb = pd.merge(totals, pivot_qty, on=COL_CATEGORY, how='left')
    comb = pd.merge(comb, pivot_net, on=COL_CATEGORY, how='left')
    
    out_cols = [COL_CATEGORY]
    for o in sorted(fdf[COL_OUTLET].unique()):
        out_cols.append(f"{o} Qty")
        out_cols.append(f"{o} Net £")
    out_cols.extend(["Qty", "Net Sales £"])
    
    return comb[out_cols].fillna(0)


def _finalise_items(df: pd.DataFrame) -> pd.DataFrame:
    """Add rank column and TOTAL footer row."""
    df = df.copy().reset_index(drop=True)
    df.insert(0, "#SL", range(1, len(df) + 1))
    
    total = {"#SL": "", COL_ITEM: "TOTAL", COL_CATEGORY: ""}
    for col in df.columns:
        if col not in ["#SL", COL_ITEM, COL_CATEGORY]:
            total[col] = df[col].sum()
            
    return pd.concat([df, pd.DataFrame([total])], ignore_index=True)


def _finalise_cats(df: pd.DataFrame) -> pd.DataFrame:
    """Add rank column and TOTAL footer row (category report)."""
    df = df.copy().reset_index(drop=True)
    df.insert(0, "#SL", range(1, len(df) + 1))
    
    total = {"#SL": "", COL_CATEGORY: "TOTAL"}
    for col in df.columns:
        if col not in ["#SL", COL_CATEGORY]:
            total[col] = df[col].sum()
            
    return pd.concat([df, pd.DataFrame([total])], ignore_index=True)


# ── Public report functions ───────────────────────────────────────────────────

def top10_sq(df, label, start, end,
             outlets=None, categories=None, exclude_items=None):
    fdf  = filter_sq(df, start, end, outlets, categories, exclude_items)
    is_comb = (label == "Combined")
    summ = _summarise_items(fdf, breakdown=is_comb).sort_values("Qty", ascending=False).head(10)
    return _finalise_items(summ), _title("Top 10 Best Sellers", label, start, end)


def worst5_sq(df, label, start, end,
              outlets=None, categories=None, exclude_items=None):
    fdf  = filter_sq(df, start, end, outlets, categories, exclude_items)
    is_comb = (label == "Combined")
    summ = _summarise_items(fdf, breakdown=is_comb).sort_values("Qty", ascending=True).head(5)
    return _finalise_items(summ), _title("Worst 5 Sellers", label, start, end)


def all_items_sq(df, label, start, end,
                 outlets=None, categories=None, exclude_items=None):
    fdf  = filter_sq(df, start, end, outlets, categories, exclude_items)
    is_comb = (label == "Combined")
    summ = _summarise_items(fdf, breakdown=is_comb).sort_values("Qty", ascending=False)
    return _finalise_items(summ), _title("All Items", label, start, end)


def category_totals_sq(df, label, start, end,
                       outlets=None, categories=None, exclude_items=None):
    fdf  = filter_sq(df, start, end, outlets, categories, exclude_items)
    is_comb = (label == "Combined")
    summ = _summarise_categories(fdf, breakdown=is_comb).sort_values("Qty", ascending=False)
    return _finalise_cats(summ), _title("Category Totals", label, start, end)


def weekly_by_day_sq(df, label, start, end,
                     outlets=None, categories=None, exclude_items=None):
    fdf = filter_sq(df, start, end, outlets, categories, exclude_items)
    
    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    cols_order = [COL_ITEM] + days + ["Total"]

    if fdf.empty:
        return pd.DataFrame(columns=cols_order), _title("Weekly", label, start, end)

    fdf = fdf.copy()
    fdf["Day"] = fdf[COL_DATE].dt.day_name()
    
    pivot = fdf.pivot_table(
        index=[COL_ITEM, COL_CATEGORY],
        columns="Day",
        values=COL_QTY,
        aggfunc="sum",
        fill_value=0
    ).reset_index()

    for d in days:
        if d not in pivot.columns:
            pivot[d] = 0

    pivot["Total"] = pivot[days].sum(axis=1)
    pivot = pivot.sort_values(by=[COL_CATEGORY, COL_ITEM], ascending=True)

    item_rows = []
    cat_totals = []
    
    # Build item rows and category subtotals
    for cat, group in pivot.groupby(COL_CATEGORY):
        grp_items = group.drop(columns=[COL_CATEGORY])
        item_rows.append(grp_items)
        
        subtotal = {COL_ITEM: f"{cat} Total"}
        for d in days + ["Total"]:
            subtotal[d] = group[d].sum()
        cat_totals.append(pd.DataFrame([subtotal]))
        
    df_items = pd.concat(item_rows, ignore_index=True) if item_rows else pd.DataFrame()
    df_cats  = pd.concat(cat_totals, ignore_index=True) if cat_totals else pd.DataFrame()
    
    # Blank row separator
    blank = {c: "" for c in cols_order}
    
    # Grand total
    grand = {COL_ITEM: "Total"}
    for d in days + ["Total"]:
        grand[d] = pivot[d].sum()
        
    final_dfs = [df_items, pd.DataFrame([blank]), df_cats, pd.DataFrame([grand]), pd.DataFrame([blank])]
    df_out = pd.concat(final_dfs, ignore_index=True)

    return df_out[cols_order], _title("Weekly", label, start, end)
