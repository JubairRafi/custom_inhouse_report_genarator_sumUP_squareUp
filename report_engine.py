"""
report_engine.py
================
Pure data-transformation logic for the Report Generator Manager.
No I/O, no formatting — only aggregation and ranking.

Design
------
All report logic goes through two generic functions:
    ranked_combined(kings_df, stmartin_df, start, end, n, ascending, categories)
    ranked_single(df, outlet_label, start, end, n, ascending, categories)

Report-type helpers (top10_*, worst5_*) are thin wrappers that set n/ascending,
keeping the call-sites in app.py very simple and making it trivial to add
new report types (e.g. "Worst 10", "Top 5") in the future.
"""

from __future__ import annotations

from datetime import date

import pandas as pd

from data_loader import filter_by_date_and_category, COL_ITEM, COL_QTY, COL_CATEGORY

# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _date_range_label(start_date: date, end_date: date) -> str:
    """Short label: '01 Jan 2025 – 31 Jan 2025'."""
    fmt = "%d %b %Y"
    return f"{start_date.strftime(fmt)} \u2013 {end_date.strftime(fmt)}"


def _date_range_label_long(start_date: date, end_date: date) -> str:
    """Long label: '23 February to 01 March 2025'."""
    return f"{start_date.strftime('%d %B')} to {end_date.strftime('%d %B %Y')}"


def _group_by_item(df: pd.DataFrame) -> pd.Series:
    """Aggregate Qty by Item Name."""
    return df.groupby(COL_ITEM)[COL_QTY].sum()


def _group_by_category(df: pd.DataFrame) -> pd.Series:
    """Aggregate Qty by Category."""
    if COL_CATEGORY not in df.columns:
        # Fallback if categories are missing
        return df.assign(**{COL_CATEGORY: "Uncategorised"}).groupby(COL_CATEGORY)[COL_QTY].sum()
    return df.groupby(COL_CATEGORY)[COL_QTY].sum()


def _add_serial_column(df: pd.DataFrame) -> pd.DataFrame:
    """Prepend a '#SL' serial-number column (1-based)."""
    df = df.reset_index(drop=True)
    df.insert(0, "#SL", range(1, len(df) + 1))
    return df


def _append_total_row(df: pd.DataFrame, qty_columns: list[str]) -> pd.DataFrame:
    """
    Append a grand-total summary row.
    '#SL' left empty; Item Name = 'Total'; qty_columns summed.
    """
    total: dict = {"#SL": "", "Item Name": "Total"}
    for col in qty_columns:
        total[col] = df[col].sum()
    total_row = pd.DataFrame([total], columns=df.columns)
    return pd.concat([df, total_row], ignore_index=True)


# ---------------------------------------------------------------------------
# Generic ranked engines  (the two core functions everything builds on)
# ---------------------------------------------------------------------------

def ranked_combined(
    kings_df: pd.DataFrame,
    stmartin_df: pd.DataFrame,
    start_date: date,
    end_date: date,
    n: int | None = 10,
    ascending: bool = False,
    categories: list[str] | None = None,
    exclude_items: list[str] | None = None,
    add_serial: bool = True,
) -> tuple[pd.DataFrame, str]:
    """
    Combined report for both outlets.

    Parameters
    ----------
    n             : max rows to return; None = all items.
    ascending     : False = best sellers; True = worst sellers.
    categories    : optional category whitelist; None = all.
    exclude_items : optional item blacklist; None = exclude nothing.
    add_serial    : whether to prepend the #SL column.
    """
    if kings_df is None or stmartin_df is None:
        raise ValueError("Both Kings Road and St Martin files must be provided.")

    kings_f    = filter_by_date_and_category(kings_df,    start_date, end_date, categories, exclude_items)
    stmartin_f = filter_by_date_and_category(stmartin_df, start_date, end_date, categories, exclude_items)

    kings_g    = _group_by_item(kings_f).rename("Kings Road")
    stmartin_g = _group_by_item(stmartin_f).rename("St Martin")

    merged = pd.merge(
        kings_g.reset_index(),
        stmartin_g.reset_index(),
        on=COL_ITEM,
        how="outer",
    ).fillna(0)

    merged["Kings Road"] = merged["Kings Road"].astype(int)
    merged["St Martin"]  = merged["St Martin"].astype(int)
    merged["Total"]      = merged["Kings Road"] + merged["St Martin"]
    merged = merged.sort_values("Total", ascending=ascending)
    if n is not None:
        merged = merged.head(n)
    merged = merged.rename(columns={COL_ITEM: "Item Name"})

    report = merged[["Item Name", "Kings Road", "St Martin", "Total"]].copy()
    if add_serial:
        report = _add_serial_column(report)
    report = _append_total_row(report, ["Kings Road", "St Martin", "Total"])

    label     = _date_range_label(start_date, end_date)
    direction = "Worst" if ascending else "Top"
    count_str = str(n) if n is not None else "All"
    title     = f"{label} \u2013 {direction} {count_str} Best Seller St George"
    return report, title


def ranked_single(
    df: pd.DataFrame,
    outlet_label: str,
    start_date: date,
    end_date: date,
    n: int | None = 10,
    ascending: bool = False,
    categories: list[str] | None = None,
    exclude_items: list[str] | None = None,
    add_serial: bool = True,
) -> tuple[pd.DataFrame, str]:
    """
    Single-outlet ranked report.

    Parameters
    ----------
    outlet_label  : display name, e.g. 'Kings' or 'St Martin'.
    n             : max rows; None = all items.
    ascending     : False = best; True = worst.
    categories    : optional category whitelist.
    exclude_items : optional item blacklist.
    add_serial    : whether to prepend the #SL column.
    """
    if df is None:
        raise ValueError(f"No data provided for outlet '{outlet_label}'.")

    filtered = filter_by_date_and_category(df, start_date, end_date, categories, exclude_items)
    grouped  = _group_by_item(filtered).reset_index()
    grouped.columns = ["Item Name", "Qty"]
    grouped = grouped.sort_values("Qty", ascending=ascending)
    if n is not None:
        grouped = grouped.head(n)
    grouped["Qty"] = grouped["Qty"].astype(int)

    report = grouped.copy()
    if add_serial:
        report = _add_serial_column(report)
    report = _append_total_row(report, ["Qty"])

    direction = "Worst" if ascending else "Top"
    count_str = str(n) if n is not None else "All"
    label     = _date_range_label(start_date, end_date)
    title     = f"{label} \u2013 {direction} {count_str} {outlet_label}"
    return report, title


# ---------------------------------------------------------------------------
# Named report helpers — kept thin so app.py stays simple
# ---------------------------------------------------------------------------

def top10_combined(kings_df, stmartin_df, start, end, categories=None, exclude_items=None):
    """Top 10 Combined Best Sellers."""
    return ranked_combined(kings_df, stmartin_df, start, end,
                           n=10, ascending=False, categories=categories,
                           exclude_items=exclude_items, add_serial=True)


def top10_single(df, outlet_label, start, end, categories=None, exclude_items=None):
    """Top 10 Single-Outlet Best Sellers."""
    return ranked_single(df, outlet_label, start, end,
                         n=10, ascending=False, categories=categories,
                         exclude_items=exclude_items, add_serial=True)


def worst5_combined(kings_df, stmartin_df, start, end, categories=None, exclude_items=None):
    """Worst 5 Combined Sellers."""
    return ranked_combined(kings_df, stmartin_df, start, end,
                           n=5, ascending=True, categories=categories,
                           exclude_items=exclude_items, add_serial=True)


def worst5_single(df, outlet_label, start, end, categories=None, exclude_items=None):
    """Worst 5 Single-Outlet Sellers."""
    return ranked_single(df, outlet_label, start, end,
                         n=5, ascending=True, categories=categories,
                         exclude_items=exclude_items, add_serial=True)


def all_items_combined(
    kings_df, stmartin_df, start, end, categories=None, exclude_items=None
):
    """
    All items combined — no row limit, no #SL, title in long format.
    Returns (DataFrame, title_str).
    """
    rc, _ = ranked_combined(
        kings_df, stmartin_df, start, end,
        n=None, ascending=False, categories=categories,
        exclude_items=exclude_items, add_serial=False,
    )
    title = _date_range_label_long(start, end)
    return rc, title


def all_items_single(
    df, outlet_label, start, end, categories=None, exclude_items=None
):
    """
    All items for one outlet — no row limit, no #SL, title in long format.
    """
    rs, _ = ranked_single(
        df, outlet_label, start, end,
        n=None, ascending=False, categories=categories,
        exclude_items=exclude_items, add_serial=False,
    )
    title = _date_range_label_long(start, end)
    return rs, title


def category_totals_combined(
    kings_df: pd.DataFrame,
    stmartin_df: pd.DataFrame,
    start_date: date,
    end_date: date,
    categories: list[str] | None = None,
    exclude_items: list[str] | None = None,
) -> tuple[pd.DataFrame, str]:
    """
    All categories combined — no row limit, no #SL, first column named after date range.
    Returns (DataFrame, title_str).
    """
    if kings_df is None or stmartin_df is None:
        raise ValueError("Both Kings Road and St Martin files must be provided.")

    kings_f    = filter_by_date_and_category(kings_df,    start_date, end_date, categories, exclude_items)
    stmartin_f = filter_by_date_and_category(stmartin_df, start_date, end_date, categories, exclude_items)

    kings_g    = _group_by_category(kings_f).rename("Kings") # Match screenshot column header
    stmartin_g = _group_by_category(stmartin_f).rename("St Martin")

    merged = pd.merge(
        stmartin_g.reset_index(), # Screenshot has St Martin first
        kings_g.reset_index(),
        on=COL_CATEGORY,
        how="outer",
    ).fillna(0)

    merged["St Martin"] = merged["St Martin"].astype(int)
    merged["Kings"]  = merged["Kings"].astype(int)
    merged["Total"]  = merged["St Martin"] + merged["Kings"]

    # Format the category name logic (append " Total" like the screenshot "Pastries Total")
    merged[COL_CATEGORY] = merged[COL_CATEGORY].astype(str) + " Total"

    merged = merged.sort_values("Total", ascending=False)
    
    label = _date_range_label_long(start_date, end_date)
    merged = merged.rename(columns={COL_CATEGORY: label})

    report = merged[[label, "St Martin", "Kings", "Total"]].copy()
    report = _append_total_row(report, ["St Martin", "Kings", "Total"])

    title = label
    return report, title


def category_totals_single(
    df: pd.DataFrame,
    outlet_label: str,
    start_date: date,
    end_date: date,
    categories: list[str] | None = None,
    exclude_items: list[str] | None = None,
) -> tuple[pd.DataFrame, str]:
    """
    Category totals for one outlet.
    """
    if df is None:
        raise ValueError(f"No data provided for outlet '{outlet_label}'.")

    filtered = filter_by_date_and_category(df, start_date, end_date, categories, exclude_items)
    grouped  = _group_by_category(filtered).reset_index()
    grouped.columns = [COL_CATEGORY, "Qty"]

    grouped[COL_CATEGORY] = grouped[COL_CATEGORY].astype(str) + " Total"
    grouped = grouped.sort_values("Qty", ascending=False)
    grouped["Qty"] = grouped["Qty"].astype(int)

    label = _date_range_label_long(start_date, end_date)
    grouped = grouped.rename(columns={COL_CATEGORY: label})

    report = grouped.copy()
    report = _append_total_row(report, ["Qty"])

    title = f"{label} \u2013 {outlet_label}"
    return report, title
