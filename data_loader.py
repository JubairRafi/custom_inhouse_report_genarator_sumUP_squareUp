"""
data_loader.py
==============
Handles all I/O and date filtering for the Report Generator Manager.

Responsibilities
----------------
- Read uploaded Excel files into DataFrames.
- Normalise column names so the rest of the pipeline is robust to minor
  header variations (extra spaces, casing, etc.).
- Provide a centralised date-filter function used by every report.

Expected minimum columns (after normalisation):
    Date | Item Name | Qty
"""

from __future__ import annotations

import io
from datetime import date

import pandas as pd

# ---------------------------------------------------------------------------
# Column-name aliases — add more aliases here if outlet files ever differ
# ---------------------------------------------------------------------------
DATE_ALIASES: list[str] = ["date", "sales date", "sale date", "trans date"]
ITEM_ALIASES: list[str] = ["item name", "item", "product", "product name", "description"]
QTY_ALIASES: list[str] = ["qty", "quantity", "sold qty", "units", "units sold"]
CATEGORY_ALIASES: list[str] = [
    "category", "cat", "product category", "item category",
    "type", "dept", "department", "group", "section",
]

# Canonical internal names
COL_DATE = "Date"
COL_ITEM = "Item Name"
COL_QTY = "Qty"
COL_CATEGORY = "Category"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _find_column(columns: pd.Index, aliases: list[str]) -> str | None:
    """Return the first column in *columns* whose lower-stripped name is in *aliases*."""
    normalised = {c.strip().lower(): c for c in columns}
    for alias in aliases:
        if alias in normalised:
            return normalised[alias]
    return None


def _normalise_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Rename recognised columns to their canonical names.
    Category column is optional — normalised if found, silently ignored otherwise.
    Raises ValueError if a *required* column cannot be found.
    """
    rename_map: dict[str, str] = {}

    date_col = _find_column(df.columns, DATE_ALIASES)
    item_col = _find_column(df.columns, ITEM_ALIASES)
    qty_col = _find_column(df.columns, QTY_ALIASES)
    cat_col = _find_column(df.columns, CATEGORY_ALIASES)

    missing = []
    if date_col is None:
        missing.append("Date")
    else:
        rename_map[date_col] = COL_DATE

    if item_col is None:
        missing.append("Item Name")
    else:
        rename_map[item_col] = COL_ITEM

    if qty_col is None:
        missing.append("Qty")
    else:
        rename_map[qty_col] = COL_QTY

    if cat_col is not None:
        rename_map[cat_col] = COL_CATEGORY

    if missing:
        raise ValueError(
            f"Could not find required column(s): {', '.join(missing)}.\n"
            f"Columns found in file: {list(df.columns)}"
        )

    df = df.rename(columns=rename_map)
    df[COL_DATE] = pd.to_datetime(df[COL_DATE], errors="coerce")
    df[COL_QTY] = pd.to_numeric(df[COL_QTY], errors="coerce").fillna(0)
    return df


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def load_outlet_file(uploaded_file) -> pd.DataFrame:
    """
    Read an uploaded Excel file (Streamlit UploadedFile or file-like object)
    and return a normalised DataFrame with canonical column names.

    Parameters
    ----------
    uploaded_file : Streamlit UploadedFile or file-like
        The Excel file to load.

    Returns
    -------
    pd.DataFrame
        DataFrame with columns: Date, Item Name, Qty (plus any extra columns).

    Raises
    ------
    ValueError
        If required columns are missing.
    """
    raw = pd.read_excel(uploaded_file, sheet_name=0)
    df = _normalise_columns(raw)
    return df


def get_categories(df: pd.DataFrame) -> list[str]:
    """
    Return a sorted list of unique category values found in the DataFrame.
    Returns an empty list if the Category column is not present.
    """
    if COL_CATEGORY not in df.columns:
        return []
    return sorted(df[COL_CATEGORY].dropna().astype(str).unique().tolist())


def get_items(df: pd.DataFrame) -> list[str]:
    """
    Return a sorted list of unique item names found in the DataFrame.
    """
    if COL_ITEM not in df.columns:
        return []
    return sorted(df[COL_ITEM].dropna().astype(str).unique().tolist())


def get_items_for_categories(
    df: pd.DataFrame,
    categories: list[str] | None = None,
) -> list[str]:
    """
    Return item names scoped to the given categories.
    If categories is None or empty, returns all item names.
    """
    if COL_ITEM not in df.columns:
        return []
    if categories and COL_CATEGORY in df.columns:
        df = df[df[COL_CATEGORY].astype(str).isin(categories)]
    return sorted(df[COL_ITEM].dropna().astype(str).unique().tolist())


def filter_by_date(df: pd.DataFrame, start_date: date, end_date: date) -> pd.DataFrame:
    """
    Return rows where Date falls within [start_date, end_date] (inclusive).
    Kept for backwards compatibility — delegates to filter_by_date_and_category.
    """
    return filter_by_date_and_category(df, start_date, end_date, categories=None)


def filter_by_date_and_category(
    df: pd.DataFrame,
    start_date: date,
    end_date: date,
    categories: list[str] | None = None,
    exclude_items: list[str] | None = None,
) -> pd.DataFrame:
    """
    Filter rows by date range, optional category whitelist, and optional item blacklist.

    Parameters
    ----------
    categories    : include only rows whose Category is in this list (None = all).
    exclude_items : remove rows whose Item Name is in this list (None = remove nothing).
    """
    mask = (df[COL_DATE].dt.date >= start_date) & (df[COL_DATE].dt.date <= end_date)
    result = df.loc[mask].copy()

    if categories and COL_CATEGORY in result.columns:
        result = result[result[COL_CATEGORY].astype(str).isin(categories)]

    if exclude_items and COL_ITEM in result.columns:
        result = result[~result[COL_ITEM].astype(str).isin(exclude_items)]

    return result
