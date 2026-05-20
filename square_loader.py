"""
square_loader.py
================
Handles Square POS CSV export files.
Format: UTF-16, tab-separated, French column headers.
"""
from __future__ import annotations

import io
from datetime import date

import pandas as pd

# ── Raw Square column names (French) ─────────────────────────────────────────
_SQ_DATE      = "Date"
_SQ_CATEGORY  = "Catégorie"
_SQ_ARTICLE   = "Article"
_SQ_QTY       = "Qté"
_SQ_NET       = "Ventes nettes"
_SQ_GROSS     = "Ventes brutes"
_SQ_OUTLET    = "Point de vente"
_SQ_ACTIVITY  = "Type d'activité"

# ── Canonical column names used throughout the pipeline ──────────────────────
COL_DATE     = "Date"
COL_ITEM     = "Item Name"
COL_CATEGORY = "Category"
COL_QTY      = "Qty"
COL_NET      = "Net Sales"
COL_GROSS    = "Gross Sales"
COL_OUTLET   = "Outlet"


# ── Parsers ──────────────────────────────────────────────────────────────────

def _parse_euro_currency(series: pd.Series) -> pd.Series:
    """'4,50 £'  →  4.50  (float)"""
    return (
        series.astype(str)
              .str.replace("£", "", regex=False)
              .str.replace("\xa0", "", regex=False)
              .str.replace("\u202f", "", regex=False)
              .str.replace(" ", "", regex=False)
              .str.replace(",", ".", regex=False)
              .pipe(pd.to_numeric, errors="coerce")
              .fillna(0.0)
    )


def _parse_euro_float(series: pd.Series) -> pd.Series:
    """'2,0'  →  2.0  (float)"""
    return (
        series.astype(str)
              .str.replace(",", ".", regex=False)
              .str.strip()
              .pipe(pd.to_numeric, errors="coerce")
              .fillna(0.0)
    )


def _shorten_outlet_name(name: str) -> str:
    """Shorten long POS outlet names to just their location."""
    name = str(name).strip()
    replace_map = [
        ("St George Coffee & Wine Bar ", ""),
        ("St George Coffee and Wine Bar ", ""),
        ("Chill Since '93 10NR", "Chill Since NR"),
        ("Chill Since '93 61KR", "Chill Since KR"),
        ("Chill Since '93 ", ""),
        ("112 ST Martin", "St Martin"),
    ]
    for old, new in replace_map:
        name = name.replace(old, new)
    return name.strip()


# ── Public API ────────────────────────────────────────────────────────────────

def load_square_csv(uploaded_file) -> pd.DataFrame:
    """
    Read a Square POS CSV export and return a normalised DataFrame.

    Parameters
    ----------
    uploaded_file : Streamlit UploadedFile or path-like

    Returns
    -------
    pd.DataFrame with columns: Date, Item Name, Category,
                               Qty, Net Sales, Gross Sales, Outlet

    Raises
    ------
    ValueError — if required columns are missing.
    """
    if hasattr(uploaded_file, "read"):
        raw = uploaded_file.read()
    else:
        with open(uploaded_file, "rb") as fh:
            raw = fh.read()

    content = raw.decode("utf-16")
    df_raw = pd.read_csv(io.StringIO(content), sep="\t", low_memory=False)

    # Validate required columns
    required = {_SQ_DATE, _SQ_ARTICLE, _SQ_QTY}
    missing = required - set(df_raw.columns)
    if missing:
        raise ValueError(
            f"Required column(s) not found: {missing}\n"
            f"Columns in file: {list(df_raw.columns)}"
        )

    # Keep payment rows only (exclude refunds, voids, etc.)
    if _SQ_ACTIVITY in df_raw.columns:
        mask = df_raw[_SQ_ACTIVITY].astype(str).str.strip().str.lower().isin(
            ["paiement", "payment"]
        )
        df_raw = df_raw[mask].copy()

    # Build canonical DataFrame
    out = pd.DataFrame()
    out[COL_DATE]     = pd.to_datetime(df_raw[_SQ_DATE], errors="coerce", dayfirst=False)
    out[COL_ITEM]     = df_raw[_SQ_ARTICLE].astype(str).str.strip()
    out[COL_CATEGORY] = (
        df_raw[_SQ_CATEGORY].astype(str).str.strip()
        if _SQ_CATEGORY in df_raw.columns else "Uncategorised"
    )
    out[COL_QTY]   = _parse_euro_float(df_raw[_SQ_QTY])
    out[COL_NET]   = (
        _parse_euro_currency(df_raw[_SQ_NET])
        if _SQ_NET in df_raw.columns else 0.0
    )
    out[COL_GROSS] = (
        _parse_euro_currency(df_raw[_SQ_GROSS])
        if _SQ_GROSS in df_raw.columns else 0.0
    )
    out[COL_OUTLET] = (
        df_raw[_SQ_OUTLET].apply(_shorten_outlet_name)
        if _SQ_OUTLET in df_raw.columns else "Unknown Outlet"
    )

    return out.reset_index(drop=True)


def get_outlets(df: pd.DataFrame) -> list[str]:
    if COL_OUTLET not in df.columns:
        return []
    return sorted(df[COL_OUTLET].dropna().astype(str).unique().tolist())


def get_categories_sq(
    df: pd.DataFrame,
    outlets: list[str] | None = None,
) -> list[str]:
    if COL_CATEGORY not in df.columns:
        return []
    d = df if not outlets else df[df[COL_OUTLET].isin(outlets)]
    return sorted(d[COL_CATEGORY].dropna().astype(str).unique().tolist())


def get_items_sq(
    df: pd.DataFrame,
    categories: list[str] | None = None,
    outlets: list[str] | None = None,
) -> list[str]:
    if COL_ITEM not in df.columns:
        return []
    d = df.copy()
    if outlets:
        d = d[d[COL_OUTLET].isin(outlets)]
    if categories:
        d = d[d[COL_CATEGORY].isin(categories)]
    return sorted(d[COL_ITEM].dropna().astype(str).unique().tolist())


def filter_sq(
    df: pd.DataFrame,
    start_date: date,
    end_date: date,
    outlets: list[str] | None = None,
    categories: list[str] | None = None,
    exclude_items: list[str] | None = None,
) -> pd.DataFrame:
    mask = (
        (df[COL_DATE].dt.date >= start_date) &
        (df[COL_DATE].dt.date <= end_date)
    )
    result = df.loc[mask].copy()
    if outlets:
        result = result[result[COL_OUTLET].isin(outlets)]
    if categories:
        result = result[result[COL_CATEGORY].isin(categories)]
    if exclude_items:
        result = result[~result[COL_ITEM].isin(exclude_items)]
    return result
