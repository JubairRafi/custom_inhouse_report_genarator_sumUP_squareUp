"""
create_sample_data.py
=====================
Run this script once to generate two sample outlet Excel files for testing.

Usage
-----
    python create_sample_data.py

Outputs
-------
    sample_kings_road.xlsx
    sample_st_martin.xlsx
"""

import random
from datetime import date, timedelta

import pandas as pd

CATEGORIES = [
    "Pastries", "sandwich", "Bread", "Breakfast", "Cookies", "Tart",
    "Yogurt", "Orange Juice", "Ginger Shot", "Carrot Juice", "Salads",
    "Mini Market"
]

ITEMS = [
    "Chicken Burger", "Beef Burger", "Fish & Chips", "Caesar Salad",
    "Prawn Cocktail", "Steak Pie", "Vegetable Curry", "Lamb Chops",
    "Pasta Carbonara", "BBQ Ribs", "Spring Rolls", "Grilled Salmon",
    "Mushroom Risotto", "Club Sandwich", "Cheesecake Slice",
]

def _generate_sales(outlet_name: str, n_rows: int = 300) -> pd.DataFrame:
    """Generate random sales rows for a single outlet."""
    start = date(2025, 1, 1)
    rows = []
    for _ in range(n_rows):
        sale_date = start + timedelta(days=random.randint(0, 364))
        item = random.choice(ITEMS)
        qty = random.randint(1, 30)
        cat = random.choice(CATEGORIES)
        rows.append({"Date": sale_date, "Item Name": item, "Qty": qty, "Category": cat})
    return pd.DataFrame(rows)

if __name__ == "__main__":
    kings_df = _generate_sales("Kings Road")
    stmartin_df = _generate_sales("St Martin")

    kings_df.to_excel("sample_kings_road.xlsx", index=False)
    stmartin_df.to_excel("sample_st_martin.xlsx", index=False)

    print("[OK] Created sample_kings_road.xlsx and sample_st_martin.xlsx")
    print(f"   Kings Road rows  : {len(kings_df)}")
    print(f"   St Martin rows   : {len(stmartin_df)}")
