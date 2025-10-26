#!/usr/bin/env python3
"""
end_to_end_excel_demo.py
1) Create a sample Excel file
2) Read/import it
3) Manipulate the data
4) Export results to another Excel file
"""

import pandas as pd
from pathlib import Path
from datetime import datetime

def create_sample_excel(path: str) -> None:
    """Create a small, realistic Excel file you can process."""
    data = {
        "Date": [
            datetime(2025, 1, 5),
            datetime(2025, 1, 10),
            datetime(2025, 2, 3),
            datetime(2025, 2, 15),
            datetime(2025, 3, 7),
            datetime(2025, 3, 20),
        ],
        "Category": ["Food", "Transport", "Food", "Rent", "Transport", "Food"],
        "Amount": [50, 20, 30, 500, 25, 45],
        "Notes": ["Groceries", "Bus fare", "Lunch", "Monthly rent", "Taxi", "Dinner"],
    }
    df = pd.DataFrame(data)
    df.to_excel(path, index=False, engine="openpyxl")
    print(f"✅ Created sample Excel: {path}")

def load_excel_first_sheet(path: str) -> pd.DataFrame:
    """Import Excel (first sheet) as DataFrame and normalize headers."""
    df = pd.read_excel(path, sheet_name=0, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    return df

def manipulate(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame | None]:
    """
    Do a few realistic manipulations:
      - Clean rows
      - Ensure numeric Amount
      - Add VAT column
      - Create YearMonth
      - Filter positive amounts
      - Group totals by Category (+ by month if Date exists)
      - Optional pivot Category x Month
    Returns (df_clean, df_pos, totals, pivot_or_None)
    """
    # Drop fully empty rows
    df = df.dropna(how="all").copy()

    # Ensure Amount exists and is numeric
    if "Amount" not in df.columns:
        df["Amount"] = 1.0
    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0.0)

    # Parse Date if present and build YearMonth
    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        df["YearMonth"] = df["Date"].dt.to_period("M").astype(str)

    # Add a derived column (example: VAT at 20%)
    df["VAT_20"] = (df["Amount"] * 0.20).round(2)

    # Filter positive amounts only
    df_pos = df[df["Amount"] > 0].copy()

    # Group totals by Category (and by month if YearMonth exists)
    if "Category" not in df.columns:
        df["Category"] = "All"
        df_pos["Category"] = "All"

    if "YearMonth" in df.columns:
        totals = (df_pos.groupby(["Category", "YearMonth"], as_index=False)["Amount"]
                        .sum()
                        .sort_values(["Category", "YearMonth"]))
        pivot = df_pos.pivot_table(index="Category",
                                   columns="YearMonth",
                                   values="Amount",
                                   aggfunc="sum",
                                   fill_value=0).sort_index()
    else:
        totals = (df_pos.groupby(["Category"], as_index=False)["Amount"]
                        .sum()
                        .sort_values("Category"))
        pivot = None

    return df, df_pos, totals, pivot

def export_results(input_path: str,
                   df_clean: pd.DataFrame,
                   df_pos: pd.DataFrame,
                   totals: pd.DataFrame,
                   pivot: pd.DataFrame | None) -> str:
    """Save outputs next to the input as *_OUT.xlsx."""
    out_path = Path(input_path).with_suffix("").as_posix() + "_OUT.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as xlw:
        df_clean.to_excel(xlw, sheet_name="Cleaned", index=False)
        df_pos.to_excel(xlw, sheet_name="PositiveOnly", index=False)
        totals.to_excel(xlw, sheet_name="Totals", index=False)
        if pivot is not None:
            pivot.to_excel(xlw, sheet_name="Pivot")
    print(f"✅ Wrote results to: {out_path}")
    return out_path

def main():
    sample_file = "Input.xlsx"

    # 1) Create Excel
    create_sample_excel(sample_file)

    # 2) Import it
    df = load_excel_first_sheet(sample_file)
    print("Loaded columns:", df.columns.tolist())

    # 3) Manipulate
    df_clean, df_pos, totals, pivot = manipulate(df)

    # 4) Export
    export_results(sample_file, df_clean, df_pos, totals, pivot)

if __name__ == "__main__":
    main()