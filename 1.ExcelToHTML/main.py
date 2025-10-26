#!/usr/bin/env python3
"""
excel_to_html.py — Read an Excel file and output HTML.

Usage examples:
    python excel_to_html.py data.xlsx
    python excel_to_html.py data.xlsx --sheet "January"
    python excel_to_html.py data.xlsx --sheet all --out out.html
"""

import argparse
import sys
import pandas as pd

def df_to_html(df: pd.DataFrame, table_id: str = "table1") -> str:
    # Basic, clean HTML table without index
    html_table = df.to_html(index=False, border=0, classes="tbl", table_id=table_id)
    # Wrap with minimal CSS for readability
    html = f"""<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Excel to HTML</title>
<style>
    body {{ font-family: Arial, Helvetica, sans-serif; margin: 24px; }}
    h2 {{ margin-top: 32px; }}
    table.tbl {{ border-collapse: collapse; width: 100%; }}
    table.tbl th, table.tbl td {{ border: 1px solid #ddd; padding: 8px; }}
    table.tbl tr:nth-child(even) {{ background: #f9f9f9; }}
    table.tbl th {{ background: #f1f1f1; text-align: left; }}
</style>
</head>
<body>
{html_table}
</body>
</html>"""
    return html

def dfs_to_html(maps: dict) -> str:
    # Multiple sheets: stack them with headings
    parts = []
    for i, (sheet, df) in enumerate(maps.items(), start=1):
        table = df.to_html(index=False, border=0, classes="tbl", table_id=f"tbl_{i}")
        parts.append(f"<h2>{sheet}</h2>\n{table}")
    return f"""<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Excel to HTML</title>
<style>
    body {{ font-family: Arial, Helvetica, sans-serif; margin: 24px; }}
    h2 {{ margin-top: 32px; }}
    table.tbl {{ border-collapse: collapse; width: 100%; }}
    table.tbl th, table.tbl td {{ border: 1px solid #ddd; padding: 8px; }}
    table.tbl tr:nth-child(even) {{ background: #f9f9f9; }}
    table.tbl th {{ background: #f1f1f1; text-align: left; }}
</style>
</head>
<body>
{''.join(parts)}
</body>
</html>"""

def main():
    parser = argparse.ArgumentParser(description="Convert Excel to HTML.")
    parser.add_argument("excel", help="Path to the .xlsx file")
    parser.add_argument("--sheet", default=None,
                        help='Sheet name to load (e.g., "January"). Use "all" for all sheets. Default: first sheet.')
    parser.add_argument("--out", default=None, help="Optional output HTML file. If omitted, prints to stdout.")
    parser.add_argument("--na", default="", help='NA/NaN representation in output (default: empty string).')
    args = parser.parse_args()

    # Read Excel
    if args.sheet and args.sheet.lower() == "all":
        dfs = pd.read_excel(args.excel, sheet_name=None, dtype=str)
        # Normalize NA
        for k in dfs:
            dfs[k] = dfs[k].fillna(args.na)
        html = dfs_to_html(dfs)
    else:
        # Single sheet (name or first by default)
        df = pd.read_excel(args.excel, sheet_name=args.sheet, dtype=str)
        df = df.fillna(args.na)
        html = df_to_html(df, table_id="table1")

    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"Written HTML to: {args.out}")
    else:
        # Print to stdout
        sys.stdout.write(html)

if __name__ == "__main__":
    main()
