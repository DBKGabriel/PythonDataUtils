import os
import pathlib
import sqlite3
import tkinter as tk
from tkinter import filedialog

import pandas as pd

def excel_to_sqlite(xlsx_path, db_path=None, table_name=None):
    """
    Converts the first Excel sheet of the selected workbook into a SQLite table.

    Parameters
    ----------
    xlsx_path : str | Path
        Path to the Excel workbook (.xlsx).
    db_path : str | Path, optional
        Destination .db file.  Defaults to <workbook-stem>.db
    table_name : str, optional
        Name of the table in SQLite.  Defaults to <workbook-stem>.
    """
    xlsx_path = pathlib.Path(xlsx_path)
    stem = xlsx_path.stem

    if db_path is None:
        db_path = xlsx_path.with_suffix(".db")
    if table_name is None:
        table_name = stem

    print(f"Loading {xlsx_path} …")
    df = pd.read_excel(xlsx_path)
    print(f"    → {len(df):,} rows read")

    conn = sqlite3.connect(db_path)
    df.to_sql(table_name, conn, if_exists='replace', index=False)
    conn.close()
    print(f"    Wrote table '{table_name}' to {db_path}")

def choose_excel_files():
    """Opens GUI and returns a list of selected .xlsx files."""
    root = tk.Tk()
    root.withdraw()
    paths = filedialog.askopenfilenames(
        title="Select Excel files to convert",
        filetypes=[("Excel workbooks", "*.xlsx")],
    )
    return list(paths)

if __name__ == "__main__":
    print("Running in:", os.getcwd())

    excel_files = choose_excel_files()
    if not excel_files:
        print("No files selected. Why'd you run this code, silly?")
        quit()

    for xlsx in excel_files:
        excel_to_sqlite(xlsx)

    print("All done!")
