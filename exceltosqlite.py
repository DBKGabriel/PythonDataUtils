import os
import pathlib
import sqlite3
import tkinter as tk
from tkinter import filedialog
import argparse
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

    print(f"Loading {xlsx_path}")
    df = pd.read_excel(xlsx_path)
    print(f"    {len(df):,} rows read")

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

def parse_arguments():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(
        description="""Convert .xlsx  to .db 
        Specify multiple files with a space between each""",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python exceltosqlite.py                           # Use GUI to select files
  python exceltosqlite.py file1.xlsx file2.xlsx    # Convert specific files
  python exceltosqlite.py *.xlsx                   # Convert all xlsx files
        """
    )
        
    # Specify files to convert
    parser.add_argument(
        'files', 
        nargs='*', 
        help='Excel files to convert (.xlsx)'
    ) 
    
    return parser.parse_args()

if __name__ == "__main__":
    print("Running in:", os.getcwd())

    args = parse_arguments()
    
    # Get files from command line or GUI
    if args.files:
        # Files specified as positional arguments
        excel_files = args.files
        print(f"Processing {len(excel_files)} file(s) from command line...")
    else:
        # No files specified, use GUI
        print("Select files to convert.")
        excel_files = choose_excel_files()
        if not excel_files:
            print("No files selected. Why'd you run this code, silly?")
            quit()

    # Validate files exist
    valid_files = []
    for file_path in excel_files:
        path = pathlib.Path(file_path)
        if path.exists():
            valid_files.append(str(path))
        else:
            print(f"Sorry bud, couldn't find it: {file_path}")
    
    if not valid_files:
        print("What'd you do? None of those were nothing!")
        quit()

    # Convert all valid files
    for xlsx in valid_files:
        try:
            excel_to_sqlite(xlsx)
        except Exception as e:
            print(f"Ope {xlsx}: {e}")