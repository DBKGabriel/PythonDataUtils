import os
import pathlib
import sqlite3
import tkinter as tk
from tkinter import filedialog
import argparse
import pandas as pd
import datetime as dt
timestamp = dt.datetime.now()


def get_unique_db_path(base_path):
    """
    To avoid overwriting an existing file, this appends the date and time
    the conversion occured to the new .db file if the original filename 
    is already taken.
    e.g., sales.db -> sales_20250602_143022.db (if sales.db exists)
    """
    base_path = pathlib.Path(base_path)
    
    # If file doesn't exist, use original name
    if not base_path.exists():
        return base_path
    
    # File exists, create unique name with timestamp
    stem = base_path.stem
    suffix = base_path.suffix
    timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    
    unique_path = base_path.parent / f"{stem}_{timestamp}{suffix}"
    
    # Let's handle the wildest edgecase imagineable and add a counter if there's STILL a collision
    counter = 1
    while unique_path.exists():
        unique_path = base_path.parent / f"{stem}_{timestamp}_{counter}{suffix}"
        counter += 1
    
    return unique_path

def excel_to_sqlite(xlsx_path, db_path=None, table_name=None):
    """
    Converts the first Excel sheet of the selected workbook(s) into a SQLite table.

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
        proposed_db_path = xlsx_path.with_suffix(".db")
        db_path = get_unique_db_path(proposed_db_path)
        
        # Inform user if we had to change the name
        if db_path != proposed_db_path:
            print(f"Database {proposed_db_path.name} already exists!")
            print(f"Creating {db_path.name} instead...")
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
  python exceltosqlite.py                          # Use GUI to select files
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