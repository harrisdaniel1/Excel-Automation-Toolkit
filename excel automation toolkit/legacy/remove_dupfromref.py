"""
Remove Records Tool
-------------------
- Picks a SOURCE file and a REFERENCE file.
- Removes any row from the SOURCE whose old_company_number also exists in the REFERENCE.
- No merging. The cleaned SOURCE data is written to a NEW timestamped Excel file.
- Output folder is chosen via a dialog. File is saved as:
      Remove_Result_YYYYMMDD_HHMMSS.xlsx
- Output file has 3 sheets: Summary, Output, Removed_Records
"""

import pandas as pd
import os
from datetime import datetime
from tkinter import Tk, filedialog, messagebox

PRIMARY_KEY = "old_company_number"


def normalize_key(series):
    return (
        series.fillna("")
        .astype(str)
        .str.strip()
        .str.replace(".0", "", regex=False)
    )


root = Tk()
root.withdraw()

# 1. Pick SOURCE file (the file to clean)
source_file = filedialog.askopenfilename(
    title="Select SOURCE Excel File (to clean)",
    filetypes=[("Excel Files", "*.xlsx *.xls")],
)
if not source_file:
    raise SystemExit

print("Reading Source File...")
source_df = pd.read_excel(source_file, engine="openpyxl")
print(f"Source File Loaded: {len(source_df):,} rows")

# 2. Pick REFERENCE file (the removal list)
reference_file = filedialog.askopenfilename(
    title="Select REFERENCE Excel File (removal list)",
    filetypes=[("Excel Files", "*.xlsx *.xls")],
)
if not reference_file:
    raise SystemExit

print("Reading Reference File...")
reference_df = pd.read_excel(reference_file, engine="openpyxl")
print(f"Reference File Loaded: {len(reference_df):,} rows")

# 3. Pick OUTPUT folder
output_folder = filedialog.askdirectory(title="Select Output Folder")
if not output_folder:
    raise SystemExit

# 4. Validate primary key exists in both files
for df, f in [(source_df, source_file), (reference_df, reference_file)]:
    if PRIMARY_KEY not in df.columns:
        messagebox.showerror(
            "Missing Column",
            f"Column '{PRIMARY_KEY}' not found in {os.path.basename(f)}",
        )
        raise SystemExit

# 5. Normalize keys
source_df[PRIMARY_KEY] = normalize_key(source_df[PRIMARY_KEY])
reference_df[PRIMARY_KEY] = normalize_key(reference_df[PRIMARY_KEY])

# 6. Find SOURCE rows whose key exists in the REFERENCE -> these get removed
reference_keys = set(reference_df[PRIMARY_KEY])
remove_mask = source_df[PRIMARY_KEY].isin(reference_keys)

removed_records = source_df[remove_mask].copy()       # rows being deleted
cleaned_output = source_df[~remove_mask].copy()       # remaining source rows

# 7. Build Summary
ts = datetime.now()
summary = pd.DataFrame(
    {
        "Metric": [
            "Primary Key",
            "Source Rows",
            "Reference Rows",
            "Rows Removed",
            "Final Rows",
            "Timestamp",
        ],
        "Value": [
            PRIMARY_KEY,
            len(source_df),
            len(reference_df),
            len(removed_records),
            len(cleaned_output),
            ts.strftime("%Y-%m-%d %H:%M:%S"),
        ],
    }
)

# 8. Write NEW timestamped output file
outfile = os.path.join(
    output_folder, f"Remove_Result_{ts.strftime('%Y%m%d_%H%M%S')}.xlsx"
)

with pd.ExcelWriter(outfile, engine="openpyxl") as writer:
    summary.to_excel(writer, sheet_name="Summary", index=False)
    cleaned_output.to_excel(writer, sheet_name="Output", index=False)
    removed_records.to_excel(writer, sheet_name="Removed_Records", index=False)

messagebox.showinfo("Success", f"Removal completed!\n\nSaved to:\n{outfile}")
print("Done")
