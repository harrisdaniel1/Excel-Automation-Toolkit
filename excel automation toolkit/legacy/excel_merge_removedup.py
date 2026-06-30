import pandas as pd
import os
from datetime import datetime
from tkinter import Tk, filedialog, messagebox

# =====================================================
# CONFIGURATION
# CHANGE THIS COLUMN IF REQUIRED
# =====================================================
PRIMARY_KEY = "old_company_number"


def normalize_key(series):
    return (
        series.fillna("")
        .astype(str)
        .str.strip()
        .str.replace(".0", "", regex=False)
    )


# =====================================================
# FILE PICKER
# =====================================================
root = Tk()
root.withdraw()

print("Select Source File...")
source_file = filedialog.askopenfilename(
    title="Select Source File",
    filetypes=[("Excel Files", "*.xlsx *.xls")]
)

if not source_file:
    exit()

print("Reading Source File...")
source_df = pd.read_excel(source_file, engine="openpyxl")
print(f"Source Rows Loaded : {len(source_df):,}")

print()

print("Select Reference File...")
reference_file = filedialog.askopenfilename(
    title="Select Reference File",
    filetypes=[("Excel Files", "*.xlsx *.xls")]
)

if not reference_file:
    exit()

print("Reading Reference File...")
reference_df = pd.read_excel(reference_file, engine="openpyxl")
print(f"Reference Rows Loaded : {len(reference_df):,}")

print()

output_folder = filedialog.askdirectory(
    title="Select Output Folder"
)

if not output_folder:
    exit()

# =====================================================
# VALIDATION
# =====================================================
for df, filename in [
    (source_df, source_file),
    (reference_df, reference_file)
]:

    if PRIMARY_KEY not in df.columns:

        messagebox.showerror(
            "Missing Column",
            f"Column '{PRIMARY_KEY}' not found.\n\n"
            f"File:\n{os.path.basename(filename)}"
        )

        exit()

# =====================================================
# NORMALIZE
# =====================================================
source_df[PRIMARY_KEY] = normalize_key(
    source_df[PRIMARY_KEY]
)

reference_df[PRIMARY_KEY] = normalize_key(
    reference_df[PRIMARY_KEY]
)

# =====================================================
# FIND DUPLICATES
# =====================================================
source_keys = set(source_df[PRIMARY_KEY])
reference_keys = set(reference_df[PRIMARY_KEY])

duplicate_keys = source_keys.intersection(reference_keys)

print(f"Duplicate Keys Found : {len(duplicate_keys):,}")

# =====================================================
# REMOVE DUPLICATES FROM BOTH FILES
# =====================================================
source_unique = source_df[
    ~source_df[PRIMARY_KEY].isin(
        duplicate_keys
    )
].copy()

reference_unique = reference_df[
    ~reference_df[PRIMARY_KEY].isin(
        duplicate_keys
    )
].copy()

# =====================================================
# REMOVED RECORDS
# =====================================================
removed_source = source_df[
    source_df[PRIMARY_KEY].isin(
        duplicate_keys
    )
]

removed_reference = reference_df[
    reference_df[PRIMARY_KEY].isin(
        duplicate_keys
    )
]

removed_records = pd.concat(
    [
        removed_source,
        removed_reference
    ],
    ignore_index=True
)

# =====================================================
# OUTPUT
# =====================================================
output_df = pd.concat(
    [
        source_unique,
        reference_unique
    ],
    ignore_index=True
)

# =====================================================
# SUMMARY
# =====================================================
timestamp = datetime.now()

summary = pd.DataFrame({

    "Metric": [

        "Primary Key",

        "Source Rows",

        "Reference Rows",

        "Duplicate Keys Found",

        "Unique Source Rows",

        "Unique Reference Rows",

        "Final Output Rows",

        "Process Timestamp"

    ],

    "Value": [

        PRIMARY_KEY,

        len(source_df),

        len(reference_df),

        len(duplicate_keys),

        len(source_unique),

        len(reference_unique),

        len(output_df),

        timestamp.strftime(
            "%Y-%m-%d %H:%M:%S"
        )

    ]

})

# =====================================================
# SAVE
# =====================================================
outfile = os.path.join(

    output_folder,

    f"New_Record_Result_{timestamp.strftime('%Y%m%d_%H%M%S')}.xlsx"

)

with pd.ExcelWriter(
    outfile,
    engine="openpyxl"
) as writer:

    summary.to_excel(
        writer,
        sheet_name="Summary",
        index=False
    )

    output_df.to_excel(
        writer,
        sheet_name="Output",
        index=False
    )

    removed_records.to_excel(
        writer,
        sheet_name="Removed_Records",
        index=False
    )

print()
print("==============================")
print("PROCESS COMPLETED")
print("==============================")
print(f"Source Rows          : {len(source_df):,}")
print(f"Reference Rows       : {len(reference_df):,}")
print(f"Duplicate Keys       : {len(duplicate_keys):,}")
print(f"Output Rows          : {len(output_df):,}")
print("==============================")

messagebox.showinfo(
    "Success",
    f"Process completed successfully!\n\n"
    f"Output saved to:\n{outfile}"
)