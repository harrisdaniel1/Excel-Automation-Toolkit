import pandas as pd
import os
from datetime import datetime
from tkinter import Tk, filedialog

# Hide Tkinter root window
root = Tk()
root.withdraw()

# =========================
# SELECT INPUT FILE
# =========================
input_file = filedialog.askopenfilename(
    title="Select Excel File",
    filetypes=[("Excel Files", "*.xlsx *.xls")]
)

if not input_file:
    print("No file selected.")
    exit()

# =========================
# SELECT OUTPUT LOCATION
# =========================
output_location = filedialog.askdirectory(
    title="Select Output Location"
)

if not output_location:
    print("No output location selected.")
    exit()

# =========================
# READ EXCEL
# =========================
print("Reading Excel file...")

df = pd.read_excel(input_file, engine="openpyxl")

total_rows = len(df)

if total_rows == 0:
    print("Selected file contains no data.")
    exit()

print(f"Total Rows: {total_rows}")

# =========================
# GET ROWS PER FILE
# =========================
while True:
    try:
        rows_per_file = int(input("Enter rows per file: "))

        if rows_per_file <= 0:
            print("Rows per file must be greater than 0.")
            continue

        break

    except ValueError:
        print("Please enter a valid number.")

# =========================
# CALCULATE EXPECTED FILES
# =========================
expected_files = (total_rows + rows_per_file - 1) // rows_per_file

# =========================
# CREATE TIMESTAMP FOLDER
# =========================
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

output_folder = os.path.join(
    output_location,
    f"Criteria1_{timestamp}"
)

os.makedirs(output_folder, exist_ok=True)

# =========================
# SUMMARY
# =========================
print("\n========== SUMMARY ==========")
print(f"Input File     : {input_file}")
print(f"Output Folder  : {output_folder}")
print(f"Total Rows     : {total_rows}")
print(f"Rows Per File  : {rows_per_file}")
print(f"Expected Files : {expected_files}")
print("=============================")

confirm = input("Proceed? (Y/N): ").strip().upper()

if confirm != "Y":
    print("Operation cancelled.")
    exit()

# =========================
# SPLIT FILE
# =========================
for i in range(0, total_rows, rows_per_file):

    chunk = df.iloc[i:i + rows_per_file]

    part_num = (i // rows_per_file) + 1

    output_path = os.path.join(
        output_folder,
        f"Criteria1_part_{part_num}.xlsx"
    )

    chunk.to_excel(
        output_path,
        index=False
    )

    if part_num % 50 == 0:
        print(f"Exported {part_num} files...")

print("\nSuccess!")
print(f"Files saved in: {output_folder}")