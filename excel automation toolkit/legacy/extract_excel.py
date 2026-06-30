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
    title="Select Excel File for Extraction",
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
try:
    df = pd.read_excel(input_file, engine="openpyxl")
except Exception as e:
    print(f"Error reading Excel file: {e}")
    exit()

total_rows = len(df)

if total_rows == 0:
    print("Selected file contains no data.")
    exit()

print(f"Total Rows: {total_rows}")

# =========================
# GET EXTRACTION PERCENTAGE
# =========================
while True:
    try:
        pct_str = input("Enter percentage to extract (0.01 to 100.0): ").strip()
        percentage = float(pct_str)

        if not (0.0 < percentage <= 100.0):
            print("Percentage must be greater than 0% and up to 100%.")
            continue

        break

    except ValueError:
        print("Please enter a valid decimal number.")

# =========================
# GET RANDOM SEED
# =========================
while True:
    seed_str = input("Enter random seed (or press Enter for default 42): ").strip()
    if not seed_str:
        seed = 42
        break
    try:
        seed = int(seed_str)
        break
    except ValueError:
        print("Seed must be a valid integer.")

# =========================
# TRACEABILITY OPTION
# =========================
add_index_raw = input("Add original Excel row index column? (Y/N, default Y): ").strip().upper()
add_index = (add_index_raw != "N")

# =========================
# CALCULATE TARGET ROWS
# =========================
target_rows = max(1, round(total_rows * (percentage / 100.0)))
target_rows = min(total_rows, target_rows)

# =========================
# EXECUTE RANDOM EXTRACTION
# =========================
print(f"\nExtracting {target_rows} rows ({percentage}%) using Random Sampling (seed={seed})...")

extracted_df = df.sample(n=target_rows, random_state=seed).sort_index()

if add_index:
    extracted_df.insert(0, "Original_Excel_Row", extracted_df.index + 2)

# =========================
# CONSTRUCT OUTPUT PATH
# =========================
base_name = os.path.splitext(os.path.basename(input_file))[0]
timestamp  = datetime.now().strftime("%Y%m%d_%H%M%S")
pct_slug   = f"{percentage:g}".replace(".", "_")

output_filename = f"{base_name}_extract_{pct_slug}pct_random_{timestamp}.xlsx"
output_path     = os.path.join(output_location, output_filename)

# =========================
# SUMMARY & CONFIRM
# =========================
print("\n========== SUMMARY ==========")
print(f"Input File     : {input_file}")
print(f"Output File    : {output_path}")
print(f"Total Rows     : {total_rows}")
print(f"Percentage     : {percentage}%")
print(f"Target Rows    : {target_rows}")
print(f"Method         : Random Sampling")
print(f"Random Seed    : {seed}")
print(f"Add Excel Row  : {add_index}")
print("=============================")

confirm = input("Proceed? (Y/N): ").strip().upper()

if confirm != "Y":
    print("Operation cancelled.")
    exit()

print("Writing Excel file...")
try:
    extracted_df.to_excel(output_path, index=False)
    print("\nSuccess!")
    print(f"Extracted file saved to: {output_path}")
except Exception as e:
    print(f"Error writing Excel file: {e}")
