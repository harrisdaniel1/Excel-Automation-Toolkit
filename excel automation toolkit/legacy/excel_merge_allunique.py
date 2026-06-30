#script ni tak buang duplicate dekat source file
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

current_file = filedialog.askopenfilename(title="Select Current Excel File", filetypes=[("Excel Files","*.xlsx *.xls")])
if not current_file:
    raise SystemExit

print("Reading Current File...")
current_df = pd.read_excel(current_file, engine="openpyxl")
print(f"Current File Loaded: {len(current_df):,} rows")

reference_file = filedialog.askopenfilename(title="Select Reference Excel File", filetypes=[("Excel Files","*.xlsx *.xls")])
if not reference_file:
    raise SystemExit

print("Reading Reference File...")
reference_df = pd.read_excel(reference_file, engine="openpyxl")
print(f"Reference File Loaded: {len(reference_df):,} rows")

output_folder = filedialog.askdirectory(title="Select Output Folder")
if not output_folder:
    raise SystemExit

for df,f in [(current_df,current_file),(reference_df,reference_file)]:
    if PRIMARY_KEY not in df.columns:
        messagebox.showerror("Missing Column",
        f"Column '{PRIMARY_KEY}' not found in {os.path.basename(f)}")
        raise SystemExit

current_df[PRIMARY_KEY]=normalize_key(current_df[PRIMARY_KEY])
reference_df[PRIMARY_KEY]=normalize_key(reference_df[PRIMARY_KEY])

current_keys=set(current_df[PRIMARY_KEY])
dup_mask=reference_df[PRIMARY_KEY].isin(current_keys)

removed_records=reference_df[dup_mask].copy()
new_records=reference_df[~dup_mask].copy()

output_df=pd.concat([current_df,new_records],ignore_index=True)

ts=datetime.now()
summary=pd.DataFrame({
"Metric":["Primary Key","Current Rows","Reference Rows","Duplicates Ignored","New Rows Appended","Final Rows","Timestamp"],
"Value":[PRIMARY_KEY,len(current_df),len(reference_df),len(removed_records),len(new_records),len(output_df),ts.strftime("%Y-%m-%d %H:%M:%S")]
})

outfile=os.path.join(output_folder,f"Merge_Result_{ts.strftime('%Y%m%d_%H%M%S')}.xlsx")

with pd.ExcelWriter(outfile,engine="openpyxl") as writer:
    summary.to_excel(writer,sheet_name="Summary",index=False)
    output_df.to_excel(writer,sheet_name="Output",index=False)
    removed_records.to_excel(writer,sheet_name="Removed_Records",index=False)

messagebox.showinfo("Success",f"Merge completed!\n\nSaved to:\n{outfile}")
print("Done")
