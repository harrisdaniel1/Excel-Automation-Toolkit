# Code Walkthrough - Excel Splitter GUI

This document explains the concept and implementation of `excel_splitter_gui.py` block by block, making it easy to understand, maintain, or upload as documentation for a GitHub repository.

---

## 1. Imports and High DPI Support

```python
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
from datetime import datetime

# Enable High DPI awareness on Windows for crisp fonts and layout
try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass
```

### Concept:
- **`tkinter` / `ttk`**: Standard Python library for GUI components. We use `tk` for basic layout wrappers and `ttk` for styled native widgets (like `Progressbar`).
- **`pandas`**: Used to load, chunk, and export Excel data.
- **`ctypes.windll.shcore`**: By default, Windows scales Tkinter interfaces on high-DPI displays, which makes text and borders blurry. This `ctypes` block tells Windows to let the application handle its own DPI scaling, rendering fonts and lines sharp and clean.

---

## 2. Application Class and State Initialization

```python
class ExcelSplitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Splitter v1.0")
        self.root.geometry("680x580")
        self.root.resizable(False, False)
        self.root.configure(bg="#f1f5f9")  # Modern light slate-100 background

        # State variables
        self.selected_file = ""
        self.selected_output_folder = ""
        self.total_rows = 0
        self.df_cache = None

        # Build UI
        self.create_widgets()
```

### Concept:
- **Object-Oriented Structure**: Rather than using global variables (which makes the codebase fragile), all properties and UI states are tied to `self` inside a class structure.
- **State variables**:
  - `self.selected_file` & `self.selected_output_folder`: Paths selected by the user.
  - `self.total_rows`: Total rows found in the imported dataset.
  - `self.df_cache`: The loaded Pandas DataFrame, kept in memory to optimize reading times.
- **Window size**: Set to a fixed size of `680x580` pixels to fit all visual controls cleanly.

---

## 3. UI Layout - Header Banner

```python
        # 1. Header Banner
        header_frame = tk.Frame(self.root, bg="#1e1b4b")
        header_frame.pack(fill="x", side="top")
        
        title_label = tk.Label(
            header_frame,
            text="Excel Splitter",
            font=("Segoe UI", 16, "bold"),
            fg="#ffffff",
            bg="#1e1b4b"
        )
        title_label.pack(anchor="w", padx=24, pady=(15, 2))
        
        subtitle_label = tk.Label(
            header_frame,
            text="Split large Excel spreadsheets into smaller files in seconds",
            font=("Segoe UI", 9),
            fg="#c7d2fe",
            bg="#1e1b4b"
        )
        subtitle_label.pack(anchor="w", padx=24, pady=(0, 15))
```

### Concept:
- Creates a dark indigo (`#1e1b4b`) title block at the top of the window, providing a premium visual hierarchy.
- Uses `anchor="w"` to align the title and subtitle to the left.

---

## 4. UI Layout - The Settings Card

```python
        # 2. Main Container
        container = tk.Frame(self.root, bg="#f1f5f9")
        container.pack(fill="both", expand=True, padx=24, pady=(10, 15))

        # 3. Card Panel (slate-300 border, white background)
        card_border = tk.Frame(container, bg="#cbd5e1")
        card_border.pack(fill="x", pady=(0, 12))

        card = tk.Frame(card_border, bg="#ffffff")
        card.pack(fill="x", padx=1, pady=1)

        # Grid configuration for Card
        card.columnconfigure(0, minsize=110)
        card.columnconfigure(1, weight=1)
        card.columnconfigure(2, minsize=120)
```

### Concept:
- **Card Border Trick**: To create a thin, clean 1-pixel border, we wrap the white `card` inside a slightly larger `card_border` frame colored slate-300 (`#cbd5e1`) and set inner padding (`padx=1, pady=1`).
- **Grid Weights**: We use the `grid` layout manager inside the card. `columnconfigure(1, weight=1)` ensures that the middle column (holding the file and directory paths) stretches horizontally to take up all leftover space.

---

## 5. UI Layout - Rows inside the Settings Card

```python
        # Row 0: Input File
        lbl_input = tk.Label(card, text="Input File:", font=("Segoe UI", 9, "bold"), fg="#334155", bg="#ffffff")
        lbl_input.grid(row=0, column=0, sticky="w", padx=15, pady=(15, 8))

        input_display_border = tk.Frame(card, bg="#cbd5e1")
        input_display_border.grid(row=0, column=1, sticky="ew", padx=(0, 10), pady=(15, 8))

        self.input_file_label = tk.Label(input_display_border, text="No file selected", font=("Segoe UI", 9), fg="#94a3b8", bg="#f8fafc", anchor="w", height=2, padx=10)
        self.input_file_label.pack(fill="x", padx=1, pady=1)

        self.browse_input_btn = tk.Button(card, text="Browse File", command=self.browse_input_file, bg="#4f46e5", fg="#ffffff", activebackground="#4338ca", relief="flat", bd=0, font=("Segoe UI", 9, "bold"), width=12, height=1, cursor="hand2")
        self.browse_input_btn.grid(row=0, column=2, sticky="ew", padx=(0, 15), pady=(15, 8))
```

### Concept:
- **Input File row**: Contains a label, a custom path display box (modeled using our border trick with `#f8fafc` background), and a flat indigo "Browse File" button.
- **Output Folder row** and **Rows Per File row** follow an identical layout design using clean flat styling parameters like `relief="flat"`, `bd=0`, and `cursor="hand2"`.

---

## 6. UI Layout - Statistics Dashboard

```python
        # Row 3: Stats Box Container
        stats_frame = tk.Frame(card, bg="#ffffff")
        stats_frame.grid(row=3, column=0, columnspan=3, sticky="ew", padx=15, pady=(10, 15))
        stats_frame.columnconfigure(0, weight=1)
        stats_frame.columnconfigure(1, weight=1)

        # Stats Box 1 (Total Rows)
        box1_border = tk.Frame(stats_frame, bg="#e2e8f0")
        box1_border.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        box1 = tk.Frame(box1_border, bg="#f8fafc")
        box1.pack(fill="both", expand=True, padx=1, pady=1)

        lbl1 = tk.Label(box1, text="TOTAL DATA ROWS", font=("Segoe UI", 8, "bold"), fg="#64748b", bg="#f8fafc")
        lbl1.pack(pady=(10, 2), padx=15, anchor="w")

        self.total_rows_label = tk.Label(box1, text="0", font=("Segoe UI", 14, "bold"), fg="#0f172a", bg="#f8fafc")
        self.total_rows_label.pack(pady=(0, 10), padx=15, anchor="w")
```

### Concept:
- Displays summary statistics side-by-side using an embedded sub-grid (`columnconfigure(0)` and `columnconfigure(1)` with equal weight).
- Creates flat panel cards containing micro-labels (light gray, uppercase bold) and numeric value labels (large, dark bold).

---

## 7. Status, Progress Bar, and Action Buttons

```python
        # 4. Status Indicator (above progress bar)
        status_container = tk.Frame(container, bg="#f1f5f9")
        status_container.pack(fill="x", pady=(2, 2))
        ...
        # 5. Progress Bar
        self.progress = ttk.Progressbar(container, orient="horizontal", mode="determinate", style="Modern.Horizontal.TProgressbar")
        self.progress.pack(fill="x", pady=(0, 10))

        # 6. Action Buttons
        btn_frame = tk.Frame(container, bg="#f1f5f9")
        btn_frame.pack(pady=5)

        self.start_btn = tk.Button(btn_frame, text="Start Splitting", command=self.start_splitting, bg="#10b981", fg="#ffffff", activebackground="#059669", relief="flat", bd=0, font=("Segoe UI", 10, "bold"), width=22, height=2, cursor="hand2")
        self.start_btn.pack(side="left", padx=10)
```

### Concept:
- **Status Indicator**: Sits right above the progress bar to show text messages of what is currently happening.
- **Progress Bar**: Styled with `Modern.Horizontal.TProgressbar` to display a flat indigo bar on a white track.
- **Action Buttons**: Centered in `btn_frame` (accomplished by packing without `fill="x"`). Standardized to `width=22` and `height=2` to look clean, balanced, and identical in size.

---

## 8. Logic Method - Set Status

```python
    def set_status(self, text, status_type="info"):
        color_map = {
            "info": "#475569",     # Slate gray
            "success": "#10b981",  # Emerald green
            "warning": "#f59e0b",  # Amber/orange
            "error": "#ef4444",    # Rose/red
            "running": "#3b82f6"   # Blue
        }
        color = color_map.get(status_type, "#475569")
        self.status_label.config(text=text, fg=color)
        self.root.update_idletasks()
```

### Concept:
- Dynamically colors the status message depending on the context:
  - Information or Idle: Slate Gray.
  - Active operations (running): Blue.
  - Completed successfully: Green.
  - Warning or Exception: Orange / Red.
- `update_idletasks()` is called to instantly refresh the GUI window without waiting for the main loop cycle.

---

## 9. Logic Method - Browse and Load Input File

```python
    def browse_input_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if not file_path:
            return

        self.selected_file = file_path
        file_name = os.path.basename(file_path)
        self.input_file_label.config(text=file_name, fg="#0f172a")
        self.set_status("Reading file...", "running")

        try:
            # Load and cache dataframe
            self.df_cache = pd.read_excel(file_path, engine="openpyxl")
            self.total_rows = len(self.df_cache)
            self.total_rows_label.config(text=f"{self.total_rows:,}")
            self.update_expected_files()
            self.set_status("File loaded successfully", "success")
        except Exception as e:
            self.df_cache = None
            self.total_rows = 0
            self.total_rows_label.config(text="0")
            self.update_expected_files()
            self.set_status("Failed to read file", "error")
            messagebox.showerror("Error", f"Failed to read Excel file.\n\n{e}")
```

### Concept:
- **`filedialog.askopenfilename`**: Prompts the standard Windows open-file dialog.
- **In-Memory Caching**: We load the dataframe once here (`pd.read_excel`) and cache it in `self.df_cache`. This guarantees that when clicking "Start Splitting", we don't reload the file, optimizing execution speeds.
- **Comma Formatting**: Formatting using `{self.total_rows:,}` makes numbers like `10543` look like `10,543`.

---

## 10. Logic Method - Update Expected Files

```python
    def update_expected_files(self):
        try:
            val = self.rows_entry.get().strip()
            if not val:
                self.expected_files_label.config(text="-")
                return
            rows_per_file = int(val)
            if rows_per_file <= 0:
                self.expected_files_label.config(text="-")
                return

            expected_files = (self.total_rows + rows_per_file - 1) // rows_per_file
            self.expected_files_label.config(text=f"{expected_files:,}")
        except ValueError:
            self.expected_files_label.config(text="-")
```

### Concept:
- Calculates how many chunks the output will yield in real-time as the user types into the Entry box.
- **Formula**: `(total_rows + rows_per_file - 1) // rows_per_file` is an integer-based ceiling division. E.g., splitting 100 rows by 34 per file gives `(100 + 33) // 34 = 3` files.

---

## 11. Logic Method - Start Splitting and Exporting

```python
    def start_splitting(self):
        # Validation checks
        if not self.selected_file:
            messagebox.showwarning("Missing File", "Please select an Excel file.")
            return
        if not self.selected_output_folder:
            messagebox.showwarning("Missing Output Folder", "Please select an output folder.")
            return
        try:
            rows_per_file = int(self.rows_entry.get())
            if rows_per_file <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Invalid Value", "Rows per file must be greater than 0.")
            return

        self.toggle_ui_state("disabled")
        self.set_status("Processing...", "running")
        self.progress["value"] = 0
        self.root.update_idletasks()
```

### Concept:
- Assures all inputs are valid before proceeding.
- Calls `self.toggle_ui_state("disabled")` to disable browse buttons, entries, and action buttons. This prevents the user from clicking buttons again or changing values mid-execution.

```python
        try:
            if self.df_cache is None:
                self.df_cache = pd.read_excel(self.selected_file, engine="openpyxl")
                self.total_rows = len(self.df_cache)
                self.total_rows_label.config(text=f"{self.total_rows:,}")

            expected_files = (self.total_rows + rows_per_file - 1) // rows_per_file
            if expected_files == 0:
                raise ValueError("The selected file contains no data rows to split.")

            # Create timestamp folder
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_folder = os.path.join(self.selected_output_folder, f"Criteria1_{timestamp}")
            os.makedirs(output_folder, exist_ok=True)

            for i in range(0, self.total_rows, rows_per_file):
                chunk = self.df_cache.iloc[i : i + rows_per_file]
                part_num = (i // rows_per_file) + 1
                output_path = os.path.join(output_folder, f"Criteria1_part_{part_num}.xlsx")

                # Export chunk to new Excel file
                chunk.to_excel(output_path, index=False)

                # Progress percentage calculation
                pct = (part_num / expected_files) * 100
                self.progress["value"] = pct
                self.set_status(f"Exported part {part_num} of {expected_files}...", "running")

            self.set_status("Completed", "success")
            messagebox.showinfo("Success", f"{expected_files} files successfully created!\n\nSaved in:\n{output_folder}")
```

### Concept:
- **Timestamp Subfolder**: Automatically builds a subdirectory named `Criteria1_YYYYMMDD_HHMMSS` inside the selected output directory to organize output chunks cleanly.
- **DataFrame Chunking**: `self.df_cache.iloc[i : i + rows_per_file]` slices a chunk of row indices from index `i` to `i + rows_per_file`.
- **`to_excel(..., index=False)`**: Saves the sliced chunk to a separate file, hiding the Pandas index column.
- **Real-time Progress updates**: Loops update `self.progress["value"]` with the percentage of progress and triggers `self.root.update_idletasks()` to update the bar animation in real-time.

```python
        except Exception as e:
            self.set_status("Error during splitting", "error")
            messagebox.showerror("Error", f"An error occurred while splitting the file:\n\n{e}")
        finally:
            self.toggle_ui_state("normal")
```

### Concept:
- If a permission error occurs (e.g. file locked) or output directory is write-protected, it catches the error safely, displays a message box, and guarantees that buttons are re-enabled in the `finally` block.

---

## 12. Main Initialization Block

```python
if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelSplitterApp(root)
    root.mainloop()
```

### Concept:
- standard entry point. Checks if the file is run directly (rather than imported), instantiates the Tkinter root window, launches our custom class, and triggers `mainloop()` to run the GUI interface.
