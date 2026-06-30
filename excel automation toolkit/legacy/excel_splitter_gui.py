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

    def create_widgets(self):
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

        # Row 0: Input File
        lbl_input = tk.Label(
            card,
            text="Input File:",
            font=("Segoe UI", 9, "bold"),
            fg="#334155",
            bg="#ffffff"
        )
        lbl_input.grid(row=0, column=0, sticky="w", padx=15, pady=(15, 8))

        input_display_border = tk.Frame(card, bg="#cbd5e1")
        input_display_border.grid(row=0, column=1, sticky="ew", padx=(0, 10), pady=(15, 8))

        self.input_file_label = tk.Label(
            input_display_border,
            text="No file selected",
            font=("Segoe UI", 9),
            fg="#94a3b8",
            bg="#f8fafc",
            anchor="w",
            height=2,
            padx=10
        )
        self.input_file_label.pack(fill="x", padx=1, pady=1)

        self.browse_input_btn = tk.Button(
            card,
            text="Browse File",
            command=self.browse_input_file,
            bg="#4f46e5",
            fg="#ffffff",
            activebackground="#4338ca",
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            font=("Segoe UI", 9, "bold"),
            width=12,
            height=1,
            cursor="hand2"
        )
        self.browse_input_btn.grid(row=0, column=2, sticky="ew", padx=(0, 15), pady=(15, 8))

        # Row 1: Output Folder
        lbl_output = tk.Label(
            card,
            text="Output Folder:",
            font=("Segoe UI", 9, "bold"),
            fg="#334155",
            bg="#ffffff"
        )
        lbl_output.grid(row=1, column=0, sticky="w", padx=15, pady=8)

        output_display_border = tk.Frame(card, bg="#cbd5e1")
        output_display_border.grid(row=1, column=1, sticky="ew", padx=(0, 10), pady=8)

        self.output_folder_label = tk.Label(
            output_display_border,
            text="No folder selected",
            font=("Segoe UI", 9),
            fg="#94a3b8",
            bg="#f8fafc",
            anchor="w",
            height=2,
            padx=10
        )
        self.output_folder_label.pack(fill="x", padx=1, pady=1)

        self.browse_output_btn = tk.Button(
            card,
            text="Browse Folder",
            command=self.browse_output_folder,
            bg="#4f46e5",
            fg="#ffffff",
            activebackground="#4338ca",
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            font=("Segoe UI", 9, "bold"),
            width=12,
            height=1,
            cursor="hand2"
        )
        self.browse_output_btn.grid(row=1, column=2, sticky="ew", padx=(0, 15), pady=8)

        # Row 2: Rows Per File
        lbl_rows = tk.Label(
            card,
            text="Rows Per File:",
            font=("Segoe UI", 9, "bold"),
            fg="#334155",
            bg="#ffffff"
        )
        lbl_rows.grid(row=2, column=0, sticky="w", padx=15, pady=8)

        entry_border = tk.Frame(card, bg="#cbd5e1")
        entry_border.grid(row=2, column=1, sticky="w", padx=(0, 10), pady=8)

        self.rows_entry = tk.Entry(
            entry_border,
            font=("Segoe UI", 10),
            fg="#1e293b",
            bg="#ffffff",
            relief="flat",
            bd=0,
            width=12
        )
        self.rows_entry.pack(padx=5, pady=5)
        self.rows_entry.insert(0, "34")
        self.rows_entry.bind("<KeyRelease>", lambda e: self.update_expected_files())

        lbl_unit = tk.Label(
            card,
            text="rows / file",
            font=("Segoe UI", 9, "italic"),
            fg="#64748b",
            bg="#ffffff"
        )
        lbl_unit.grid(row=2, column=2, sticky="w", padx=(0, 15), pady=8)

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

        # Stats Box 2 (Expected Files)
        box2_border = tk.Frame(stats_frame, bg="#e2e8f0")
        box2_border.grid(row=0, column=1, sticky="ew", padx=(8, 0))

        box2 = tk.Frame(box2_border, bg="#f8fafc")
        box2.pack(fill="both", expand=True, padx=1, pady=1)

        lbl2 = tk.Label(box2, text="EXPECTED FILES", font=("Segoe UI", 8, "bold"), fg="#64748b", bg="#f8fafc")
        lbl2.pack(pady=(10, 2), padx=15, anchor="w")

        self.expected_files_label = tk.Label(box2, text="0", font=("Segoe UI", 14, "bold"), fg="#0f172a", bg="#f8fafc")
        self.expected_files_label.pack(pady=(0, 10), padx=15, anchor="w")

        # 4. Status Indicator (above progress bar)
        status_container = tk.Frame(container, bg="#f1f5f9")
        status_container.pack(fill="x", pady=(2, 2))

        lbl_status_title = tk.Label(
            status_container,
            text="System Status:",
            font=("Segoe UI", 9, "bold"),
            fg="#475569",
            bg="#f1f5f9"
        )
        lbl_status_title.pack(side="left")

        self.status_label = tk.Label(
            status_container,
            text="Ready",
            font=("Segoe UI", 9, "bold"),
            fg="#475569",
            bg="#f1f5f9"
        )
        self.status_label.pack(side="left", padx=5)

        # 5. Progress Bar
        self.progress = ttk.Progressbar(
            container,
            orient="horizontal",
            mode="determinate",
            style="Modern.Horizontal.TProgressbar"
        )
        self.progress.pack(fill="x", pady=(0, 10))

        # 6. Action Buttons
        btn_frame = tk.Frame(container, bg="#f1f5f9")
        btn_frame.pack(pady=5)

        self.start_btn = tk.Button(
            btn_frame,
            text="Start Splitting",
            command=self.start_splitting,
            bg="#10b981",
            fg="#ffffff",
            activebackground="#059669",
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            font=("Segoe UI", 10, "bold"),
            width=22,
            height=2,
            cursor="hand2"
        )
        self.start_btn.pack(side="left", padx=10)

        self.reset_btn = tk.Button(
            btn_frame,
            text="Reset",
            command=self.reset_fields,
            bg="#cbd5e1",
            fg="#334155",
            activebackground="#94a3b8",
            activeforeground="#334155",
            relief="flat",
            bd=0,
            font=("Segoe UI", 10, "bold"),
            width=22,
            height=2,
            cursor="hand2"
        )
        self.reset_btn.pack(side="left", padx=10)

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
            self.df_cache = pd.read_excel(
                file_path,
                engine="openpyxl"
            )
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
            messagebox.showerror(
                "Error",
                f"Failed to read Excel file.\n\n{e}"
            )

    def browse_output_folder(self):
        folder_path = filedialog.askdirectory(
            title="Select Output Folder"
        )

        if not folder_path:
            return

        self.selected_output_folder = folder_path
        folder_name = os.path.basename(folder_path)
        if not folder_name:
            folder_name = folder_path

        self.output_folder_label.config(text=folder_name, fg="#0f172a")
        self.set_status("Output folder selected", "info")

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

    def start_splitting(self):
        # Validation
        if not self.selected_file:
            messagebox.showwarning(
                "Missing File",
                "Please select an Excel file."
            )
            return

        if not self.selected_output_folder:
            messagebox.showwarning(
                "Missing Output Folder",
                "Please select an output folder."
            )
            return

        try:
            rows_per_file = int(self.rows_entry.get())
            if rows_per_file <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning(
                "Invalid Value",
                "Rows per file must be greater than 0."
            )
            return

        # Disable UI controls during splitting
        self.toggle_ui_state("disabled")
        self.set_status("Processing...", "running")
        self.progress["value"] = 0
        self.root.update_idletasks()

        try:
            # Ensure we have data loaded
            if self.df_cache is None:
                self.df_cache = pd.read_excel(self.selected_file, engine="openpyxl")
                self.total_rows = len(self.df_cache)
                self.total_rows_label.config(text=f"{self.total_rows:,}")

            expected_files = (self.total_rows + rows_per_file - 1) // rows_per_file
            if expected_files == 0:
                raise ValueError("The selected file contains no data rows to split.")

            # Create timestamp folder
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_folder = os.path.join(
                self.selected_output_folder,
                f"Criteria1_{timestamp}"
            )
            os.makedirs(output_folder, exist_ok=True)

            for i in range(0, self.total_rows, rows_per_file):
                chunk = self.df_cache.iloc[i : i + rows_per_file]
                part_num = (i // rows_per_file) + 1
                output_path = os.path.join(
                    output_folder,
                    f"Criteria1_part_{part_num}.xlsx"
                )

                chunk.to_excel(
                    output_path,
                    index=False
                )

                # Progress percentage
                pct = (part_num / expected_files) * 100
                self.progress["value"] = pct
                self.set_status(f"Exported part {part_num} of {expected_files}...", "running")

            self.set_status("Completed", "success")
            messagebox.showinfo(
                "Success",
                f"{expected_files} files successfully created!\n\nSaved in:\n{output_folder}"
            )

        except Exception as e:
            self.set_status("Error during splitting", "error")
            messagebox.showerror(
                "Error",
                f"An error occurred while splitting the file:\n\n{e}"
            )
        finally:
            self.toggle_ui_state("normal")

    def toggle_ui_state(self, state):
        self.start_btn.config(state=state)
        self.reset_btn.config(state=state)
        self.browse_input_btn.config(state=state)
        self.browse_output_btn.config(state=state)
        self.rows_entry.config(state=state)

    def reset_fields(self):
        self.selected_file = ""
        self.selected_output_folder = ""
        self.total_rows = 0
        self.df_cache = None

        self.input_file_label.config(text="No file selected", fg="#94a3b8")
        self.output_folder_label.config(text="No folder selected", fg="#94a3b8")

        self.rows_entry.config(state="normal")
        self.rows_entry.delete(0, tk.END)
        self.rows_entry.insert(0, "34")

        self.total_rows_label.config(text="0")
        self.expected_files_label.config(text="0")
        self.set_status("Ready", "info")
        self.progress["value"] = 0

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelSplitterApp(root)
    root.mainloop()