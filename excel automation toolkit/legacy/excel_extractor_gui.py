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

class ExcelExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Row Extractor v1.0")
        self.root.geometry("680x620")
        self.root.resizable(False, False)
        self.root.configure(bg="#f1f5f9")  # Modern light slate-100 background

        # State variables
        self.selected_file = ""
        self.selected_output_folder = ""
        self.total_rows = 0
        self.df_cache = None

        # Configure style
        self.configure_styles()

        # Build UI
        self.create_widgets()

    def configure_styles(self):
        style = ttk.Style(self.root)
        # Use clam theme as base to allow progressbar color overrides
        style.theme_use('clam')
        style.configure(
            "Modern.Horizontal.TProgressbar",
            troughcolor="#e2e8f0",
            background="#3b82f6",
            thickness=12,
            borderwidth=0
        )

    def create_widgets(self):
        # 1. Header Banner
        header_frame = tk.Frame(self.root, bg="#1e1b4b")
        header_frame.pack(fill="x", side="top")
        
        title_label = tk.Label(
            header_frame,
            text="Excel Row Extractor",
            font=("Segoe UI", 16, "bold"),
            fg="#ffffff",
            bg="#1e1b4b"
        )
        title_label.pack(anchor="w", padx=24, pady=(15, 2))
        
        subtitle_label = tk.Label(
            header_frame,
            text="Extract representative subsets of Excel sheets by percentage",
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
        card.columnconfigure(0, minsize=140)
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

        # Row 2: Extraction Percentage
        lbl_pct = tk.Label(
            card,
            text="Percentage:",
            font=("Segoe UI", 9, "bold"),
            fg="#334155",
            bg="#ffffff"
        )
        lbl_pct.grid(row=2, column=0, sticky="w", padx=15, pady=8)

        pct_entry_border = tk.Frame(card, bg="#cbd5e1")
        pct_entry_border.grid(row=2, column=1, sticky="w", padx=(0, 10), pady=8)

        self.pct_entry = tk.Entry(
            pct_entry_border,
            font=("Segoe UI", 10),
            fg="#1e293b",
            bg="#ffffff",
            relief="flat",
            bd=0,
            width=12
        )
        self.pct_entry.pack(padx=5, pady=5)
        self.pct_entry.insert(0, "1.0")
        self.pct_entry.bind("<KeyRelease>", lambda e: self.update_target_rows())

        lbl_pct_unit = tk.Label(
            card,
            text="% of rows (0.01 - 100)",
            font=("Segoe UI", 9, "italic"),
            fg="#64748b",
            bg="#ffffff"
        )
        lbl_pct_unit.grid(row=2, column=2, sticky="w", padx=(0, 15), pady=8)

        # Row 3: Sampling Method
        lbl_method = tk.Label(
            card,
            text="Sampling Method:",
            font=("Segoe UI", 9, "bold"),
            fg="#334155",
            bg="#ffffff"
        )
        lbl_method.grid(row=3, column=0, sticky="w", padx=15, pady=8)

        self.method_combo = ttk.Combobox(
            card,
            values=["Random Sampling", "Systematic Sampling", "First N% (Sequential)"],
            state="readonly",
            font=("Segoe UI", 9)
        )
        self.method_combo.grid(row=3, column=1, columnspan=2, sticky="ew", padx=(0, 15), pady=8)
        self.method_combo.current(0)
        self.method_combo.bind("<<ComboboxSelected>>", self.on_method_changed)

        # Row 4: Seed and Traceability Options
        lbl_options = tk.Label(
            card,
            text="Options:",
            font=("Segoe UI", 9, "bold"),
            fg="#334155",
            bg="#ffffff"
        )
        lbl_options.grid(row=4, column=0, sticky="w", padx=15, pady=8)

        options_frame = tk.Frame(card, bg="#ffffff")
        options_frame.grid(row=4, column=1, columnspan=2, sticky="ew", padx=(0, 15), pady=8)

        self.seed_lbl = tk.Label(
            options_frame,
            text="Seed:",
            font=("Segoe UI", 9),
            fg="#334155",
            bg="#ffffff"
        )
        self.seed_lbl.pack(side="left", padx=(0, 5))

        seed_entry_border = tk.Frame(options_frame, bg="#cbd5e1")
        seed_entry_border.pack(side="left", padx=(0, 15))

        self.seed_entry = tk.Entry(
            seed_entry_border,
            font=("Segoe UI", 9),
            fg="#1e293b",
            bg="#ffffff",
            relief="flat",
            bd=0,
            width=8
        )
        self.seed_entry.pack(padx=3, pady=3)
        self.seed_entry.insert(0, "42")

        self.trace_var = tk.BooleanVar(value=True)
        self.trace_cb = tk.Checkbutton(
            options_frame,
            text="Add 'Original_Excel_Row' column",
            variable=self.trace_var,
            font=("Segoe UI", 9),
            fg="#334155",
            bg="#ffffff",
            activebackground="#ffffff",
            activeforeground="#334155"
        )
        self.trace_cb.pack(side="left")

        # Row 5: Stats Dashboard
        stats_frame = tk.Frame(card, bg="#ffffff")
        stats_frame.grid(row=5, column=0, columnspan=3, sticky="ew", padx=15, pady=(10, 15))
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

        # Stats Box 2 (Target Rows)
        box2_border = tk.Frame(stats_frame, bg="#e2e8f0")
        box2_border.grid(row=0, column=1, sticky="ew", padx=(8, 0))

        box2 = tk.Frame(box2_border, bg="#f8fafc")
        box2.pack(fill="both", expand=True, padx=1, pady=1)

        lbl2 = tk.Label(box2, text="TARGET ROWS TO EXTRACT", font=("Segoe UI", 8, "bold"), fg="#64748b", bg="#f8fafc")
        lbl2.pack(pady=(10, 2), padx=15, anchor="w")

        self.target_rows_label = tk.Label(box2, text="0", font=("Segoe UI", 14, "bold"), fg="#0f172a", bg="#f8fafc")
        self.target_rows_label.pack(pady=(0, 10), padx=15, anchor="w")

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
            text="Start Extraction",
            command=self.start_extraction,
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
            self.update_target_rows()
            self.set_status("File loaded successfully", "success")
        except Exception as e:
            self.df_cache = None
            self.total_rows = 0
            self.total_rows_label.config(text="0")
            self.update_target_rows()
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

    def update_target_rows(self):
        try:
            val = self.pct_entry.get().strip()
            if not val:
                self.target_rows_label.config(text="-")
                return
            pct = float(val)
            if not (0.0 < pct <= 100.0):
                self.target_rows_label.config(text="-")
                return

            if self.total_rows == 0:
                self.target_rows_label.config(text="0")
                return

            target_rows = max(1, round(self.total_rows * (pct / 100.0)))
            target_rows = min(self.total_rows, target_rows)
            self.target_rows_label.config(text=f"{target_rows:,}")
        except ValueError:
            self.target_rows_label.config(text="-")

    def on_method_changed(self, event=None):
        method = self.method_combo.get()
        if method == "Random Sampling":
            self.seed_entry.config(state="normal")
            self.seed_lbl.config(fg="#334155")
        else:
            self.seed_entry.config(state="disabled")
            self.seed_lbl.config(fg="#94a3b8")

    def start_extraction(self):
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
            pct = float(self.pct_entry.get())
            if not (0.0 < pct <= 100.0):
                raise ValueError
        except ValueError:
            messagebox.showwarning(
                "Invalid Percentage",
                "Percentage must be a decimal number greater than 0% and up to 100%."
            )
            return

        method = self.method_combo.get()
        seed = 42
        if method == "Random Sampling":
            try:
                seed = int(self.seed_entry.get())
            except ValueError:
                messagebox.showwarning(
                    "Invalid Seed",
                    "Random seed must be a valid integer."
                )
                return

        # Disable UI controls during processing
        self.toggle_ui_state("disabled")
        self.set_status("Processing...", "running")
        self.progress["value"] = 10
        self.root.update_idletasks()

        try:
            # Ensure data is loaded
            if self.df_cache is None:
                self.df_cache = pd.read_excel(self.selected_file, engine="openpyxl")
                self.total_rows = len(self.df_cache)
                self.total_rows_label.config(text=f"{self.total_rows:,}")

            if self.total_rows == 0:
                raise ValueError("The selected file contains no data rows to extract.")

            self.progress["value"] = 30
            self.set_status("Extracting subset...", "running")
            self.root.update_idletasks()

            # Calculate target row count
            target_rows = max(1, round(self.total_rows * (pct / 100.0)))
            target_rows = min(self.total_rows, target_rows)

            # Perform extraction based on method
            if method == "Random Sampling":
                extracted_df = self.df_cache.sample(n=target_rows, random_state=seed)
                extracted_df = extracted_df.sort_index()
            elif method == "Systematic Sampling":
                indices = []
                if target_rows <= 1:
                    indices = [0] if self.total_rows > 0 else []
                elif target_rows >= self.total_rows:
                    indices = list(range(self.total_rows))
                else:
                    for i in range(target_rows):
                        idx = int(round(i * (self.total_rows - 1) / (target_rows - 1)))
                        idx = max(0, min(self.total_rows - 1, idx))
                        indices.append(idx)
                    indices = sorted(list(set(indices)))
                extracted_df = self.df_cache.iloc[indices].copy()
            else:  # First N% (Sequential)
                extracted_df = self.df_cache.head(target_rows).copy()

            self.progress["value"] = 60
            self.set_status("Preparing final sheet...", "running")
            self.root.update_idletasks()

            # Add original Excel Row Index if selected
            if self.trace_var.get():
                original_excel_rows = extracted_df.index + 2
                extracted_df.insert(0, "Original_Excel_Row", original_excel_rows)

            # Create output filename and write excel file
            base_name = os.path.splitext(os.path.basename(self.selected_file))[0]
            method_slugs = {
                "Random Sampling": "random",
                "Systematic Sampling": "systematic",
                "First N% (Sequential)": "sequential"
            }
            method_slug = method_slugs.get(method, "extracted")
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            pct_slug = f"{pct:g}".replace(".", "_")

            output_filename = f"{base_name}_extract_{pct_slug}pct_{method_slug}_{timestamp}.xlsx"
            output_path = os.path.join(self.selected_output_folder, output_filename)

            self.progress["value"] = 80
            self.set_status("Writing file to disk...", "running")
            self.root.update_idletasks()

            extracted_df.to_excel(output_path, index=False)

            self.progress["value"] = 100
            self.set_status("Completed successfully", "success")
            messagebox.showinfo(
                "Success",
                f"Successfully extracted {target_rows} rows ({pct}%) using {method}!\n\nSaved to:\n{output_path}"
            )

        except Exception as e:
            self.progress["value"] = 0
            self.set_status("Error during extraction", "error")
            messagebox.showerror(
                "Error",
                f"An error occurred while extracting the data:\n\n{e}"
            )
        finally:
            self.toggle_ui_state("normal")

    def toggle_ui_state(self, state):
        self.start_btn.config(state=state)
        self.reset_btn.config(state=state)
        self.browse_input_btn.config(state=state)
        self.browse_output_btn.config(state=state)
        self.pct_entry.config(state=state)
        self.method_combo.config(state=state if state == "disabled" else "readonly")
        if self.method_combo.get() == "Random Sampling" or state == "disabled":
            self.seed_entry.config(state=state)
        self.trace_cb.config(state=state)

    def reset_fields(self):
        self.selected_file = ""
        self.selected_output_folder = ""
        self.total_rows = 0
        self.df_cache = None

        self.input_file_label.config(text="No file selected", fg="#94a3b8")
        self.output_folder_label.config(text="No folder selected", fg="#94a3b8")

        self.pct_entry.config(state="normal")
        self.pct_entry.delete(0, tk.END)
        self.pct_entry.insert(0, "1.0")

        self.method_combo.config(state="readonly")
        self.method_combo.current(0)
        self.on_method_changed()

        self.seed_entry.config(state="normal")
        self.seed_entry.delete(0, tk.END)
        self.seed_entry.insert(0, "42")

        self.trace_var.set(True)
        self.trace_cb.config(state="normal")

        self.total_rows_label.config(text="0")
        self.target_rows_label.config(text="0")
        
        self.set_status("Ready", "info")
        self.progress["value"] = 0

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelExtractorApp(root)
    root.mainloop()
