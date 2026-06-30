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


class ExcelToolkitApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Toolkit v1.0")
        self.root.geometry("680x600")
        self.root.minsize(700, 700)
        self.root.configure(bg="#f1f5f9")

        # Configure shared ttk styles
        self._configure_styles()

        # Build all three frames upfront
        self.frame_home = tk.Frame(self.root, bg="#f1f5f9")
        self.frame_splitter = tk.Frame(self.root, bg="#f1f5f9")
        self.frame_extractor = tk.Frame(self.root, bg="#f1f5f9")
        self.frame_remover = tk.Frame(self.root, bg="#f1f5f9")

        self._build_home(self.frame_home)
        self._build_splitter(self.frame_splitter)
        self._build_extractor(self.frame_extractor)
        self._build_remover(self.frame_remover)

        # Show home screen first
        self._show_frame(self.frame_home)

    # =========================================================
    #  STYLE CONFIGURATION
    # =========================================================

    def _configure_styles(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure(
            "Modern.Horizontal.TProgressbar",
            troughcolor="#e2e8f0",
            background="#3b82f6",
            thickness=12,
            borderwidth=0,
        )

    # =========================================================
    #  FRAME NAVIGATION
    # =========================================================

    def _show_frame(self, frame):
        for f in (self.frame_home, self.frame_splitter, self.frame_extractor, self.frame_remover):
            f.pack_forget()
        frame.pack(fill="both", expand=True)

    def _go_home(self):
        self._show_frame(self.frame_home)

    # =========================================================
    #  SHARED WIDGET HELPERS
    # =========================================================

    def _make_header(self, parent, title, subtitle, show_back=False):
        """Renders the dark indigo header banner. Optionally includes a back button."""
        header = tk.Frame(parent, bg="#1e1b4b")
        header.pack(fill="x", side="top")

        # Back button (left side)
        if show_back:
            back_btn = tk.Button(
                header,
                text="← Back",
                command=self._go_home,
                bg="#1e1b4b",
                fg="#c7d2fe",
                activebackground="#312e81",
                activeforeground="#ffffff",
                relief="flat",
                bd=0,
                font=("Segoe UI", 9, "bold"),
                cursor="hand2",
            )
            back_btn.pack(anchor="w", padx=24, pady=(15, 0))

        title_lbl = tk.Label(
            header,
            text=title,
            font=("Segoe UI", 16, "bold"),
            fg="#ffffff",
            bg="#1e1b4b",
        )
        title_lbl.pack(anchor="w", padx=24, pady=((4 if show_back else 15), 2))

        sub_lbl = tk.Label(
            header,
            text=subtitle,
            font=("Segoe UI", 9),
            fg="#c7d2fe",
            bg="#1e1b4b",
        )
        sub_lbl.pack(anchor="w", padx=24, pady=(0, 15))

        return header

    def _make_card(self, parent):
        """Returns a white card Frame (1px slate border trick)."""
        border = tk.Frame(parent, bg="#cbd5e1")
        border.pack(fill="x", pady=(0, 12))
        card = tk.Frame(border, bg="#ffffff")
        card.pack(fill="x", padx=1, pady=1)
        return card

    def _make_path_row(self, card, row_idx, label_text, btn_text, btn_cmd, first=False):
        """Renders a labelled path-display row inside a card grid."""
        pady = (15, 8) if first else 8

        lbl = tk.Label(
            card,
            text=label_text,
            font=("Segoe UI", 9, "bold"),
            fg="#334155",
            bg="#ffffff",
        )
        lbl.grid(row=row_idx, column=0, sticky="w", padx=15, pady=pady)

        border = tk.Frame(card, bg="#cbd5e1")
        border.grid(row=row_idx, column=1, sticky="ew", padx=(0, 10), pady=pady)

        display = tk.Label(
            border,
            text="No file selected" if "File" in label_text else "No folder selected",
            font=("Segoe UI", 9),
            fg="#94a3b8",
            bg="#f8fafc",
            anchor="w",
            height=2,
            padx=10,
        )
        display.pack(fill="x", padx=1, pady=1)

        btn = tk.Button(
            card,
            text=btn_text,
            command=btn_cmd,
            bg="#4f46e5",
            fg="#ffffff",
            activebackground="#4338ca",
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            font=("Segoe UI", 9, "bold"),
            width=12,
            height=1,
            cursor="hand2",
        )
        btn.grid(row=row_idx, column=2, sticky="ew", padx=(0, 15), pady=pady)

        return display, btn

    def _make_stat_box(self, parent, col, label_text, padx=(0, 0)):
        """Renders a single stat card and returns the value label."""
        border = tk.Frame(parent, bg="#e2e8f0")
        border.grid(row=0, column=col, sticky="ew", padx=padx)

        inner = tk.Frame(border, bg="#f8fafc")
        inner.pack(fill="both", expand=True, padx=1, pady=1)

        tk.Label(
            inner,
            text=label_text,
            font=("Segoe UI", 8, "bold"),
            fg="#64748b",
            bg="#f8fafc",
        ).pack(pady=(10, 2), padx=15, anchor="w")

        value_lbl = tk.Label(
            inner,
            text="0",
            font=("Segoe UI", 14, "bold"),
            fg="#0f172a",
            bg="#f8fafc",
        )
        value_lbl.pack(pady=(0, 10), padx=15, anchor="w")
        return value_lbl

    def _make_status_and_progress(self, parent):
        """Renders the status label + progress bar and returns both."""
        status_row = tk.Frame(parent, bg="#f1f5f9")
        status_row.pack(fill="x", pady=(2, 2))

        tk.Label(
            status_row,
            text="System Status:",
            font=("Segoe UI", 9, "bold"),
            fg="#475569",
            bg="#f1f5f9",
        ).pack(side="left")

        status_lbl = tk.Label(
            status_row,
            text="Ready",
            font=("Segoe UI", 9, "bold"),
            fg="#475569",
            bg="#f1f5f9",
        )
        status_lbl.pack(side="left", padx=5)

        progress = ttk.Progressbar(
            parent,
            orient="horizontal",
            mode="determinate",
            style="Modern.Horizontal.TProgressbar",
        )
        progress.pack(fill="x", pady=(0, 10))

        return status_lbl, progress

    def _set_status(self, status_lbl, text, status_type="info"):
        color_map = {
            "info":    "#475569",
            "success": "#10b981",
            "warning": "#f59e0b",
            "error":   "#ef4444",
            "running": "#3b82f6",
        }
        status_lbl.config(text=text, fg=color_map.get(status_type, "#475569"))
        self.root.update_idletasks()

    # =========================================================
    #  HOME SCREEN
    # =========================================================

    def _build_home(self, parent):
        self._make_header(
            parent,
            title="Excel Processing Toolkit",
            subtitle="Split large spreadsheets or extract random row samples",
        )

        # Main content container
        container = tk.Frame(parent, bg="#f1f5f9")
        container.pack(fill="both", expand=True, padx=40, pady=30)

        # Tagline
        tk.Label(
            container,
            text="Choose a tool to get started",
            font=("Segoe UI", 11),
            fg="#64748b",
            bg="#f1f5f9",
        ).pack(pady=(0, 24))

        # Card holder
        cards_row = tk.Frame(container, bg="#f1f5f9")
        cards_row.pack(fill="x")
        cards_row.columnconfigure(0, weight=1)
        cards_row.columnconfigure(1, weight=1)

        self._make_menu_card(
            cards_row,
            col=0,
            accent="#4f46e5",
            hover="#4338ca",
            icon="✂",
            title="Excel Splitter",
            description="Split one large Excel file\ninto multiple smaller files\nby row count.",
            command=lambda: self._show_frame(self.frame_splitter),
        )

        self._make_menu_card(
            cards_row,
            col=1,
            accent="#0d9488",
            hover="#0f766e",
            icon="⚗",
            title="Excel Extractor",
            description="Randomly extract a percentage\nof rows from an Excel file\nusing a fixed seed.",
            command=lambda: self._show_frame(self.frame_extractor),
        )

        # Second row: additional tools
        cards_row2 = tk.Frame(container, bg="#f1f5f9")
        cards_row2.pack(fill="x", pady=(12, 0))
        cards_row2.columnconfigure(0, weight=1)

        self._make_menu_card(
            cards_row2,
            col=0,
            accent="#b91c1c",
            hover="#991b1b",
            icon="🧹",
            title="Remove Duplicates From Reference",
            description="Remove rows from a SOURCE file\nthat have keys present in a REFERENCE file.",
            command=lambda: self._show_frame(self.frame_remover),
        )

        # Footer
        tk.Label(
            container,
            text="Excel Toolkit v1.0  •  Powered by Python, Pandas & Tkinter",
            font=("Segoe UI", 8),
            fg="#94a3b8",
            bg="#f1f5f9",
        ).pack(side="bottom", pady=20)

    def _make_menu_card(self, parent, col, accent, hover, icon, title, description, command):
        """Renders a large clickable tool-selection card."""
        padx = (0, 12) if col == 0 else (12, 0)

        outer = tk.Frame(parent, bg="#cbd5e1")
        outer.grid(row=0, column=col, sticky="nsew", padx=padx)

        inner = tk.Frame(outer, bg="#ffffff", cursor="hand2")
        inner.pack(fill="both", expand=True, padx=1, pady=1)

        # Icon circle
        icon_lbl = tk.Label(
            inner,
            text=icon,
            font=("Segoe UI", 28),
            fg=accent,
            bg="#ffffff",
        )
        icon_lbl.pack(pady=(28, 6))

        title_lbl = tk.Label(
            inner,
            text=title,
            font=("Segoe UI", 13, "bold"),
            fg="#0f172a",
            bg="#ffffff",
        )
        title_lbl.pack()

        desc_lbl = tk.Label(
            inner,
            text=description,
            font=("Segoe UI", 9),
            fg="#64748b",
            bg="#ffffff",
            justify="center",
        )
        desc_lbl.pack(pady=(6, 20))

        # Action button
        btn = tk.Button(
            inner,
            text=f"Open {title} →",
            command=command,
            bg=accent,
            fg="#ffffff",
            activebackground=hover,
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            font=("Segoe UI", 10, "bold"),
            width=20,
            height=2,
            cursor="hand2",
        )
        btn.pack(pady=(0, 28))

        # Make entire card clickable
        for widget in (inner, icon_lbl, title_lbl, desc_lbl):
            widget.bind("<Button-1>", lambda e, cmd=command: cmd())
            widget.bind("<Enter>", lambda e, f=inner: f.config(bg="#f8fafc"))
            widget.bind("<Leave>", lambda e, f=inner: f.config(bg="#ffffff"))

    # =========================================================
    #  SPLITTER SCREEN
    # =========================================================

    def _build_splitter(self, parent):
        # State
        self._sp_file = ""
        self._sp_folder = ""
        self._sp_total_rows = 0
        self._sp_df = None

        self._make_header(
            parent,
            title="Excel Splitter",
            subtitle="Split large Excel spreadsheets into smaller files in seconds",
            show_back=True,
        )

        container = tk.Frame(parent, bg="#f1f5f9")
        container.pack(fill="both", expand=True, padx=24, pady=(10, 15))

        # Settings card
        card = self._make_card(container)
        card.columnconfigure(0, minsize=130)
        card.columnconfigure(1, weight=1)
        card.columnconfigure(2, minsize=120)

        # Row 0: Input file
        self._sp_input_lbl, self._sp_browse_file_btn = self._make_path_row(
            card, 0, "Input File:", "Browse File", self._sp_browse_file, first=True
        )

        # Row 1: Output folder
        self._sp_output_lbl, self._sp_browse_folder_btn = self._make_path_row(
            card, 1, "Output Folder:", "Browse Folder", self._sp_browse_folder
        )

        # Row 2: Rows per file entry
        tk.Label(
            card,
            text="Rows Per File:",
            font=("Segoe UI", 9, "bold"),
            fg="#334155",
            bg="#ffffff",
        ).grid(row=2, column=0, sticky="w", padx=15, pady=8)

        entry_border = tk.Frame(card, bg="#cbd5e1")
        entry_border.grid(row=2, column=1, sticky="w", padx=(0, 10), pady=8)

        self._sp_rows_entry = tk.Entry(
            entry_border,
            font=("Segoe UI", 10),
            fg="#1e293b",
            bg="#ffffff",
            relief="flat",
            bd=0,
            width=12,
        )
        self._sp_rows_entry.pack(padx=5, pady=5)
        self._sp_rows_entry.insert(0, "34")
        self._sp_rows_entry.bind("<KeyRelease>", lambda e: self._sp_update_stats())

        tk.Label(
            card,
            text="rows / file",
            font=("Segoe UI", 9, "italic"),
            fg="#64748b",
            bg="#ffffff",
        ).grid(row=2, column=2, sticky="w", padx=(0, 15), pady=8)

        # Row 3: Stats
        stats_frame = tk.Frame(card, bg="#ffffff")
        stats_frame.grid(row=3, column=0, columnspan=3, sticky="ew", padx=15, pady=(10, 15))
        stats_frame.columnconfigure(0, weight=1)
        stats_frame.columnconfigure(1, weight=1)

        self._sp_total_rows_lbl = self._make_stat_box(stats_frame, 0, "TOTAL DATA ROWS", padx=(0, 8))
        self._sp_expected_lbl   = self._make_stat_box(stats_frame, 1, "EXPECTED FILES",  padx=(8, 0))

        # Status + progress
        self._sp_status_lbl, self._sp_progress = self._make_status_and_progress(container)

        # Buttons
        btn_frame = tk.Frame(container, bg="#f1f5f9")
        btn_frame.pack(pady=5)

        self._sp_start_btn = tk.Button(
            btn_frame,
            text="Start Splitting",
            command=self._sp_start,
            bg="#10b981",
            fg="#ffffff",
            activebackground="#059669",
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            font=("Segoe UI", 10, "bold"),
            width=22,
            height=2,
            cursor="hand2",
        )
        self._sp_start_btn.pack(side="left", padx=10)

        self._sp_reset_btn = tk.Button(
            btn_frame,
            text="Reset",
            command=self._sp_reset,
            bg="#cbd5e1",
            fg="#334155",
            activebackground="#94a3b8",
            activeforeground="#334155",
            relief="flat",
            bd=0,
            font=("Segoe UI", 10, "bold"),
            width=22,
            height=2,
            cursor="hand2",
        )
        self._sp_reset_btn.pack(side="left", padx=10)

    def _sp_browse_file(self):
        path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls")],
        )
        if not path:
            return
        self._sp_file = path
        self._sp_input_lbl.config(text=os.path.basename(path), fg="#0f172a")
        self._set_status(self._sp_status_lbl, "Reading file...", "running")
        try:
            self._sp_df = pd.read_excel(path, engine="openpyxl")
            self._sp_total_rows = len(self._sp_df)
            self._sp_total_rows_lbl.config(text=f"{self._sp_total_rows:,}")
            self._sp_update_stats()
            self._set_status(self._sp_status_lbl, "File loaded successfully", "success")
        except Exception as e:
            self._sp_df = None
            self._sp_total_rows = 0
            self._sp_total_rows_lbl.config(text="0")
            self._set_status(self._sp_status_lbl, "Failed to read file", "error")
            messagebox.showerror("Error", f"Failed to read Excel file.\n\n{e}")

    def _sp_browse_folder(self):
        path = filedialog.askdirectory(title="Select Output Folder")
        if not path:
            return
        self._sp_folder = path
        name = os.path.basename(path) or path
        self._sp_output_lbl.config(text=name, fg="#0f172a")
        self._set_status(self._sp_status_lbl, "Output folder selected", "info")

    def _sp_update_stats(self):
        try:
            val = int(self._sp_rows_entry.get().strip())
            if val <= 0:
                raise ValueError
            expected = (self._sp_total_rows + val - 1) // val
            self._sp_expected_lbl.config(text=f"{expected:,}")
        except ValueError:
            self._sp_expected_lbl.config(text="-")

    def _sp_toggle_ui(self, state):
        for w in (self._sp_start_btn, self._sp_reset_btn,
                  self._sp_browse_file_btn, self._sp_browse_folder_btn,
                  self._sp_rows_entry):
            w.config(state=state)

    def _sp_start(self):
        if not self._sp_file:
            messagebox.showwarning("Missing File", "Please select an Excel file.")
            return
        if not self._sp_folder:
            messagebox.showwarning("Missing Output Folder", "Please select an output folder.")
            return
        try:
            rows_per_file = int(self._sp_rows_entry.get())
            if rows_per_file <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Invalid Value", "Rows per file must be a positive integer.")
            return

        self._sp_toggle_ui("disabled")
        self._set_status(self._sp_status_lbl, "Processing...", "running")
        self._sp_progress["value"] = 0
        self.root.update_idletasks()

        try:
            if self._sp_df is None:
                self._sp_df = pd.read_excel(self._sp_file, engine="openpyxl")
                self._sp_total_rows = len(self._sp_df)
                self._sp_total_rows_lbl.config(text=f"{self._sp_total_rows:,}")

            expected = (self._sp_total_rows + rows_per_file - 1) // rows_per_file
            if expected == 0:
                raise ValueError("The selected file contains no data rows to split.")

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_folder = os.path.join(self._sp_folder, f"Split_{timestamp}")
            os.makedirs(out_folder, exist_ok=True)

            for i in range(0, self._sp_total_rows, rows_per_file):
                part_num = (i // rows_per_file) + 1
                chunk = self._sp_df.iloc[i : i + rows_per_file]
                chunk.to_excel(os.path.join(out_folder, f"Part_{part_num}.xlsx"), index=False)
                self._sp_progress["value"] = (part_num / expected) * 100
                self._set_status(self._sp_status_lbl, f"Exported part {part_num} of {expected}...", "running")

            self._set_status(self._sp_status_lbl, "Completed successfully", "success")
            messagebox.showinfo(
                "Success",
                f"{expected} files successfully created!\n\nSaved in:\n{out_folder}",
            )
        except Exception as e:
            self._sp_progress["value"] = 0
            self._set_status(self._sp_status_lbl, "Error during splitting", "error")
            messagebox.showerror("Error", f"An error occurred:\n\n{e}")
        finally:
            self._sp_toggle_ui("normal")

    def _sp_reset(self):
        self._sp_file = ""
        self._sp_folder = ""
        self._sp_total_rows = 0
        self._sp_df = None
        self._sp_input_lbl.config(text="No file selected", fg="#94a3b8")
        self._sp_output_lbl.config(text="No folder selected", fg="#94a3b8")
        self._sp_rows_entry.config(state="normal")
        self._sp_rows_entry.delete(0, tk.END)
        self._sp_rows_entry.insert(0, "34")
        self._sp_total_rows_lbl.config(text="0")
        self._sp_expected_lbl.config(text="0")
        self._set_status(self._sp_status_lbl, "Ready", "info")
        self._sp_progress["value"] = 0

    # =========================================================
    #  EXTRACTOR SCREEN
    # =========================================================

    def _build_extractor(self, parent):
        # State
        self._ex_file = ""
        self._ex_folder = ""
        self._ex_total_rows = 0
        self._ex_df = None

        self._make_header(
            parent,
            title="Excel Extractor",
            subtitle="Randomly extract a percentage of rows from an Excel file",
            show_back=True,
        )

        container = tk.Frame(parent, bg="#f1f5f9")
        container.pack(fill="both", expand=True, padx=24, pady=(10, 15))

        # Settings card
        card = self._make_card(container)
        card.columnconfigure(0, minsize=130)
        card.columnconfigure(1, weight=1)
        card.columnconfigure(2, minsize=120)

        # Row 0: Input file
        self._ex_input_lbl, self._ex_browse_file_btn = self._make_path_row(
            card, 0, "Input File:", "Browse File", self._ex_browse_file, first=True
        )

        # Row 1: Output folder
        self._ex_output_lbl, self._ex_browse_folder_btn = self._make_path_row(
            card, 1, "Output Folder:", "Browse Folder", self._ex_browse_folder
        )

        # Row 2: Percentage entry
        tk.Label(
            card,
            text="Percentage:",
            font=("Segoe UI", 9, "bold"),
            fg="#334155",
            bg="#ffffff",
        ).grid(row=2, column=0, sticky="w", padx=15, pady=5)

        pct_border = tk.Frame(card, bg="#cbd5e1")
        pct_border.grid(row=2, column=1, sticky="w", padx=(0, 10), pady=5)

        self._ex_pct_entry = tk.Entry(
            pct_border,
            font=("Segoe UI", 10),
            fg="#1e293b",
            bg="#ffffff",
            relief="flat",
            bd=0,
            width=12,
        )
        self._ex_pct_entry.pack(padx=5, pady=4)
        self._ex_pct_entry.insert(0, "1.0")
        self._ex_pct_entry.bind("<KeyRelease>", lambda e: self._ex_update_stats())

        tk.Label(
            card,
            text="% of rows",
            font=("Segoe UI", 9, "italic"),
            fg="#64748b",
            bg="#ffffff",
        ).grid(row=2, column=2, sticky="w", padx=(0, 15), pady=5)

        # Row 3: Random seed
        tk.Label(
            card,
            text="Random Seed:",
            font=("Segoe UI", 9, "bold"),
            fg="#334155",
            bg="#ffffff",
        ).grid(row=3, column=0, sticky="w", padx=15, pady=5)

        seed_border = tk.Frame(card, bg="#cbd5e1")
        seed_border.grid(row=3, column=1, sticky="w", padx=(0, 10), pady=5)

        self._ex_seed_entry = tk.Entry(
            seed_border,
            font=("Segoe UI", 10),
            fg="#1e293b",
            bg="#ffffff",
            relief="flat",
            bd=0,
            width=12,
        )
        self._ex_seed_entry.pack(padx=5, pady=4)
        self._ex_seed_entry.insert(0, "42")

        tk.Label(
            card,
            text="(for reproducibility)",
            font=("Segoe UI", 9, "italic"),
            fg="#64748b",
            bg="#ffffff",
        ).grid(row=3, column=2, sticky="w", padx=(0, 15), pady=5)

        # Row 4: Traceability checkbox
        self._ex_trace_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            card,
            text="Add 'Original_Excel_Row' reference column",
            variable=self._ex_trace_var,
            font=("Segoe UI", 9),
            fg="#334155",
            bg="#ffffff",
            activebackground="#ffffff",
            activeforeground="#334155",
        ).grid(row=4, column=0, columnspan=3, sticky="w", padx=15, pady=(4, 10))

        # Row 5: Stats
        stats_frame = tk.Frame(card, bg="#ffffff")
        stats_frame.grid(row=5, column=0, columnspan=3, sticky="ew", padx=15, pady=(0, 10))
        stats_frame.columnconfigure(0, weight=1)
        stats_frame.columnconfigure(1, weight=1)

        self._ex_total_rows_lbl  = self._make_stat_box(stats_frame, 0, "TOTAL DATA ROWS",       padx=(0, 8))
        self._ex_target_rows_lbl = self._make_stat_box(stats_frame, 1, "TARGET ROWS TO EXTRACT", padx=(8, 0))

        # Status + progress
        self._ex_status_lbl, self._ex_progress = self._make_status_and_progress(container)

        # Buttons
        btn_frame = tk.Frame(container, bg="#f1f5f9")
        btn_frame.pack(pady=5)

        self._ex_start_btn = tk.Button(
            btn_frame,
            text="Start Extraction",
            command=self._ex_start,
            bg="#0d9488",
            fg="#ffffff",
            activebackground="#0f766e",
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            font=("Segoe UI", 10, "bold"),
            width=22,
            height=2,
            cursor="hand2",
        )
        self._ex_start_btn.pack(side="left", padx=10)

        self._ex_reset_btn = tk.Button(
            btn_frame,
            text="Reset",
            command=self._ex_reset,
            bg="#cbd5e1",
            fg="#334155",
            activebackground="#94a3b8",
            activeforeground="#334155",
            relief="flat",
            bd=0,
            font=("Segoe UI", 10, "bold"),
            width=22,
            height=2,
            cursor="hand2",
        )
        self._ex_reset_btn.pack(side="left", padx=10)

    def _ex_browse_file(self):
        path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls")],
        )
        if not path:
            return
        self._ex_file = path
        self._ex_input_lbl.config(text=os.path.basename(path), fg="#0f172a")
        self._set_status(self._ex_status_lbl, "Reading file...", "running")
        try:
            self._ex_df = pd.read_excel(path, engine="openpyxl")
            self._ex_total_rows = len(self._ex_df)
            self._ex_total_rows_lbl.config(text=f"{self._ex_total_rows:,}")
            self._ex_update_stats()
            self._set_status(self._ex_status_lbl, "File loaded successfully", "success")
        except Exception as e:
            self._ex_df = None
            self._ex_total_rows = 0
            self._ex_total_rows_lbl.config(text="0")
            self._set_status(self._ex_status_lbl, "Failed to read file", "error")
            messagebox.showerror("Error", f"Failed to read Excel file.\n\n{e}")

    def _ex_browse_folder(self):
        path = filedialog.askdirectory(title="Select Output Folder")
        if not path:
            return
        self._ex_folder = path
        name = os.path.basename(path) or path
        self._ex_output_lbl.config(text=name, fg="#0f172a")
        self._set_status(self._ex_status_lbl, "Output folder selected", "info")

    def _ex_update_stats(self):
        if self._ex_total_rows == 0:
            self._ex_target_rows_lbl.config(text="0")
            return
        try:
            pct = float(self._ex_pct_entry.get())
            if not (0.0 < pct <= 100.0):
                raise ValueError
            target = max(1, round(self._ex_total_rows * (pct / 100.0)))
            target = min(self._ex_total_rows, target)
            self._ex_target_rows_lbl.config(text=f"{target:,}")
        except ValueError:
            self._ex_target_rows_lbl.config(text="-")

    def _ex_toggle_ui(self, state):
        for w in (self._ex_start_btn, self._ex_reset_btn,
                  self._ex_browse_file_btn, self._ex_browse_folder_btn,
                  self._ex_pct_entry, self._ex_seed_entry):
            w.config(state=state)

    def _ex_start(self):
        if not self._ex_file:
            messagebox.showwarning("Missing File", "Please select an Excel file.")
            return
        if not self._ex_folder:
            messagebox.showwarning("Missing Output Folder", "Please select an output folder.")
            return
        try:
            pct = float(self._ex_pct_entry.get())
            if not (0.0 < pct <= 100.0):
                raise ValueError
        except ValueError:
            messagebox.showwarning("Invalid Percentage", "Enter a number between 0.01 and 100.")
            return
        try:
            seed = int(self._ex_seed_entry.get())
        except ValueError:
            messagebox.showwarning("Invalid Seed", "Random seed must be a valid integer.")
            return

        self._ex_toggle_ui("disabled")
        self._set_status(self._ex_status_lbl, "Processing...", "running")
        self._ex_progress["value"] = 10
        self.root.update_idletasks()

        try:
            if self._ex_df is None:
                self._ex_df = pd.read_excel(self._ex_file, engine="openpyxl")
                self._ex_total_rows = len(self._ex_df)
                self._ex_total_rows_lbl.config(text=f"{self._ex_total_rows:,}")

            if self._ex_total_rows == 0:
                raise ValueError("The selected file contains no data rows to extract.")

            target = max(1, round(self._ex_total_rows * (pct / 100.0)))
            target = min(self._ex_total_rows, target)

            self._ex_progress["value"] = 30
            self._set_status(self._ex_status_lbl, "Extracting random sample...", "running")
            self.root.update_idletasks()

            extracted_df = self._ex_df.sample(n=target, random_state=seed).sort_index()

            if self._ex_trace_var.get():
                extracted_df.insert(0, "Original_Excel_Row", extracted_df.index + 2)

            self._ex_progress["value"] = 70
            self._set_status(self._ex_status_lbl, "Writing file to disk...", "running")
            self.root.update_idletasks()

            base_name = os.path.splitext(os.path.basename(self._ex_file))[0]
            timestamp  = datetime.now().strftime("%Y%m%d_%H%M%S")
            pct_slug   = f"{pct:g}".replace(".", "_")
            out_name   = f"{base_name}_extracted_{int(pct)}pct.xlsx"
            out_path   = os.path.join(self._ex_folder, out_name)

            extracted_df.to_excel(out_path, index=False)

            self._ex_progress["value"] = 100
            self._set_status(self._ex_status_lbl, "Completed successfully", "success")
            messagebox.showinfo(
                "Success",
                f"Extraction finished.\nGenerated file: {os.path.basename(out_path)}",
            )
        except Exception as e:
            self._set_status(self._ex_status_lbl, "Error during extraction", "error")
            messagebox.showerror("Error", f"An error occurred:\n\n{e}")
        finally:
            self._ex_toggle_ui("normal")

    def _ex_reset(self):
        self._ex_file = ""
        self._ex_folder = ""
        self._ex_total_rows = 0
        self._ex_df = None
        self._ex_input_lbl.config(text="No file selected")
        self._ex_output_lbl.config(text="No folder selected")
        self._ex_pct_entry.config(state="normal")
        self._ex_pct_entry.delete(0, tk.END)
        self._ex_pct_entry.insert(0, "1.0")
        self._ex_seed_entry.config(state="normal")
        self._ex_seed_entry.delete(0, tk.END)
        self._ex_seed_entry.insert(0, "42")
        self._ex_trace_var.set(True)
        self._ex_total_rows_lbl.config(text="0")
        self._ex_target_rows_lbl.config(text="0")
        self._set_status(self._ex_status_lbl, "Ready", "info")
        self._ex_progress["value"] = 0

    # =========================================================
    #  REMOVER SCREEN (Remove records from SOURCE using a REFERENCE)
    # =========================================================

    def _build_remover(self, parent):
        # State
        self._rd_source = ""
        self._rd_reference = ""
        self._rd_folder = ""
        self._rd_source_df = None
        self._rd_reference_df = None
        self._rd_total_rows = 0
        self._rd_ref_rows = 0
        self._rd_key_var = tk.StringVar(value="")

        self._make_header(
            parent,
            title="Remove Records From Reference",
            subtitle="Remove rows from a SOURCE file that exist in a REFERENCE file",
            show_back=True,
        )

        container = tk.Frame(parent, bg="#f1f5f9")
        container.pack(fill="both", expand=True, padx=24, pady=(10, 15))

        card = self._make_card(container)
        card.columnconfigure(0, minsize=160)
        card.columnconfigure(1, weight=1)
        card.columnconfigure(2, minsize=120)

        # Row 0: Source file
        self._rd_source_lbl, self._rd_browse_source_btn = self._make_path_row(
            card, 0, "Source File (to clean):", "Browse File", self._rd_browse_source, first=True
        )

        # Row 1: Reference file
        self._rd_reference_lbl, self._rd_browse_reference_btn = self._make_path_row(
            card, 1, "Reference File (removal list):", "Browse File", self._rd_browse_reference
        )

        # Row 2: Output folder
        self._rd_output_lbl, self._rd_browse_folder_btn = self._make_path_row(
            card, 2, "Output Folder:", "Browse Folder", self._rd_browse_folder
        )

        # Row 3: Primary key selector
        tk.Label(
            card,
            text="Primary Key:",
            font=("Segoe UI", 9, "bold"),
            fg="#334155",
            bg="#ffffff",
        ).grid(row=3, column=0, sticky="w", padx=15, pady=(8, 6))

        from tkinter import ttk as _ttk

        self._rd_key_combo = _ttk.Combobox(
            card,
            textvariable=self._rd_key_var,
            values=[],
            font=("Segoe UI", 10),
            width=32,
            state="disabled",
        )
        self._rd_key_combo.grid(row=3, column=1, sticky="w", padx=(0, 10), pady=(8, 6))

        self._rd_refresh_keys_btn = tk.Button(
            card,
            text="Refresh Columns",
            command=lambda: self._rd_update_pk_options(force=True),
            bg="#4f46e5",
            fg="#ffffff",
            activebackground="#4338ca",
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            font=("Segoe UI", 9, "bold"),
            width=14,
            cursor="hand2",
            state="disabled",
        )
        self._rd_refresh_keys_btn.grid(row=3, column=2, sticky="ew", padx=(0, 15), pady=(8, 6))

        # Row 4: Stats
        stats_frame = tk.Frame(card, bg="#ffffff")
        stats_frame.grid(row=4, column=0, columnspan=3, sticky="ew", padx=15, pady=(0, 10))
        stats_frame.columnconfigure(0, weight=1)
        stats_frame.columnconfigure(1, weight=1)
        stats_frame.columnconfigure(2, weight=1)

        self._rd_total_rows_lbl = self._make_stat_box(stats_frame, 0, "SOURCE ROWS")
        self._rd_ref_rows_lbl   = self._make_stat_box(stats_frame, 1, "REFERENCE ROWS")
        self._rd_removed_lbl    = self._make_stat_box(stats_frame, 2, "ROWS TO REMOVE")

        # Status + progress
        self._rd_status_lbl, self._rd_progress = self._make_status_and_progress(container)

        # Buttons
        btn_frame = tk.Frame(container, bg="#f1f5f9")
        btn_frame.pack(pady=5)

        self._rd_start_btn = tk.Button(
            btn_frame,
            text="Start Removal",
            command=self._rd_start,
            bg="#ef4444",
            fg="#ffffff",
            activebackground="#dc2626",
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            font=("Segoe UI", 10, "bold"),
            width=22,
            height=2,
            cursor="hand2",
        )
        self._rd_start_btn.pack(side="left", padx=10)

        self._rd_reset_btn = tk.Button(
            btn_frame,
            text="Reset",
            command=self._rd_reset,
            bg="#cbd5e1",
            fg="#334155",
            activebackground="#94a3b8",
            activeforeground="#334155",
            relief="flat",
            bd=0,
            font=("Segoe UI", 10, "bold"),
            width=22,
            height=2,
            cursor="hand2",
        )
        self._rd_reset_btn.pack(side="left", padx=10)

    def _normalize_key(self, series):
        return (
            series.fillna("")
            .astype(str)
            .str.strip()
            .str.replace(".0", "", regex=False)
        )

    def _rd_browse_source(self):
        path = filedialog.askopenfilename(
            title="Select SOURCE Excel File (to clean)",
            filetypes=[("Excel Files", "*.xlsx *.xls")],
        )
        if not path:
            return
        self._rd_source = path
        self._rd_source_lbl.config(text=os.path.basename(path), fg="#0f172a")
        self._set_status(self._rd_status_lbl, "Reading source file...", "running")
        try:
            self._rd_source_df = pd.read_excel(path, engine="openpyxl")
            self._rd_total_rows = len(self._rd_source_df)
            self._rd_total_rows_lbl.config(text=f"{self._rd_total_rows:,}")
            self._rd_update_stats()
            self._rd_update_pk_options()
            self._set_status(self._rd_status_lbl, "Source loaded", "success")
        except Exception as e:
            self._rd_source_df = None
            self._rd_total_rows = 0
            self._rd_total_rows_lbl.config(text="0")
            self._set_status(self._rd_status_lbl, "Failed to read source", "error")
            messagebox.showerror("Error", f"Failed to read source file.\n\n{e}")

    def _rd_browse_reference(self):
        path = filedialog.askopenfilename(
            title="Select REFERENCE Excel File (removal list)",
            filetypes=[("Excel Files", "*.xlsx *.xls")],
        )
        if not path:
            return
        self._rd_reference = path
        self._rd_reference_lbl.config(text=os.path.basename(path), fg="#0f172a")
        self._set_status(self._rd_status_lbl, "Reading reference file...", "running")
        try:
            self._rd_reference_df = pd.read_excel(path, engine="openpyxl")
            self._rd_ref_rows = len(self._rd_reference_df)
            self._rd_ref_rows_lbl.config(text=f"{self._rd_ref_rows:,}")
            self._rd_update_stats()
            self._rd_update_pk_options()
            self._set_status(self._rd_status_lbl, "Reference loaded", "success")
        except Exception as e:
            self._rd_reference_df = None
            self._rd_ref_rows = 0
            self._rd_ref_rows_lbl.config(text="0")
            self._set_status(self._rd_status_lbl, "Failed to read reference", "error")
            messagebox.showerror("Error", f"Failed to read reference file.\n\n{e}")

    def _rd_browse_folder(self):
        path = filedialog.askdirectory(title="Select Output Folder")
        if not path:
            return
        self._rd_folder = path
        name = os.path.basename(path) or path
        self._rd_output_lbl.config(text=name, fg="#0f172a")
        self._set_status(self._rd_status_lbl, "Output folder selected", "info")

    def _rd_update_stats(self):
        try:
            if self._rd_source_df is None or self._rd_reference_df is None:
                self._rd_removed_lbl.config(text="0")
                return
            # Ensure key exists
            key = (self._rd_key_var.get() or "").strip()
            if not key:
                self._rd_removed_lbl.config(text="-")
                return
            if key not in self._rd_source_df.columns or key not in self._rd_reference_df.columns:
                self._rd_removed_lbl.config(text="-")
                return

            src_keys = set(self._normalize_key(self._rd_source_df[key]))
            ref_keys = set(self._normalize_key(self._rd_reference_df[key]))
            remove_count = sum(1 for k in src_keys if k in ref_keys)  # approximate unique key overlap
            # For a precise count of rows removed, compute mask
            remove_mask = self._normalize_key(self._rd_source_df[key]).isin(ref_keys)
            precise_remove = int(remove_mask.sum())
            self._rd_removed_lbl.config(text=f"{precise_remove:,}")
        except Exception:
            self._rd_removed_lbl.config(text="-")

    def _rd_toggle_ui(self, state):
        for w in (self._rd_start_btn, self._rd_reset_btn,
                  self._rd_browse_source_btn, self._rd_browse_reference_btn,
                  self._rd_browse_folder_btn):
            w.config(state=state)

    def _rd_update_pk_options(self, force=False):
        """Populate the PK combobox with column names from loaded files."""
        cols = []
        try:
            if self._rd_source_df is not None:
                cols.extend(list(self._rd_source_df.columns))
            if self._rd_reference_df is not None:
                cols.extend(list(self._rd_reference_df.columns))
            cols = sorted(set(cols))
            if cols:
                current = self._rd_key_var.get()
                self._rd_key_combo['values'] = cols
                # enable combobox and refresh button
                self._rd_key_combo.config(state='readonly')
                self._rd_refresh_keys_btn.config(state='normal')
                # If current is empty or not in options, set to default if available
                if (not current or current not in cols) or force:
                    default = "old_company_number" if "old_company_number" in cols else cols[0]
                    self._rd_key_var.set(default)
            else:
                # no columns available yet; disable combobox and refresh
                self._rd_key_combo['values'] = []
                self._rd_key_var.set("")
                self._rd_key_combo.config(state='disabled')
                self._rd_refresh_keys_btn.config(state='disabled')
        except Exception:
            pass

    def _rd_start(self):
        if not self._rd_source:
            messagebox.showwarning("Missing File", "Please select a SOURCE Excel file.")
            return
        if not self._rd_reference:
            messagebox.showwarning("Missing File", "Please select a REFERENCE Excel file.")
            return
        if not self._rd_folder:
            messagebox.showwarning("Missing Output Folder", "Please select an output folder.")
            return

        self._rd_toggle_ui("disabled")
        self._set_status(self._rd_status_lbl, "Processing...", "running")
        self._rd_progress["value"] = 10
        self.root.update_idletasks()

        try:
            if self._rd_source_df is None:
                self._rd_source_df = pd.read_excel(self._rd_source, engine="openpyxl")
                self._rd_total_rows = len(self._rd_source_df)
                self._rd_total_rows_lbl.config(text=f"{self._rd_total_rows:,}")
            if self._rd_reference_df is None:
                self._rd_reference_df = pd.read_excel(self._rd_reference, engine="openpyxl")
                self._rd_ref_rows = len(self._rd_reference_df)
                self._rd_ref_rows_lbl.config(text=f"{self._rd_ref_rows:,}")

            key = (self._rd_key_var.get() or "").strip()
            if not key:
                raise ValueError("Primary key not specified")
            if key not in self._rd_source_df.columns or key not in self._rd_reference_df.columns:
                raise ValueError(f"Column '{key}' not found in one of the files")

            # Normalize
            self._rd_source_df[key] = self._normalize_key(self._rd_source_df[key])
            self._rd_reference_df[key] = self._normalize_key(self._rd_reference_df[key])

            reference_keys = set(self._rd_reference_df[key])
            remove_mask = self._rd_source_df[key].isin(reference_keys)

            removed_records = self._rd_source_df[remove_mask].copy()
            cleaned_output = self._rd_source_df[~remove_mask].copy()

            ts = datetime.now()
            out_file = os.path.join(self._rd_folder, f"Remove_Result_{ts.strftime('%Y%m%d_%H%M%S')}.xlsx")

            with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
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
                            key,
                            len(self._rd_source_df),
                            len(self._rd_reference_df),
                            len(removed_records),
                            len(cleaned_output),
                            ts.strftime("%Y-%m-%d %H:%M:%S"),
                        ],
                    }
                )
                summary.to_excel(writer, sheet_name="Summary", index=False)
                cleaned_output.to_excel(writer, sheet_name="Output", index=False)
                removed_records.to_excel(writer, sheet_name="Removed_Records", index=False)

            self._rd_progress["value"] = 100
            self._set_status(self._rd_status_lbl, "Completed successfully", "success")
            messagebox.showinfo("Success", f"Removal completed!\n\nSaved to:\n{out_file}")
        except Exception as e:
            self._rd_progress["value"] = 0
            self._set_status(self._rd_status_lbl, "Error during removal", "error")
            messagebox.showerror("Error", f"An error occurred:\n\n{e}")
        finally:
            self._rd_toggle_ui("normal")

    def _rd_reset(self):
        self._rd_source = ""
        self._rd_reference = ""
        self._rd_folder = ""
        self._rd_source_df = None
        self._rd_reference_df = None
        self._rd_total_rows = 0
        self._rd_ref_rows = 0
        self._rd_key_var.set("")
        self._rd_key_combo['values'] = []
        self._rd_key_combo.config(state='disabled')
        self._rd_refresh_keys_btn.config(state='disabled')
        self._rd_source_lbl.config(text="No file selected", fg="#94a3b8")
        self._rd_reference_lbl.config(text="No file selected", fg="#94a3b8")
        self._rd_output_lbl.config(text="No folder selected", fg="#94a3b8")
        self._rd_total_rows_lbl.config(text="0")
        self._rd_ref_rows_lbl.config(text="0")
        self._rd_removed_lbl.config(text="0")
        self._set_status(self._rd_status_lbl, "Ready", "info")
        self._rd_progress["value"] = 0

# =========================================================
#  ENTRY POINT
# =========================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelToolkitApp(root)
    root.mainloop()