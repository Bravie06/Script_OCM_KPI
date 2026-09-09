"""
gui_ocm_net.py
GUI for the OCM Network KPI Report Updater.
Allows selecting vendor raw files and updating the Daily / Weekly / Monthly
OCM Excel files in one click.
"""

import os
import sys
import threading
import logging
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# Ensure the script directory is in sys.path for direct execution
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from generate_ocm_net_report import (
    process_vendor_files,
    OCM_DEFAULT_FILENAMES,
    PL_SHEETS,
    PL_VENDOR_KEYS,
)


class OCMNetGUI:
    # Main KPI vendor files
    MAIN_VENDOR_LABELS = [
        ('H_2G', 'Huawei 2G'),
        ('H_3G', 'Huawei 3G'),
        ('H_4G', 'Huawei 4G'),
        ('N_2G', 'Nokia 2G'),
        ('N_3G', 'Nokia 3G'),
        ('N_4G', 'Nokia 4G'),
        ('Z_2G', 'ZTE 2G'),
        ('Z_3G', 'ZTE 3G'),
        ('Z_4G', 'ZTE 4G'),
    ]

    # Packet-loss dedicated files
    PL_VENDOR_LABELS = [
        ('H_3G_PL', 'Huawei 3G Packet Loss'),
        ('H_4G_PL', 'Huawei 4G Packet Loss'),
        ('N_3G_PL', 'Nokia 3G Packet Loss'),
        ('N_4G_PL', 'Nokia 4G Packet Loss'),
        ('Z_3G_PL', 'ZTE 3G Packet Loss'),
        ('Z_4G_PL', 'ZTE 4G Packet Loss'),
    ]

    VENDOR_LABELS = MAIN_VENDOR_LABELS + PL_VENDOR_LABELS

    GRAN_LABELS = [
        ('daily',   'Daily'),
        ('weekly',  'Weekly'),
        ('monthly', 'Monthly'),
    ]

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title('OCM Network KPI Report Updater')
        self.root.minsize(720, 760)
        self.root.resizable(True, True)

        # OCM template file paths
        default_ocm_dir = os.path.join(os.getcwd(), 'files', 'ocm net')
        self.ocm_file_vars: dict[str, tk.StringVar] = {
            gran: tk.StringVar(
                value=os.path.join(default_ocm_dir, fname)
            )
            for gran, fname in OCM_DEFAULT_FILENAMES.items()
        }

        # Granularity checkboxes
        self.gran_vars: dict[str, tk.BooleanVar] = {
            gran: tk.BooleanVar(value=True)
            for gran, _ in self.GRAN_LABELS
        }

        # Vendor raw file paths
        self.file_vars: dict[str, tk.StringVar] = {
            key: tk.StringVar() for key, _ in self.VENDOR_LABELS
        }

        # Packet-loss only mode
        self.pkt_only_var = tk.BooleanVar(value=False)

        # Default RAW FILES folder (for auto-detect / browse initial dir)
        self.raw_dir_var = tk.StringVar(
            value=os.path.join(default_ocm_dir, 'RAW FILES')
        )

        self._build_ui()
        self._setup_logging()

    # ─── UI CONSTRUCTION ──────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        outer = ttk.Frame(self.root, padding=12)
        outer.grid(row=0, column=0, sticky='nsew')
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        outer.columnconfigure(1, weight=1)

        row = 0

        # Title
        ttk.Label(
            outer,
            text='OCM Network KPI Report Updater',
            font=('Segoe UI', 13, 'bold'),
        ).grid(row=row, column=0, columnspan=3, sticky='w', pady=(0, 8))
        row += 1

        # ── OCM Template Files ────────────────────────────────────────────────
        ttk.Label(
            outer, text='OCM Template Files', font=('Segoe UI', 10, 'bold')
        ).grid(row=row, column=0, columnspan=3, sticky='w', pady=(4, 2))
        row += 1

        for gran, label in self.GRAN_LABELS:
            # Checkbox
            cb = ttk.Checkbutton(
                outer,
                text=f'{label}:',
                variable=self.gran_vars[gran],
            )
            cb.grid(row=row, column=0, sticky='w', pady=2)
            # File path entry
            ttk.Entry(outer, textvariable=self.ocm_file_vars[gran]).grid(
                row=row, column=1, sticky='ew', padx=6                      
            )
            # Browse button
            ttk.Button(
                outer, text='Browse…',
                command=lambda g=gran: self._browse_ocm_file(g),
            ).grid(row=row, column=2)
            row += 1

        ttk.Separator(outer, orient='horizontal').grid(
            row=row, column=0, columnspan=3, sticky='ew', pady=8
        )
        row += 1

        # ── Vendor Raw Files ──────────────────────────────────────────────────
        hdr_frame = ttk.Frame(outer)
        hdr_frame.grid(row=row, column=0, columnspan=3, sticky='ew')
        hdr_frame.columnconfigure(1, weight=1)

        ttk.Label(
            hdr_frame, text='Vendor Raw Files', font=('Segoe UI', 10, 'bold')
        ).grid(row=0, column=0, sticky='w')

        # Raw files folder (used for auto-detect initial dir)
        ttk.Label(hdr_frame, text='  RAW FILES folder:').grid(
            row=0, column=1, sticky='e', padx=(12, 4)
        )
        ttk.Entry(hdr_frame, textvariable=self.raw_dir_var, width=28).grid(
            row=0, column=2, sticky='ew', padx=(0, 4)
        )
        ttk.Button(hdr_frame, text='…', width=3,
                   command=self._browse_raw_dir).grid(row=0, column=3)
        row += 1

        # Packet-loss only checkbox + auto-detect button on same row
        pkt_row_frame = ttk.Frame(outer)
        pkt_row_frame.grid(row=row, column=0, columnspan=3, sticky='ew')
        ttk.Checkbutton(
            pkt_row_frame,
            text='Packet Loss sheets only  (skip main KPI files)',
            variable=self.pkt_only_var,
            command=self._toggle_pkt_only,
        ).pack(side='left')
        ttk.Button(
            pkt_row_frame,
            text='Auto-detect from RAW FILES folder',
            command=self._auto_detect,
        ).pack(side='right')
        row += 1

        # ── Main KPI vendor rows ──────────────────────────────────────────────
        ttk.Label(
            outer, text='Main KPI Files', font=('Segoe UI', 9, 'italic')
        ).grid(row=row, column=0, columnspan=3, sticky='w', pady=(4, 0))
        row += 1

        self._main_vendor_rows: list[tuple] = []
        for key, label in self.MAIN_VENDOR_LABELS:
            lbl = ttk.Label(outer, text=f'{label}:')
            lbl.grid(row=row, column=0, sticky='w', pady=2)
            ent = ttk.Entry(outer, textvariable=self.file_vars[key])
            ent.grid(row=row, column=1, sticky='ew', padx=6)
            btn = ttk.Button(
                outer, text='Browse…',
                command=lambda k=key: self._browse_vendor_file(k),
            )
            btn.grid(row=row, column=2)
            self._main_vendor_rows.append((lbl, ent, btn))
            row += 1

        # ── Packet-Loss vendor rows ───────────────────────────────────────────
        ttk.Label(
            outer, text='Packet Loss Files', font=('Segoe UI', 9, 'italic')
        ).grid(row=row, column=0, columnspan=3, sticky='w', pady=(6, 0))
        row += 1

        for key, label in self.PL_VENDOR_LABELS:
            ttk.Label(outer, text=f'{label}:').grid(
                row=row, column=0, sticky='w', pady=2
            )
            ttk.Entry(outer, textvariable=self.file_vars[key]).grid(
                row=row, column=1, sticky='ew', padx=6
            )
            ttk.Button(
                outer, text='Browse…',
                command=lambda k=key: self._browse_vendor_file(k),
            ).grid(row=row, column=2)
            row += 1

        ttk.Separator(outer, orient='horizontal').grid(
            row=row, column=0, columnspan=3, sticky='ew', pady=8
        )
        row += 1

        # ── Run button ────────────────────────────────────────────────────────
        self.run_btn = ttk.Button(
            outer,
            text='▶   Update Selected OCM Reports',
            command=self._run,
        )
        self.run_btn.grid(row=row, column=0, columnspan=3, pady=4)
        row += 1

        # Progress bar
        self.progress = ttk.Progressbar(outer, mode='indeterminate')
        self.progress.grid(row=row, column=0, columnspan=3, sticky='ew', pady=(0, 4))
        row += 1

        # Log area
        ttk.Label(outer, text='Log:').grid(
            row=row, column=0, columnspan=3, sticky='w'
        )
        row += 1
        self.log_text = scrolledtext.ScrolledText(
            outer, height=10, state='disabled',
            font=('Consolas', 9), wrap='word',
        )
        self.log_text.grid(row=row, column=0, columnspan=3, sticky='nsew')
        outer.rowconfigure(row, weight=1)

    # ─── LOGGING ──────────────────────────────────────────────────────────────

    def _setup_logging(self) -> None:
        gui = self

        class _Handler(logging.Handler):
            def emit(self_, record):  # noqa: N805
                gui._log(self_.format(record))

        handler = _Handler()
        handler.setFormatter(logging.Formatter('%(message)s'))
        mod_logger = logging.getLogger('generate_ocm_net_report')
        mod_logger.addHandler(handler)
        mod_logger.setLevel(logging.INFO)

    def _log(self, msg: str) -> None:
        def _append():
            self.log_text.configure(state='normal')
            self.log_text.insert('end', msg + '\n')
            self.log_text.see('end')
            self.log_text.configure(state='disabled')
        self.root.after(0, _append)

    # ─── BROWSE CALLBACKS ─────────────────────────────────────────────────────

    def _browse_ocm_file(self, gran: str) -> None:
        current = self.ocm_file_vars[gran].get()
        init_dir = os.path.dirname(current) if current else os.getcwd()
        fp = filedialog.askopenfilename(
            title=f'Select OCM {gran.capitalize()} Level file',
            filetypes=[('Excel files', '*.xlsx *.xls'), ('All files', '*.*')],
            initialdir=init_dir,
        )
        if fp:
            self.ocm_file_vars[gran].set(fp)

    def _browse_raw_dir(self) -> None:
        d = filedialog.askdirectory(
            initialdir=self.raw_dir_var.get() or os.getcwd()
        )
        if d:
            self.raw_dir_var.set(d)

    def _browse_vendor_file(self, key: str) -> None:
        raw_dir = self.raw_dir_var.get()
        init = raw_dir if os.path.isdir(raw_dir) else os.getcwd()
        fp = filedialog.askopenfilename(
            title=f'Select {key} vendor file',
            filetypes=[('Excel files', '*.xlsx *.xls'), ('All files', '*.*')],
            initialdir=init,
        )
        if fp:
            self.file_vars[key].set(fp)

    # ─── TOGGLE PL-ONLY ───────────────────────────────────────────────────────

    def _toggle_pkt_only(self) -> None:
        """Enable/disable main KPI vendor rows based on PL-only checkbox."""
        state = 'disabled' if self.pkt_only_var.get() else 'normal'
        for lbl, ent, btn in self._main_vendor_rows:
            lbl.configure(foreground='grey' if self.pkt_only_var.get() else '')
            ent.configure(state=state)
            btn.configure(state=state)

    # ─── AUTO-DETECT ──────────────────────────────────────────────────────────

    def _auto_detect(self) -> None:
        """Scan the RAW FILES folder and populate vendor file entries by prefix."""
        raw_dir = self.raw_dir_var.get().strip()
        if not raw_dir or not os.path.isdir(raw_dir):
            messagebox.showwarning(
                'Folder not found',
                f'RAW FILES folder not found:\n{raw_dir}\n\n'
                'Set the correct RAW FILES folder path.',
            )
            return

        # More-specific prefixes must come before overlapping ones
        patterns: dict[str, tuple] = {
            'H_3G_PL': ('HUAWEI 3G PACKET',),
            'H_4G_PL': ('HUAWEI 4G PACKET',),
            'H_2G':    ('H 2G', 'H2G'),
            'H_3G':    ('H 3G', 'H3G'),
            'H_4G':    ('H 4G', 'H4G'),
            'N_3G_PL': ('NOKIA 3G PACKET',),
            'N_4G_PL': ('NOKIA 4G PACKET',),
            'N_2G':    ('N 2G', 'N2G'),
            'N_3G':    ('N 3G', 'N3G'),
            'N_4G':    ('N 4G', 'N4G'),
            'Z_3G_PL': ('ZTE 3G PACKET',),
            'Z_4G_PL': ('ZTE 4G PACKET',),
            'Z_2G':    ('Z P',),
            'Z_3G':    ('Z 3G', 'Z3G'),
            'Z_4G':    ('Z-4G', 'Z 4G', 'Z4G'),
        }

        found = 0
        for fname in sorted(os.listdir(raw_dir)):
            if not fname.lower().endswith('.xlsx'):
                continue
            fpath = os.path.join(raw_dir, fname)
            fupper = fname.upper()
            for vg, pfxs in patterns.items():
                if any(fupper.startswith(p.upper()) for p in pfxs):
                    if not self.file_vars[vg].get():
                        self.file_vars[vg].set(fpath)
                        found += 1
                    break

        self._log(f'Auto-detect: found {found} vendor file(s) in {raw_dir}')

    # ─── RUN ──────────────────────────────────────────────────────────────────

    def _run(self) -> None:
        pkt_only = self.pkt_only_var.get()
        sheets_filter = PL_SHEETS if pkt_only else None

        # Collect selected granularities
        granularities = {
            gran for gran, _ in self.GRAN_LABELS
            if self.gran_vars[gran].get()
        }
        if not granularities:
            messagebox.showerror('Error', 'Please select at least one level to update\n(Daily, Weekly, and/or Monthly).')
            return

        # Validate OCM file paths for selected granularities
        ocm_files: dict[str, str] = {}
        for gran in granularities:
            fp = self.ocm_file_vars[gran].get().strip()
            if not fp:
                messagebox.showerror(
                    'Error',
                    f'No file selected for {gran.capitalize()} level.\n'
                    'Please browse to the OCM template file or uncheck the level.'
                )
                return
            if not os.path.isfile(fp):
                messagebox.showerror(
                    'Error',
                    f'{gran.capitalize()} template file not found:\n{fp}'
                )
                return
            ocm_files[gran] = fp

        # Collect vendor files
        vendor_files = {
            key: self.file_vars[key].get().strip()
            for key, _ in self.VENDOR_LABELS
        }
        provided = sum(1 for fp in vendor_files.values() if fp)
        if provided == 0:
            messagebox.showerror('Error', 'Please select at least one vendor file.')
            return

        # Clear log
        self.log_text.configure(state='normal')
        self.log_text.delete('1.0', 'end')
        self.log_text.configure(state='disabled')

        self.run_btn.configure(state='disabled')
        self.progress.start(12)

        def _worker():
            try:
                process_vendor_files(
                    vendor_files,
                    ocm_files,
                    granularities=granularities,
                    sheets_filter=sheets_filter,
                    log_callback=self._log,
                )
                levels = ', '.join(
                    lbl for g, lbl in self.GRAN_LABELS if g in granularities
                )
                self.root.after(
                    0,
                    lambda: messagebox.showinfo(
                        'Done',
                        f'OCM reports updated successfully!\n\nUpdated: {levels}',
                    ),
                )
            except Exception as exc:
                import traceback
                self._log(f'\nFATAL ERROR: {exc}')
                self._log(traceback.format_exc())
                self.root.after(
                    0, lambda: messagebox.showerror('Error', str(exc))
                )
            finally:
                self.root.after(0, self._reset_controls)

        threading.Thread(target=_worker, daemon=True).start()

    def _reset_controls(self) -> None:
        self.progress.stop()
        self.run_btn.configure(state='normal')


# ─── ENTRY POINT ──────────────────────────────────────────────────────────────

def main() -> None:
    root = tk.Tk()
    OCMNetGUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()

