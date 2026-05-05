"""SeedFilePage — Check Sample Import (đối soát SUM ↔ folder metadata)."""
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from _theme import (
    BG, CARD_BG, PANEL_BG, ACCENT, TEXT_MAIN, TEXT_SUB,
    TITLE_FONT, SUBTITLE_FONT, LABEL_FONT, ENTRY_FONT, HINT_FONT, ENTRY_BG,
    load_last_paths, save_last_paths, update_path_hint,
)
from _progress import SmoothProgress
from backend import run_check_sample_number


DEFAULT_LAST = {
    "check_meta_path":    "",
    "check_filesum_path": "",
    "check_output_path":  "",
}


class SeedFilePage(tk.Frame):
    def __init__(self, parent, controller, **kwargs):
        super().__init__(parent, bg=BG, **kwargs)
        self.controller = controller
        self._build_ui()

    def on_show(self):
        """Refresh entries từ last_paths khi quay lại page."""
        self._load_into_entries()

    # ── UI build ──────────────────────────────────────────
    def _build_ui(self):
        # Hero header
        header = tk.Frame(self, bg=PANEL_BG, padx=20, pady=16)
        header.grid(row=0, column=0, columnspan=4, sticky="nsew")
        row1 = tk.Frame(header, bg=PANEL_BG)
        row1.pack(anchor="w", fill="x")

        icon = tk.Canvas(row1, width=30, height=30, bg=PANEL_BG, highlightthickness=0)
        icon.create_oval(4, 4, 26, 26, fill=ACCENT, outline=ACCENT)
        icon.create_text(15, 15, text="GS", fill=PANEL_BG, font=("Bahnschrift", 9, "bold"))
        icon.pack(side="left", padx=(0, 10))

        tk.Label(row1, text="Check Sample Import",
                 font=TITLE_FONT, fg=TEXT_MAIN, bg=PANEL_BG).pack(side="left", anchor="w")

        tk.Button(
            row1, text="🏠", font=("Bahnschrift", 14),
            bg=PANEL_BG, fg=ACCENT, relief="flat",
            activebackground=PANEL_BG, activeforeground=ACCENT,
            command=self._go_home, cursor="hand2",
        ).pack(side="right", padx=(20, 0))

        tk.Label(header, text="Đối soát SUM ↔ folder metadata",
                 font=SUBTITLE_FONT, fg=TEXT_SUB, bg=PANEL_BG).pack(anchor="w", pady=(6, 0))

        # Card container
        card = ttk.Frame(self, style="Card.TFrame", padding=20)
        card.grid(row=1, column=0, columnspan=4, padx=24, pady=(14, 10), sticky="nsew")
        card.grid_columnconfigure(1, weight=1)

        last = {**DEFAULT_LAST, **load_last_paths()}

        # Row 1: Source Metadata (folder)
        self.source_entry, self.source_hint = self._add_path_row(
            card, row=1, label="Source Metadata",
            value=last.get("check_meta_path", ""), kind="folder",
            on_browse=lambda: filedialog.askdirectory(),
        )

        # Row 2: File Sum (file)
        self.filesum_entry, self.filesum_hint = self._add_path_row(
            card, row=2, label="File Sum",
            value=last.get("check_filesum_path", ""), kind="file",
            on_browse=lambda: filedialog.askopenfilename(
                filetypes=[("Excel files", "*.xlsx *.xls")]
            ),
        )

        # Row 3: Destination (folder)
        self.output_entry, self.output_hint = self._add_path_row(
            card, row=3, label="Destination",
            value=last.get("check_output_path", ""), kind="folder",
            on_browse=lambda: filedialog.askdirectory(),
        )

        # Row 4: Smooth Progress
        prog_frame = tk.Frame(card, bg=CARD_BG)
        prog_frame.grid(row=4, column=0, columnspan=4, pady=(16, 4))
        self.progress = SmoothProgress(prog_frame, length=520, bg=CARD_BG)
        self.progress.pack()

        # Row 5: Run button
        btn_frame = tk.Frame(card, bg=CARD_BG)
        btn_frame.grid(row=5, column=0, columnspan=4, pady=(10, 0))
        self.btn_run = ttk.Button(btn_frame, text="Run", style="Accent.TButton",
                                  width=16, command=self._on_run)
        self.btn_run.pack(side="left", padx=14, pady=4)

        # Row 6: Open folder (hidden ban đầu)
        self.btn_open_folder = ttk.Button(
            card, text="Open Destination Folder", style="Accent.TButton",
            command=self._open_output_folder,
        )
        self.btn_open_folder.grid(row=6, column=0, columnspan=4, pady=(10, 0))
        self.btn_open_folder.grid_remove()

        # Footer
        tk.Label(self, text="Gene Solutions • LAB Automation Tool • By BI Team",
                 font=HINT_FONT, fg=TEXT_SUB, bg=BG)\
            .grid(row=2, column=0, columnspan=4, pady=(8, 0))
        tk.Label(self, text="_____🌟🌟🌟_____",
                 font=HINT_FONT, fg=TEXT_SUB, bg=BG)\
            .grid(row=3, column=0, columnspan=4, pady=(0, 12))

        self.grid_columnconfigure(0, weight=1)

    def _add_path_row(self, parent, row, label, value, kind, on_browse):
        """Thêm 1 row gồm Label + Entry + Browse button + hint label."""
        tk.Label(parent, text=label, font=LABEL_FONT, fg=TEXT_MAIN, bg=CARD_BG)\
            .grid(row=row, column=0, padx=(10, 10), pady=6, sticky="e")

        entry = tk.Entry(parent, font=ENTRY_FONT, bg=ENTRY_BG, fg=TEXT_MAIN,
                         relief="flat", insertbackground=TEXT_MAIN)
        entry.grid(row=row, column=1, padx=6, pady=6, sticky="ew")
        entry.insert(0, value)

        hint = tk.Label(parent, text="", font=HINT_FONT, fg=TEXT_SUB, bg=CARD_BG, anchor="w")
        hint.grid(row=row, column=3, padx=(6, 8), sticky="w")
        update_path_hint(hint, value, kind)

        def _browse():
            path = on_browse()
            if path:
                entry.delete(0, tk.END)
                entry.insert(0, path)
                update_path_hint(hint, path, kind)

        ttk.Button(parent, text="👉Browse", width=10, style="Outline.TButton",
                   command=_browse)\
            .grid(row=row, column=2, padx=(8, 10), pady=6, sticky="w")
        return entry, hint

    def _load_into_entries(self):
        """Reload last_paths vào các entry (gọi khi quay lại page)."""
        last = {**DEFAULT_LAST, **load_last_paths()}
        for entry, key, kind, hint in [
            (self.source_entry,  "check_meta_path",    "folder", self.source_hint),
            (self.filesum_entry, "check_filesum_path", "file",   self.filesum_hint),
            (self.output_entry,  "check_output_path",  "folder", self.output_hint),
        ]:
            value = last.get(key, "")
            entry.delete(0, tk.END)
            entry.insert(0, value)
            update_path_hint(hint, value, kind)

    # ── Validation ────────────────────────────────────────
    def _validate(self):
        src     = self.source_entry.get().strip()
        filesum = self.filesum_entry.get().strip()
        out     = self.output_entry.get().strip()

        if not src or not filesum or not out:
            return False, "Vui lòng chọn đủ Source Metadata, File Sum và Destination."
        if not os.path.isdir(src):
            return False, "Source Metadata phải là folder chứa các file metadata."
        if not os.path.isfile(filesum):
            return False, "File Sum phải là file Excel (.xlsx / .xls)."
        if not os.path.isdir(out):
            return False, "Destination phải là folder."
        return True, ""

    # ── Actions ───────────────────────────────────────────
    def _on_run(self):
        ok, msg = self._validate()
        if not ok:
            messagebox.showerror("Error", msg)
            return

        data = {
            "check_meta_path":    self.source_entry.get().strip(),
            "check_filesum_path": self.filesum_entry.get().strip(),
            "check_output_path":  self.output_entry.get().strip(),
        }
        save_last_paths(data)

        self.btn_open_folder.grid_remove()
        self.progress.reset()
        self._set_buttons_disabled(True)

        threading.Thread(target=self._worker, args=(data,), daemon=True).start()

    def _worker(self, data):
        try:
            def progress_cb(current, total):
                pct = int(current / total * 100) if total else 100
                # Status text theo phase
                phase = {
                    30:  "Đang đọc folder metadata...",
                    60:  "Đang xử lý file SUM...",
                    85:  "Đang đối soát meta ↔ sum...",
                    100: "Đang xuất file Excel...",
                }.get(pct, None)
                self.after(0, self.progress.set_target, pct, phase)

            result = run_check_sample_number(
                meta_folder=data["check_meta_path"],
                sum_file=data["check_filesum_path"],
                output_folder=data["check_output_path"],
                progress_callback=progress_cb,
            )

            msg_done = (
                f"✓ Đã xuất 3 file Excel\n"
                f"   {data['check_output_path']}\n\n"
                f"📁 metadata_combined.xlsx  ({result['n_meta']} dòng)\n"
                f"📊 {os.path.basename(result['sum_path'])}  ({result['n_sum']} dòng)\n"
                f"⚠ meta_only.xlsx  ({result['n_meta_only']} dòng — meta không có trong sum)"
            )
            self.after(0, lambda: self.progress.finish("✓ Hoàn tất"))
            self.after(0, self.btn_open_folder.grid)
            self.after(0, lambda: messagebox.showinfo("Kết quả", msg_done))

        except Exception as e:
            err_msg = str(e)
            self.after(0, lambda: messagebox.showerror("Error", err_msg))
            self.after(0, self.progress.reset)
        finally:
            self.after(0, lambda: self._set_buttons_disabled(False))

    def _set_buttons_disabled(self, disabled: bool):
        state = "disabled" if disabled else "normal"
        self.btn_run.config(state=state)
        self.btn_open_folder.config(state=state)

    def _open_output_folder(self):
        folder = self.output_entry.get().strip()
        if os.path.isdir(folder):
            os.startfile(folder)
        else:
            messagebox.showwarning("Warning", "Destination folder không tồn tại!")

    def _go_home(self):
        from HomePage import HomePage
        self.controller.show_page(HomePage)
