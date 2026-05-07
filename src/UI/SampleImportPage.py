"""SampleImportPage — xuất SampleImport CSV + Manifest từ folder metadata."""
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from _theme import (
    BG, CARD_BG, PANEL_BG, ACCENT, TEXT_MAIN, TEXT_SUB,
    TITLE_FONT, SUBTITLE_FONT, LABEL_FONT, ENTRY_FONT, HINT_FONT, ENTRY_BG,
    load_last_paths, save_last_paths, update_path_hint, bind_hint_click,
)
from _progress import SmoothProgress
from backend import run_sample_import


DEFAULT_LAST = {
    "sample_source_mode":      "folder",
    "sample_source_path":      "",
    "sample_output_path":      "",
    "sample_nhat_ky_nam_path": "",
    "sample_nhat_ky_bac_path": "",
    "sample_goi_xn_path":      "",
}


class SampleImportPage(tk.Frame):
    def __init__(self, parent, controller, **kwargs):
        super().__init__(parent, bg=BG, **kwargs)
        self.controller = controller
        self._last_result: dict | None = None
        self._build_ui()

    def on_show(self):
        self._load_into_entries()

    # ── UI ────────────────────────────────────────────────
    def _build_ui(self):
        # Header
        header = tk.Frame(self, bg=PANEL_BG, padx=20, pady=16)
        header.grid(row=0, column=0, columnspan=4, sticky="nsew")
        row1 = tk.Frame(header, bg=PANEL_BG)
        row1.pack(anchor="w", fill="x")

        icon = tk.Canvas(row1, width=30, height=30, bg=PANEL_BG, highlightthickness=0)
        icon.create_oval(4, 4, 26, 26, fill=ACCENT, outline=ACCENT)
        icon.create_text(15, 15, text="GS", fill=PANEL_BG, font=("Bahnschrift", 9, "bold"))
        icon.pack(side="left", padx=(0, 10))

        tk.Label(row1, text="Excel Pool Allocation Tool",
                 font=TITLE_FONT, fg=TEXT_MAIN, bg=PANEL_BG).pack(side="left", anchor="w")
        tk.Button(
            row1, text="🏠", font=("Bahnschrift", 14),
            bg=PANEL_BG, fg=ACCENT, relief="flat",
            activebackground=PANEL_BG, activeforeground=ACCENT,
            command=self._go_home, cursor="hand2",
        ).pack(side="right", padx=(20, 0))

        tk.Label(header, text="Tạo File Sample Import",
                 font=SUBTITLE_FONT, fg=TEXT_SUB, bg=PANEL_BG).pack(anchor="w", pady=(6, 0))

        # Card
        card = ttk.Frame(self, style="Card.TFrame", padding=20)
        card.grid(row=1, column=0, columnspan=4, padx=24, pady=(14, 10), sticky="nsew")
        card.grid_columnconfigure(1, weight=1)

        last = {**DEFAULT_LAST, **load_last_paths()}
        # Auto-fill source từ output Combine nếu chưa có
        if not last.get("sample_source_path"):
            last["sample_source_path"] = last.get("combine_output_path", "")

        # Source mode (file/folder)
        self.mode_var = tk.StringVar(value=last.get("sample_source_mode", "folder"))
        tk.Label(card, text="Source Mode", font=LABEL_FONT, fg=TEXT_MAIN, bg=CARD_BG)\
            .grid(row=0, column=0, padx=(10, 10), pady=6, sticky="e")
        mode_frame = tk.Frame(card, bg=CARD_BG)
        mode_frame.grid(row=0, column=1, padx=6, pady=6, sticky="w")
        for txt, val in [("File", "file"), ("Folder", "folder")]:
            tk.Radiobutton(
                mode_frame, text=txt, value=val, variable=self.mode_var,
                bg=CARD_BG, fg=TEXT_MAIN, selectcolor=PANEL_BG,
                activebackground=CARD_BG, activeforeground=TEXT_MAIN,
            ).pack(side="left", padx=(0, 14))

        # Row 1: Source (file or folder per mode)
        self.source_entry, self.source_hint = self._add_path_row(
            card, row=1, label="Source",
            value=last.get("sample_source_path", ""), kind="file",
            on_browse=self._browse_source,
        )

        # Row 2: Destination
        self.output_entry, self.output_hint = self._add_path_row(
            card, row=2, label="Destination",
            value=last.get("sample_output_path", ""), kind="folder",
            on_browse=lambda: filedialog.askdirectory(),
        )

        # Row 3: Nhật ký dò miền Nam
        self.nam_entry, self.nam_hint = self._add_path_row(
            card, row=3, label="Nhật ký dò miền Nam",
            value=last.get("sample_nhat_ky_nam_path", ""), kind="file",
            on_browse=lambda: filedialog.askopenfilename(
                filetypes=[("Excel files", "*.xlsx *.xls")]),
        )

        # Row 4: Nhật ký dò miền Bắc
        self.bac_entry, self.bac_hint = self._add_path_row(
            card, row=4, label="Nhật ký dò miền Bắc",
            value=last.get("sample_nhat_ky_bac_path", ""), kind="file",
            on_browse=lambda: filedialog.askopenfilename(
                filetypes=[("Excel files", "*.xlsx *.xls")]),
        )

        # Row 5: Gói xét nghiệm
        self.goi_entry, self.goi_hint = self._add_path_row(
            card, row=5, label="Gói xét nghiệm",
            value=last.get("sample_goi_xn_path", ""), kind="file",
            on_browse=lambda: filedialog.askopenfilename(
                filetypes=[("Excel files", "*.xlsx *.xls")]),
        )

        # Row 6: Progress
        prog_frame = tk.Frame(card, bg=CARD_BG)
        prog_frame.grid(row=6, column=0, columnspan=4, pady=(16, 4))
        self.progress = SmoothProgress(prog_frame, length=520, bg=CARD_BG)
        self.progress.pack()

        # Row 8: Run + Xem kết quả
        btn_frame = tk.Frame(card, bg=CARD_BG)
        btn_frame.grid(row=8, column=0, columnspan=4, pady=(10, 0))
        self.btn_run = ttk.Button(btn_frame, text="Run", style="Accent.TButton",
                                  width=16, command=self._on_run)
        self.btn_run.pack(side="left", padx=14, pady=4)
        ttk.Button(btn_frame, text="📋 Xem kết quả", style="Outline.TButton",
                   width=16, command=self._show_summary_skipped)\
            .pack(side="left", padx=4, pady=4)

        # Row 9: Open folder
        self.btn_open_folder = ttk.Button(
            card, text="Open Destination Folder", style="Accent.TButton",
            command=self._open_output_folder,
        )
        self.btn_open_folder.grid(row=9, column=0, columnspan=4, pady=(10, 0))
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
        tk.Label(parent, text=label, font=LABEL_FONT, fg=TEXT_MAIN, bg=CARD_BG)\
            .grid(row=row, column=0, padx=(10, 10), pady=6, sticky="e")
        entry = tk.Entry(parent, font=ENTRY_FONT, bg=ENTRY_BG, fg=TEXT_MAIN,
                         relief="flat", insertbackground=TEXT_MAIN)
        entry.grid(row=row, column=1, padx=6, pady=6, sticky="ew")
        entry.insert(0, value)

        hint = tk.Label(parent, text="", font=HINT_FONT, fg=TEXT_SUB, bg=CARD_BG, anchor="w")
        hint.grid(row=row, column=3, padx=(6, 8), sticky="w")
        update_path_hint(hint, value, kind)
        bind_hint_click(hint, lambda: entry.get().strip(), kind)

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

    def _browse_source(self):
        if self.mode_var.get() == "file":
            return filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        return filedialog.askdirectory()

    def _load_into_entries(self):
        last = {**DEFAULT_LAST, **load_last_paths()}
        if not last.get("sample_source_path"):
            last["sample_source_path"] = last.get("combine_output_path", "")
        self.mode_var.set(last.get("sample_source_mode", "folder"))
        for entry, key, kind, hint in [
            (self.source_entry, "sample_source_path",      "file"   if self.mode_var.get() == "file" else "folder", self.source_hint),
            (self.output_entry, "sample_output_path",      "folder", self.output_hint),
            (self.nam_entry,    "sample_nhat_ky_nam_path", "file",   self.nam_hint),
            (self.bac_entry,    "sample_nhat_ky_bac_path", "file",   self.bac_hint),
            (self.goi_entry,    "sample_goi_xn_path",      "file",   self.goi_hint),
        ]:
            value = last.get(key, "")
            entry.delete(0, tk.END)
            entry.insert(0, value)
            update_path_hint(hint, value, kind)

    # ── Validation ────────────────────────────────────────
    def _validate(self):
        mode = self.mode_var.get()
        src  = self.source_entry.get().strip()
        out  = self.output_entry.get().strip()

        if not src or not out:
            return False, "Vui lòng chọn đủ Source và Destination."

        if mode == "file" and not os.path.isfile(src):
            return False, "Source mode=File nhưng Source không phải file."
        if mode == "folder" and not os.path.isdir(src):
            return False, "Source mode=Folder nhưng Source không phải folder."

        if not os.path.isdir(out):
            return False, "Destination phải là folder."

        for entry, name in [
            (self.nam_entry, "Nhật ký dò miền Nam"),
            (self.bac_entry, "Nhật ký dò miền Bắc"),
            (self.goi_entry, "Gói xét nghiệm"),
        ]:
            v = entry.get().strip()
            if v and not os.path.isfile(v):
                return False, f"File {name} không tồn tại."

        return True, ""

    # ── Actions ───────────────────────────────────────────
    def _on_run(self):
        ok, msg = self._validate()
        if not ok:
            messagebox.showerror("Error", msg)
            return

        data = {
            "sample_source_mode":      self.mode_var.get(),
            "sample_source_path":      self.source_entry.get().strip(),
            "sample_output_path":      self.output_entry.get().strip(),
            "sample_nhat_ky_nam_path": self.nam_entry.get().strip(),
            "sample_nhat_ky_bac_path": self.bac_entry.get().strip(),
            "sample_goi_xn_path":      self.goi_entry.get().strip(),
        }
        save_last_paths(data)

        self.btn_open_folder.grid_remove()
        self.progress.reset()
        self.progress.set_target(2, "Đang khởi động...")
        self._set_buttons_disabled(True)

        threading.Thread(target=self._worker, args=(data,), daemon=True).start()

    def _worker(self, data):
        try:
            def progress_cb(current, total):
                pct = int(current / total * 100) if total else 100
                phase = f"Đang xử lý file {current}/{total}..." if total > 1 else "Đang xử lý..."
                self.after(0, self.progress.set_target, pct, phase)

            result = run_sample_import(
                source_mode=data["sample_source_mode"],
                source_path=data["sample_source_path"],
                output_path=data["sample_output_path"],
                nhat_ky_nam_path=data.get("sample_nhat_ky_nam_path", ""),
                nhat_ky_bac_path=data.get("sample_nhat_ky_bac_path", ""),
                goi_xn_path=data.get("sample_goi_xn_path", ""),
                progress_callback=progress_cb,
            )

            self._last_result = {
                "result":       result,
                "output_path":  data["sample_output_path"],
            }

            self.after(0, lambda: self.progress.finish("✓ Hoàn tất"))
            self.after(0, self.btn_open_folder.grid)
            self.after(0, self._show_summary_skipped)

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

    # ── Summary popup ─────────────────────────────────────
    def _show_summary_skipped(self):
        """Bấm 'Xem kết quả': nếu đã chạy → show kết quả thật; chưa chạy → thông báo."""
        if self._last_result is None:
            messagebox.showinfo("Kết quả", "Chưa chạy lần nào trong session này.\nBấm Run để bắt đầu.")
            return
        self._show_summary(self._last_result["result"], self._last_result["output_path"])

    def _show_summary(self, result: dict, output_folder: str):
        file_results     = result.get("file_results", [])
        j_report_path    = result.get("j_report_path")
        desc_report_path = result.get("desc_report_path")
        df_j_report      = result.get("j_error_report")
        df_desc_report   = result.get("desc_error_report")
        n_files = sum(1 for r in file_results if r.get("csv_import"))

        win = tk.Toplevel(self)
        win.title("Kết quả xử lý")
        win.geometry("540x460")
        win.configure(bg=BG)
        win.resizable(False, False)
        win.grab_set()

        tk.Label(win, text="Kết quả xử lý", font=TITLE_FONT,
                 fg=TEXT_MAIN, bg=BG).pack(pady=(16, 2))
        tk.Label(win, text=f"Folder: {output_folder}", font=HINT_FONT,
                 fg=TEXT_SUB, bg=BG, wraplength=500).pack(pady=(0, 6))

        frame = tk.Frame(win, bg=CARD_BG, padx=20, pady=14)
        frame.pack(fill="both", expand=True, padx=20, pady=(0, 6))

        def section(title, icon):
            tk.Label(frame, text=f"{icon}  {title}", font=LABEL_FONT,
                     fg=ACCENT, bg=CARD_BG, anchor="w")\
                .pack(fill="x", pady=(10, 2))
            tk.Frame(frame, bg=ACCENT, height=1).pack(fill="x")

        def line(text, sub=False):
            tk.Label(frame, text=text,
                     font=("Bahnschrift", 10 if sub else 11),
                     fg=TEXT_SUB if sub else TEXT_MAIN,
                     bg=CARD_BG, anchor="w", wraplength=480, justify="left")\
                .pack(fill="x", padx=8, pady=1)

        # 1. File CSV đã xuất
        section("File CSV đã xuất", "📄")
        if n_files == 0:
            line("  Không có file CSV nào được xuất", sub=True)
        else:
            line(f"✔  Đã xuất {n_files} file (SampleImport + Manifest)")

        # 2. Cột J / Sample Project trống
        section("Cột J / Sample Project trống", "⚡")
        if df_j_report is None or df_j_report.empty:
            line("✔  Không có cảnh báo cột J / Sample Project")
        else:
            counts = df_j_report.groupby("File").size().to_dict() \
                     if "File" in df_j_report.columns else {}
            line(f"⚠  Phát hiện {len(df_j_report)} dòng trống "
                 f"trong {len(counts)} file:")
            for fname, cnt in counts.items():
                line(f"     • {fname}  →  {cnt} dòng", sub=True)
            line(f"     → {os.path.basename(j_report_path or '')}", sub=True)

        # 3. Description lệch bảng labcode
        section("Description lệch bảng labcode", "🔍")
        if df_desc_report is None or df_desc_report.empty:
            line("✔  Tất cả Description khớp bảng labcode")
        else:
            counts = df_desc_report.groupby("File").size().to_dict() \
                     if "File" in df_desc_report.columns else {}
            line(f"⚠  Phát hiện {len(df_desc_report)} dòng lệch "
                 f"trong {len(counts)} file:")
            for fname, cnt in counts.items():
                line(f"     • {fname}  →  {cnt} dòng", sub=True)
            line(f"     → {os.path.basename(desc_report_path or '')}", sub=True)

        ttk.Button(win, text="Đóng", style="Accent.TButton",
                   command=win.destroy).pack(pady=(6, 14))

    def _open_output_folder(self):
        folder = self.output_entry.get().strip()
        if os.path.isdir(folder):
            os.startfile(folder)
        else:
            messagebox.showwarning("Warning", "Destination folder không tồn tại!")

    def _go_home(self):
        from HomePage import HomePage
        self.controller.show_page(HomePage)
