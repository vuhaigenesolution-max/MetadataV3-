"""primreT7 — kiểm tra primer T7 trùng + summary collision/special chars."""
import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Đảm bảo Logic import được
_SRC_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if _SRC_DIR not in sys.path:
    sys.path.insert(0, _SRC_DIR)

from Logic.Primer_T7 import process_primer_t7

from _theme import (
    BG, CARD_BG, PANEL_BG, ACCENT, TEXT_MAIN, TEXT_SUB,
    TITLE_FONT, SUBTITLE_FONT, LABEL_FONT, ENTRY_FONT, HINT_FONT, ENTRY_BG,
    load_last_paths, save_last_paths, update_path_hint, bind_hint_click,
)
from _progress import SmoothProgress


LIST_BG = "#0d2138"
SEL_BG  = "#0e4d3e"
LIST_FONT = ("Bahnschrift", 10)


class PrimerT7Page(tk.Frame):
    def __init__(self, parent, controller, **kwargs):
        super().__init__(parent, bg=BG, **kwargs)
        self.controller = controller
        self.file_vars: dict = {}
        self._last_t7_result: dict | None = None
        self._build_ui()

    def on_show(self):
        self._load_into_entries()
        self.refresh_file_list()

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

        tk.Label(header, text="Máy T7 — Primer Loại Trừ",
                 font=SUBTITLE_FONT, fg=TEXT_SUB, bg=PANEL_BG).pack(anchor="w", pady=(6, 0))

        # Card
        card = ttk.Frame(self, style="Card.TFrame", padding=20)
        card.grid(row=1, column=0, columnspan=4, padx=24, pady=(14, 10), sticky="nsew")
        card.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1)

        cfg = load_last_paths()

        # Row 0: Primer loại trừ
        self.primer_entry, self.primer_hint = self._add_path_row(
            card, row=0, label="Primer loại trừ",
            value=cfg.get("t7_primer_path", ""), kind="file",
            on_browse=lambda: filedialog.askopenfilename(
                filetypes=[("Excel files", "*.xlsx *.xls")]),
            on_save_key="t7_primer_path",
        )

        # Row 1: Folder Output (ưu tiên t7_folder_override → output_path)
        folder_init = cfg.get("t7_folder_override") or cfg.get("combine_output_path", "")
        self.folder_entry, self.folder_hint = self._add_path_row(
            card, row=1, label="Folder Output",
            value=folder_init, kind="folder",
            on_browse=lambda: filedialog.askdirectory(),
            on_save_key="t7_folder_override",
            on_change=self.refresh_file_list,
        )

        # Divider
        tk.Frame(card, bg=PANEL_BG, height=1).grid(
            row=2, column=0, columnspan=4, sticky="ew", padx=10, pady=(10, 0))

        # Row 3: File list header
        file_header = tk.Frame(card, bg=CARD_BG)
        file_header.grid(row=3, column=0, columnspan=4, sticky="ew", padx=10, pady=(6, 4))
        tk.Label(file_header, text="Files trong Output Folder",
                 font=LABEL_FONT, fg=TEXT_MAIN, bg=CARD_BG).pack(side="left")
        self.file_count_lbl = tk.Label(file_header, text="(0 files)",
                                       font=HINT_FONT, fg=TEXT_SUB, bg=CARD_BG)
        self.file_count_lbl.pack(side="left", padx=(8, 0))

        ttk.Button(file_header, text="Refresh", style="Outline.TButton",
                   command=self.refresh_file_list).pack(side="right", padx=(4, 0))
        ttk.Button(file_header, text="Clear All", style="Outline.TButton",
                   command=lambda: self._toggle_all(False)).pack(side="right", padx=(4, 0))
        ttk.Button(file_header, text="Select All", style="Outline.TButton",
                   command=lambda: self._toggle_all(True)).pack(side="right", padx=(4, 0))

        # Row 4: Scrollable file list
        list_outer = tk.Frame(card, bg=CARD_BG)
        list_outer.grid(row=4, column=0, columnspan=4, sticky="nsew", padx=10, pady=(0, 8))
        card.grid_rowconfigure(4, weight=1)

        self.list_canvas = tk.Canvas(list_outer, bg=LIST_BG, highlightthickness=1,
                                     highlightbackground=PANEL_BG, height=200)
        sb = ttk.Scrollbar(list_outer, orient="vertical", command=self.list_canvas.yview)
        self.list_canvas.configure(yscrollcommand=sb.set)
        self.list_canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.list_frame = tk.Frame(self.list_canvas, bg=LIST_BG)
        self.list_window = self.list_canvas.create_window((0, 0), window=self.list_frame, anchor="nw")
        self.list_frame.bind("<Configure>", lambda _e: self.list_canvas.configure(
            scrollregion=self.list_canvas.bbox("all")))
        self.list_canvas.bind("<Configure>", lambda e: self.list_canvas.itemconfig(
            self.list_window, width=e.width))
        self.list_canvas.bind("<MouseWheel>",
            lambda e: self.list_canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))

        # Row 5: Progress
        prog_frame = tk.Frame(card, bg=CARD_BG)
        prog_frame.grid(row=5, column=0, columnspan=4, pady=(4, 4))
        self.progress = SmoothProgress(prog_frame, length=520, bg=CARD_BG)
        self.progress.pack()

        # Row 6: Buttons
        btn_frame = tk.Frame(card, bg=CARD_BG)
        btn_frame.grid(row=6, column=0, columnspan=4, pady=(6, 0))
        self.btn_run = ttk.Button(btn_frame, text="Run", style="Accent.TButton",
                                  width=16, command=self._on_run)
        self.btn_run.pack(side="left", padx=14, pady=4)
        ttk.Button(btn_frame, text="← Quay lại Combie", style="Outline.TButton",
                   width=18, command=self._go_combine).pack(side="left", padx=4, pady=4)
        ttk.Button(btn_frame, text="📋 Xem kết quả", style="Outline.TButton",
                   width=16, command=self._show_summary_skipped).pack(side="left", padx=4, pady=4)

        # Row 7: Open folder
        self.btn_open_folder = ttk.Button(card, text="Open Output Folder",
                                          style="Accent.TButton",
                                          command=self._open_output_folder)
        self.btn_open_folder.grid(row=7, column=0, columnspan=4, pady=(8, 0))
        self.btn_open_folder.grid_remove()

        # Footer
        tk.Label(self, text="Gene Solutions • LAB Automation Tool • By BI Team",
                 font=HINT_FONT, fg=TEXT_SUB, bg=BG)\
            .grid(row=2, column=0, columnspan=4, pady=(8, 0))
        tk.Label(self, text="_____🌟🌟🌟_____",
                 font=HINT_FONT, fg=TEXT_SUB, bg=BG)\
            .grid(row=3, column=0, columnspan=4, pady=(0, 12))
        self.grid_columnconfigure(0, weight=1)

    def _add_path_row(self, parent, row, label, value, kind, on_browse,
                      on_save_key=None, on_change=None):
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
                if on_save_key:
                    save_last_paths({on_save_key: path})
                if on_change:
                    on_change()

        ttk.Button(parent, text="👉Browse", width=10, style="Outline.TButton",
                   command=_browse)\
            .grid(row=row, column=2, padx=(8, 10), pady=6, sticky="w")
        return entry, hint

    def _load_into_entries(self):
        cfg = load_last_paths()
        for entry, key, kind, hint in [
            (self.primer_entry, "t7_primer_path", "file", self.primer_hint),
        ]:
            v = cfg.get(key, "")
            entry.delete(0, tk.END)
            entry.insert(0, v)
            update_path_hint(hint, v, kind)
        # Folder ưu tiên override
        folder = cfg.get("t7_folder_override") or cfg.get("combine_output_path", "")
        self.folder_entry.delete(0, tk.END)
        self.folder_entry.insert(0, folder)
        update_path_hint(self.folder_hint, folder, "folder")

    # ── File list ─────────────────────────────────────────
    def refresh_file_list(self):
        for w in self.list_frame.winfo_children():
            w.destroy()
        self.file_vars.clear()

        folder = self.folder_entry.get().strip()
        if not folder or not os.path.isdir(folder):
            tk.Label(self.list_frame,
                     text="  Chưa chọn folder hoặc folder không tồn tại",
                     font=LIST_FONT, fg=TEXT_SUB, bg=LIST_BG)\
                .pack(anchor="w", padx=12, pady=8)
            self.file_count_lbl.config(text="(0 files)")
            return

        files = sorted(f for f in os.listdir(folder)
                       if f.lower().endswith((".xlsx", ".xls")))
        if not files:
            tk.Label(self.list_frame,
                     text="  Không có file Excel nào trong folder",
                     font=LIST_FONT, fg=TEXT_SUB, bg=LIST_BG)\
                .pack(anchor="w", padx=12, pady=8)
            self.file_count_lbl.config(text="(0 files)")
            return

        for i, fname in enumerate(files):
            row_bg = LIST_BG if i % 2 == 0 else "#0c1e30"
            row = tk.Frame(self.list_frame, bg=row_bg)
            row.pack(fill="x")
            row._orig_bg = row_bg  # type: ignore

            var = tk.BooleanVar(value=False)
            self.file_vars[fname] = var

            def upd(v=var, r=row):
                bg = SEL_BG if v.get() else r._orig_bg  # type: ignore[attr-defined]
                r.config(bg=bg)
                for c in r.winfo_children():
                    if isinstance(c, tk.Label):
                        c.config(bg=bg)

            def toggle(_e, v=var, r=row):
                v.set(not v.get())
                upd(v, r)

            row.bind("<Button-1>", toggle)

            chk = tk.Checkbutton(row, variable=var, bg=row_bg,
                                 activebackground=row_bg, selectcolor=PANEL_BG,
                                 command=upd)
            chk.pack(side="left", padx=(8, 4))
            for txt, fnt in [("📄", ("Segoe UI Emoji", 10)), (fname, LIST_FONT)]:
                lbl = tk.Label(row, text=txt, font=fnt, fg=TEXT_MAIN, bg=row_bg, anchor="w")
                lbl.pack(side="left", padx=(0, 6) if txt == "📄" else (0, 0),
                         fill="x", expand=(txt != "📄"), pady=(0 if txt == "📄" else 5))
                lbl.bind("<Button-1>", toggle)

        self.file_count_lbl.config(text=f"({len(files)} files)")

    def _toggle_all(self, state: bool):
        for v in self.file_vars.values():
            v.set(state)
        for row in self.list_frame.winfo_children():
            if isinstance(row, tk.Frame):
                bg = SEL_BG if state else row._orig_bg  # type: ignore[attr-defined]
                row.config(bg=bg)
                for c in row.winfo_children():
                    if isinstance(c, tk.Label):
                        c.config(bg=bg)

    def _get_selected_files(self) -> list:
        folder = self.folder_entry.get().strip()
        return [os.path.join(folder, f) for f, v in self.file_vars.items() if v.get()]

    # ── Run ───────────────────────────────────────────────
    def _validate(self):
        primer = self.primer_entry.get().strip()
        folder = self.folder_entry.get().strip()
        if not primer:
            return False, "Vui lòng chọn file Primer loại trừ."
        if not os.path.isfile(primer):
            return False, "File Primer loại trừ không tồn tại."
        if not folder or not os.path.isdir(folder):
            return False, "Folder Output chưa chọn hoặc không tồn tại."
        if not self._get_selected_files():
            return False, "Vui lòng chọn ít nhất 1 file để chạy."
        return True, ""

    def _on_run(self):
        ok, msg = self._validate()
        if not ok:
            messagebox.showerror("Error", msg)
            return

        selected = self._get_selected_files()
        primer   = self.primer_entry.get().strip()
        folder   = self.folder_entry.get().strip()

        save_last_paths({
            "t7_primer_path":     primer,
            "t7_folder_override": folder,
        })

        self.btn_open_folder.grid_remove()
        self.progress.reset()
        self.progress.set_target(2, "Đang khởi động...")
        self.btn_run.config(state="disabled")

        def worker():
            try:
                def progress_cb(current, total):
                    pct = int(current / total * 100) if total else 100
                    self.after(0, self.progress.set_target, pct,
                               f"Đang xử lý file {current}/{total}...")

                result = process_primer_t7(
                    file_paths=selected, output_folder=folder,
                    progress_callback=progress_cb,
                )
                self._last_t7_result = result
                self.after(0, lambda: self.progress.finish("✓ Hoàn tất"))
                self.after(0, self.btn_open_folder.grid)
                self.after(0, lambda r=result, f=folder: self._show_summary(r, f))

            except Exception as exc:
                err = str(exc)
                self.after(0, lambda: messagebox.showerror("Error", err))
                self.after(0, self.progress.reset)
            finally:
                self.after(0, lambda: self.btn_run.config(state="normal"))

        threading.Thread(target=worker, daemon=True).start()

    def _show_summary_skipped(self):
        """Nếu đã chạy Run trong session này → show kết quả thật;
        chưa chạy thì show 'skipped'."""
        folder = self.folder_entry.get().strip()
        if self._last_t7_result is not None:
            self._show_summary(self._last_t7_result, folder)
        else:
            self._show_summary(
                {"message": "skipped", "total": 0, "duplicates": 0},
                folder,
            )

    def _show_summary(self, t7_result, output_folder):
        cfg = load_last_paths()
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

        # Collision
        section("Collision trong SampleImport", "⚡")
        cpf = cfg.get("last_collision_per_file", {})
        files_with_col = {k: v for k, v in cpf.items() if v > 0}
        if not cpf:
            line("  Chưa có dữ liệu collision (chạy Combie trước)", sub=True)
        elif not files_with_col:
            line("✔  Không phát hiện collision trong tất cả file")
        else:
            line(f"⚠  Phát hiện collision trong {len(files_with_col)} file:")
            for fname, cnt in files_with_col.items():
                line(f"     • {fname}  →  {cnt} cặp", sub=True)

        # Special chars
        section("Ký tự đặc biệt / khoảng trắng", "🔍")
        err_files = cfg.get("last_error_files", [])
        if not err_files and not cpf:
            line("  Chưa có dữ liệu (chạy Combie trước)", sub=True)
        elif not err_files:
            line("✔  Không phát hiện ký tự đặc biệt")
        else:
            line(f"⚠  Phát hiện lỗi trong {len(err_files)} file:")
            for ef in err_files:
                line(f"     • {ef}", sub=True)

        # T7
        section("Mẫu T7 trùng", "🧬")
        msg = t7_result.get("message", "")
        tot = t7_result.get("total", 0)
        dup = t7_result.get("duplicates", 0)
        if msg == "skipped":
            line("  Chưa chạy kiểm tra T7 lần này", sub=True)
        elif msg == "no_data":
            line("  Không đọc được T-primer từ các file đã chọn", sub=True)
        elif msg == "no_duplicate":
            line(f"✔  Không có primer T7 nào trùng  (đã check {tot} primer)")
        elif msg == "ok":
            line(f"⚠  Phát hiện {dup} primer T7 trùng / {tot} primer")
            line("     Xem file: Primer_Check_Result.xlsx", sub=True)
        else:
            line(f"  {msg}", sub=True)

        ttk.Button(win, text="Đóng", style="Accent.TButton",
                   command=win.destroy).pack(pady=(6, 14))

    def _open_output_folder(self):
        folder = self.folder_entry.get().strip()
        if os.path.isdir(folder):
            os.startfile(folder)
        else:
            messagebox.showwarning("Warning", "Folder Output không tồn tại!")

    def _go_home(self):
        from HomePage import HomePage
        self.controller.show_page(HomePage)

    def _go_combine(self):
        from CombinePage import CombinePage
        self.controller.show_page(CombinePage)
