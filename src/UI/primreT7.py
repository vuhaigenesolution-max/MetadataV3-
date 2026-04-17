import os
import sys
import json
import time
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# ── Import backend logic ──────────────────────────────────
_SRC_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if _SRC_DIR not in sys.path:
    sys.path.insert(0, _SRC_DIR)

from Logic.Primer_T7 import process_primer_t7


# =========================================================
# Go home
# =========================================================
def go_home():
    try:
        home_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "HomePage.py")
        root.destroy()
        os.startfile(home_path)
    except Exception as e:
        messagebox.showerror("Error", f"Không mở được HomePage.py\n{e}")


# =========================================================
# Persist last paths  (merge với file chung last_paths.json)
# =========================================================
def get_exe_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


CONFIG_PATH = os.path.join(get_exe_dir(), "last_paths.json")


def load_config() -> dict:
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_config(data: dict):
    existing = load_config()
    existing.update(data)
    os.makedirs(os.path.dirname(CONFIG_PATH) or ".", exist_ok=True)
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)


# =========================================================
# UI Helpers
# =========================================================
def update_path_hint(label_widget: tk.Label, path: str, kind: str):
    if not path:
        label_widget.config(text="Chưa chọn", fg="#7f8fa6")
        return
    base = os.path.basename(path.rstrip("/\\"))
    prefix = "📄" if kind == "file" else "📁"
    label_widget.config(text=f"{prefix} {base}", fg="#dcdde1")


# =========================================================
# GUI root
# =========================================================
root = tk.Tk()
root.title("🌟🌟🌟META DATA TOOL🌟🌟🌟")
root.geometry("960x640")
root.configure(bg="#0b1d2c")

BG       = "#0b1d2c"
CARD_BG  = "#0f2536"
PANEL_BG = "#102b42"
LIST_BG  = "#0d2138"
ACCENT      = "#00c4a7"
ACCENT_DARK = "#0c9c85"
SEL_BG   = "#0e4d3e"
TEXT_MAIN = "#e8f1f8"
TEXT_SUB  = "#9fb3c8"

TITLE_FONT    = ("Bahnschrift", 20, "bold")
SUBTITLE_FONT = ("Bahnschrift", 12)
LABEL_FONT    = ("Bahnschrift", 11, "bold")
ENTRY_FONT    = ("Bahnschrift", 11)
BUTTON_FONT   = ("Bahnschrift", 11, "bold")
LIST_FONT     = ("Bahnschrift", 10)

style = ttk.Style()
style.theme_use("clam")
style.configure("Card.TFrame",    background=CARD_BG)
style.configure("Accent.TButton", font=BUTTON_FONT, foreground="#0b1d2c", background=ACCENT)
style.map("Accent.TButton",
          background=[("active", ACCENT_DARK), ("disabled", "#1e3a52")],
          foreground=[("disabled", "#4c5a68")])
style.configure("Outline.TButton", font=BUTTON_FONT, foreground=TEXT_MAIN,
                background=CARD_BG, relief="solid")
style.map("Outline.TButton", background=[("active", PANEL_BG)])
style.configure("Small.TButton", font=("Bahnschrift", 9, "bold"), foreground=TEXT_MAIN,
                background=PANEL_BG, relief="solid")
style.map("Small.TButton", background=[("active", "#1a3e5c")])
style.configure("Success.Horizontal.TProgressbar",
                troughcolor=PANEL_BG, background=ACCENT,
                bordercolor=PANEL_BG, lightcolor=ACCENT, darkcolor=ACCENT_DARK)

cfg = load_config()

# =========================================================
# Header
# =========================================================
header_frame = tk.Frame(root, bg=PANEL_BG, padx=20, pady=16)
header_frame.grid(row=0, column=0, columnspan=4, sticky="nsew")
header_row = tk.Frame(header_frame, bg=PANEL_BG)
header_row.pack(anchor="w")

hero_icon = tk.Canvas(header_row, width=30, height=30, bg=PANEL_BG, highlightthickness=0)
hero_icon.create_oval(4, 4, 26, 26, fill=ACCENT, outline=ACCENT)
hero_icon.create_text(15, 15, text="GS", fill=PANEL_BG, font=("Bahnschrift", 9, "bold"))
hero_icon.pack(side="left", padx=(0, 10))

tk.Label(header_row, text="Excel Pool Allocation Tool",
         font=TITLE_FONT, fg=TEXT_MAIN, bg=PANEL_BG).pack(side="left", anchor="w")

tk.Button(header_row, text="🏠", font=("Bahnschrift", 14),
          bg=PANEL_BG, fg=ACCENT, relief="flat",
          activebackground=PANEL_BG, activeforeground=ACCENT,
          command=go_home, cursor="hand2").pack(side="right", padx=(20, 0))

tk.Label(header_frame, text="Máy T7 — Primer Loại Trừ",
         font=SUBTITLE_FONT, fg=TEXT_SUB, bg=PANEL_BG).pack(anchor="w", pady=(6, 0))

# =========================================================
# Main card
# =========================================================
card = ttk.Frame(root, style="Card.TFrame", padding=20)
card.grid(row=1, column=0, columnspan=4, padx=24, pady=(14, 10), sticky="nsew")
card.grid_columnconfigure(1, weight=1)
root.grid_rowconfigure(1, weight=1)

# ── Row 0: Primer loại trừ ───────────────────────────────
tk.Label(card, text="Primer loại trừ", font=LABEL_FONT,
         fg=TEXT_MAIN, bg=CARD_BG).grid(row=0, column=0, padx=(10, 10), pady=6, sticky="e")

primer_entry = tk.Entry(card, font=ENTRY_FONT, bg="#0d2d44",
                        fg=TEXT_MAIN, relief="flat", insertbackground=TEXT_MAIN)
primer_entry.grid(row=0, column=1, padx=6, pady=6, sticky="ew")
primer_entry.insert(0, cfg.get("t7_primer_path", ""))

primer_hint = tk.Label(card, text="", font=("Bahnschrift", 10), fg=TEXT_SUB, bg=CARD_BG, anchor="w")
primer_hint.grid(row=0, column=3, padx=(6, 8), sticky="w")
update_path_hint(primer_hint, primer_entry.get(), "file")


def browse_primer():
    path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
    if path:
        primer_entry.delete(0, tk.END)
        primer_entry.insert(0, path)
        update_path_hint(primer_hint, path, "file")
        save_config({"t7_primer_path": path})


ttk.Button(card, text="👉Browse", width=10, style="Outline.TButton", command=browse_primer)\
    .grid(row=0, column=2, padx=(8, 10), pady=6, sticky="w")

# ── Row 1: Folder Output (lấy từ output CombinePage) ─────
tk.Label(card, text="Folder Output", font=LABEL_FONT,
         fg=TEXT_MAIN, bg=CARD_BG).grid(row=1, column=0, padx=(10, 10), pady=6, sticky="e")

# Ưu tiên: t7_folder_override → output_path (CombinePage) → ""
_folder_init = cfg.get("t7_folder_override") or cfg.get("output_path", "")

folder_entry = tk.Entry(card, font=ENTRY_FONT, bg="#0d2d44",
                        fg=TEXT_MAIN, relief="flat", insertbackground=TEXT_MAIN)
folder_entry.grid(row=1, column=1, padx=6, pady=6, sticky="ew")
folder_entry.insert(0, _folder_init)

folder_hint = tk.Label(card, text="", font=("Bahnschrift", 10), fg=TEXT_SUB, bg=CARD_BG, anchor="w")
folder_hint.grid(row=1, column=3, padx=(6, 8), sticky="w")
update_path_hint(folder_hint, folder_entry.get(), "folder")


def browse_folder():
    path = filedialog.askdirectory()
    if path:
        folder_entry.delete(0, tk.END)
        folder_entry.insert(0, path)
        update_path_hint(folder_hint, path, "folder")
        save_config({"t7_folder_override": path})
        refresh_file_list()


ttk.Button(card, text="👉Browse", width=10, style="Outline.TButton", command=browse_folder)\
    .grid(row=1, column=2, padx=(8, 10), pady=6, sticky="w")

# ── Divider ───────────────────────────────────────────────
tk.Frame(card, bg=PANEL_BG, height=1).grid(
    row=2, column=0, columnspan=4, sticky="ew", padx=10, pady=(10, 0))

# ── Row 3: File list header ───────────────────────────────
file_header = tk.Frame(card, bg=CARD_BG)
file_header.grid(row=3, column=0, columnspan=4, sticky="ew", padx=10, pady=(6, 4))

tk.Label(file_header, text="Files trong Output Folder",
         font=LABEL_FONT, fg=TEXT_MAIN, bg=CARD_BG).pack(side="left")

file_count_lbl = tk.Label(file_header, text="(0 files)",
                           font=("Bahnschrift", 10), fg=TEXT_SUB, bg=CARD_BG)
file_count_lbl.pack(side="left", padx=(8, 0))

ttk.Button(file_header, text="Refresh",   style="Small.TButton",
           command=lambda: refresh_file_list()).pack(side="right", padx=(4, 0))
ttk.Button(file_header, text="Clear All", style="Small.TButton",
           command=lambda: toggle_all(False)).pack(side="right", padx=(4, 0))
ttk.Button(file_header, text="Select All", style="Small.TButton",
           command=lambda: toggle_all(True)).pack(side="right", padx=(4, 0))

# ── Row 4: Scrollable file list ───────────────────────────
list_outer = tk.Frame(card, bg=CARD_BG)
list_outer.grid(row=4, column=0, columnspan=4, sticky="nsew", padx=10, pady=(0, 8))
card.grid_rowconfigure(4, weight=1)

list_canvas = tk.Canvas(list_outer, bg=LIST_BG, highlightthickness=1,
                         highlightbackground=PANEL_BG, height=240)
list_scrollbar = ttk.Scrollbar(list_outer, orient="vertical", command=list_canvas.yview)
list_canvas.configure(yscrollcommand=list_scrollbar.set)
list_canvas.pack(side="left", fill="both", expand=True)
list_scrollbar.pack(side="right", fill="y")

list_frame = tk.Frame(list_canvas, bg=LIST_BG)
list_window = list_canvas.create_window((0, 0), window=list_frame, anchor="nw")

list_frame.bind("<Configure>", lambda _e: list_canvas.configure(
    scrollregion=list_canvas.bbox("all")))
list_canvas.bind("<Configure>", lambda e: list_canvas.itemconfig(
    list_window, width=e.width))


def _on_mousewheel(event):
    list_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


list_canvas.bind("<MouseWheel>", _on_mousewheel)
list_frame.bind("<MouseWheel>", _on_mousewheel)

file_vars: dict[str, tk.BooleanVar] = {}


def refresh_file_list():
    for w in list_frame.winfo_children():
        w.destroy()
    file_vars.clear()

    folder = folder_entry.get().strip()
    if not folder or not os.path.isdir(folder):
        tk.Label(list_frame, text="  Chưa chọn folder hoặc folder không tồn tại",
                 font=LIST_FONT, fg=TEXT_SUB, bg=LIST_BG).pack(anchor="w", padx=12, pady=8)
        file_count_lbl.config(text="(0 files)")
        return

    files = sorted(f for f in os.listdir(folder) if f.lower().endswith((".xlsx", ".xls")))

    if not files:
        tk.Label(list_frame, text="  Không có file Excel nào trong folder",
                 font=LIST_FONT, fg=TEXT_SUB, bg=LIST_BG).pack(anchor="w", padx=12, pady=8)
        file_count_lbl.config(text="(0 files)")
        return

    for i, fname in enumerate(files):
        row_bg = LIST_BG if i % 2 == 0 else "#0c1e30"
        row = tk.Frame(list_frame, bg=row_bg)
        row.pack(fill="x")
        row._orig_bg = row_bg  # type: ignore[attr-defined]

        var = tk.BooleanVar(value=False)
        file_vars[fname] = var

        def _update_row_color(v=var, r=row):
            bg = SEL_BG if v.get() else r._orig_bg  # type: ignore[attr-defined]
            r.config(bg=bg)
            for child in r.winfo_children():
                if isinstance(child, tk.Label):
                    child.config(bg=bg)

        def _toggle_row(_event, v=var, r=row):
            v.set(not v.get())
            _update_row_color(v, r)

        row.bind("<Button-1>", _toggle_row)

        chk = tk.Checkbutton(row, variable=var, bg=row_bg,
                             activebackground=row_bg, selectcolor=PANEL_BG,
                             command=_update_row_color)
        chk.pack(side="left", padx=(8, 4))

        icon_lbl = tk.Label(row, text="📄", font=("Segoe UI Emoji", 10),
                            bg=row_bg, fg=TEXT_MAIN)
        icon_lbl.pack(side="left", padx=(0, 6))
        icon_lbl.bind("<Button-1>", _toggle_row)

        name_lbl = tk.Label(row, text=fname, font=LIST_FONT,
                            fg=TEXT_MAIN, bg=row_bg, anchor="w")
        name_lbl.pack(side="left", fill="x", expand=True, pady=5)
        name_lbl.bind("<Button-1>", _toggle_row)

    file_count_lbl.config(text=f"({len(files)} files)")


def toggle_all(state: bool):
    for var in file_vars.values():
        var.set(state)
    for row in list_frame.winfo_children():
        if isinstance(row, tk.Frame):
            bg = SEL_BG if state else row._orig_bg  # type: ignore[attr-defined]
            row.config(bg=bg)
            for child in row.winfo_children():
                if isinstance(child, tk.Label):
                    child.config(bg=bg)


if folder_entry.get().strip():
    root.after(100, refresh_file_list)

# =========================================================
# Progress
# =========================================================
progress_var = tk.DoubleVar(value=0)
progress_start_time = 0.0

progress_frame = tk.Frame(card, bg=CARD_BG)
progress_frame.grid(row=5, column=0, columnspan=4, pady=(4, 4))

progress_bar = ttk.Progressbar(progress_frame, variable=progress_var, maximum=100,
                                length=460, style="Success.Horizontal.TProgressbar")
progress_bar.pack(side="left", padx=(0, 10))

progress_percent_label = tk.Label(progress_frame, text="0%",
                                   font=("Bahnschrift", 11, "bold"), fg=TEXT_MAIN, bg=CARD_BG)
progress_percent_label.pack(side="left")

elapsed_time_label = tk.Label(progress_frame, text="0.0s",
                               font=("Bahnschrift", 10), fg=TEXT_SUB, bg=CARD_BG)
elapsed_time_label.pack(side="left", padx=(12, 0))


def ui_set_progress(p):
    progress_var.set(p)
    progress_percent_label.config(text=f"{int(p)}%")
    if progress_start_time:
        elapsed = time.time() - progress_start_time
        elapsed_time_label.config(text=f"{elapsed:.1f}s")
    root.update_idletasks()


# =========================================================
# Open folder button — mở chính Folder Output, ẩn mặc định
# =========================================================
def open_output_folder():
    folder = folder_entry.get().strip()
    if os.path.isdir(folder):
        os.startfile(folder)
    else:
        messagebox.showwarning("Warning", "Folder Output không tồn tại!")


btn_open_folder = ttk.Button(card, text="Open Output Folder",
                              style="Accent.TButton", command=open_output_folder)
btn_open_folder.grid(row=7, column=0, columnspan=4, pady=(8, 0))
btn_open_folder.grid_remove()

# =========================================================
# Run
# =========================================================
btn_frame = tk.Frame(card, bg=CARD_BG)
btn_frame.grid(row=6, column=0, columnspan=4, pady=(6, 0))

btn_run = ttk.Button(btn_frame, text="Run", style="Accent.TButton", width=16)
btn_run.pack(side="left", padx=14, pady=4)


def go_back_combie():
    try:
        combie_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "CombinePage.py")
        root.destroy()
        os.startfile(combie_path)
    except Exception as e:
        messagebox.showerror("Error", f"Không mở được CombinePage.py\n{e}")


btn_back = ttk.Button(btn_frame, text="← Quay lại Combie",
                      style="Outline.TButton", width=18, command=go_back_combie)
btn_back.pack(side="left", padx=4, pady=4)

btn_summary = ttk.Button(btn_frame, text="📋 Xem kết quả",
                         style="Outline.TButton", width=16,
                         command=lambda: show_summary(
                             {"message": "skipped", "total": 0, "duplicates": 0},
                             folder_entry.get().strip()
                         ))
btn_summary.pack(side="left", padx=4, pady=4)


def get_selected_files() -> list:
    folder = folder_entry.get().strip()
    return [os.path.join(folder, fname) for fname, var in file_vars.items() if var.get()]


def validate():
    primer = primer_entry.get().strip()
    folder = folder_entry.get().strip()

    if not primer:
        return False, "Vui lòng chọn file Primer loại trừ."
    if not os.path.isfile(primer):
        return False, "File Primer loại trừ không tồn tại."
    if not folder or not os.path.isdir(folder):
        return False, "Folder Output chưa chọn hoặc không tồn tại."
    if not get_selected_files():
        return False, "Vui lòng chọn ít nhất 1 file để chạy."
    return True, ""


def show_summary(t7_result, output_folder):
    """Cửa sổ tóm tắt sau khi chạy T7 primer — gồm collision, ký tự đặc biệt, T7."""
    cfg_now = load_config()

    win = tk.Toplevel(root)
    win.title("Kết quả xử lý")
    win.geometry("540x460")
    win.configure(bg=BG)
    win.resizable(False, False)
    win.grab_set()

    tk.Label(win, text="Kết quả xử lý", font=TITLE_FONT, fg=TEXT_MAIN, bg=BG).pack(pady=(16, 2))
    tk.Label(win, text=f"Folder: {output_folder}", font=("Bahnschrift", 10),
             fg=TEXT_SUB, bg=BG, wraplength=500).pack(pady=(0, 6))

    frame = tk.Frame(win, bg=CARD_BG, padx=20, pady=14)
    frame.pack(fill="both", expand=True, padx=20, pady=(0, 6))

    def section(title, icon, color=ACCENT):
        tk.Label(frame, text=f"{icon}  {title}", font=LABEL_FONT,
                 fg=color, bg=CARD_BG, anchor="w").pack(fill="x", pady=(10, 2))
        tk.Frame(frame, bg=color, height=1).pack(fill="x")

    def row(text, sub=False):
        tk.Label(frame, text=text,
                 font=("Bahnschrift", 10 if sub else 11),
                 fg=TEXT_SUB if sub else TEXT_MAIN,
                 bg=CARD_BG, anchor="w", wraplength=480, justify="left").pack(fill="x", padx=8, pady=1)

    # ── Collision ──────────────────────────────────────────
    section("Collision trong SampleImport", "⚡")
    cpf = cfg_now.get("last_collision_per_file", {})
    files_with_col = {k: v for k, v in cpf.items() if v > 0}
    if not cpf:
        row("  Chưa có dữ liệu collision (chạy Combie trước)", sub=True)
    elif not files_with_col:
        row("✔  Không phát hiện collision trong tất cả file")
    else:
        row(f"⚠  Phát hiện collision trong {len(files_with_col)} file:")
        for fname, cnt in files_with_col.items():
            row(f"     • {fname}  →  {cnt} cặp", sub=True)

    # ── Ký tự đặc biệt ────────────────────────────────────
    section("Ký tự đặc biệt / khoảng trắng", "🔍")
    err_files = cfg_now.get("last_error_files", [])
    if not err_files and not cpf:
        row("  Chưa có dữ liệu (chạy Combie trước)", sub=True)
    elif not err_files:
        row("✔  Không phát hiện ký tự đặc biệt")
    else:
        row(f"⚠  Phát hiện lỗi trong {len(err_files)} file:")
        for ef in err_files:
            row(f"     • {ef}", sub=True)

    # ── T7 Primer ─────────────────────────────────────────
    section("Mẫu T7 trùng", "🧬")
    msg = t7_result.get("message", "")
    tot = t7_result.get("total", 0)
    dup = t7_result.get("duplicates", 0)
    if msg == "skipped":
        row("  Chưa chạy kiểm tra T7 lần này", sub=True)
    elif msg == "no_data":
        row("  Không đọc được T-primer từ các file đã chọn", sub=True)
    elif msg == "no_duplicate":
        row(f"✔  Không có primer T7 nào trùng  (đã check {tot} primer)")
    elif msg == "ok":
        row(f"⚠  Phát hiện {dup} primer T7 trùng / {tot} primer")
        row("     Xem file: Primer_Check_Result.xlsx", sub=True)
    else:
        row(f"  {msg}", sub=True)

    ttk.Button(win, text="Đóng", style="Accent.TButton",
               command=win.destroy).pack(pady=(6, 14))


def on_run():
    ok, msg = validate()
    if not ok:
        messagebox.showerror("Error", msg)
        return

    selected = get_selected_files()
    primer   = primer_entry.get().strip()
    folder   = folder_entry.get().strip()

    save_config({
        "t7_primer_path":    primer,
        "t7_folder_override": folder,   # lưu riêng để không ghi đè output_path của Combie
    })

    btn_open_folder.grid_remove()
    ui_set_progress(0)
    btn_run.config(state="disabled")

    def worker():
        global progress_start_time
        progress_start_time = time.time()
        try:
            def on_progress(current, total_files):
                pct = int(current / total_files * 100) if total_files else 100
                root.after(0, ui_set_progress, pct)

            result = process_primer_t7(
                file_paths=selected,
                output_folder=folder,
                progress_callback=on_progress,
            )

            root.after(0, ui_set_progress, 100)
            root.after(0, btn_open_folder.grid)
            root.after(0, lambda r=result, f=folder: show_summary(r, f))

        except Exception as exc:
            root.after(0, lambda: messagebox.showerror("Error", str(exc)))
            root.after(0, ui_set_progress, 0)
        finally:
            root.after(0, lambda: btn_run.config(state="normal"))

    threading.Thread(target=worker, daemon=True).start()


btn_run.config(command=on_run)

# =========================================================
# Footer
# =========================================================
tk.Label(root, text="Gene Solutions • LAB Automation Tool • By BI Team",
         font=("Bahnschrift", 10), fg=TEXT_SUB, bg=BG).grid(
    row=2, column=0, columnspan=4, pady=(8, 0))

tk.Label(root, text="_____🌟🌟🌟_____",
         font=("Bahnschrift", 10), fg=TEXT_SUB, bg=BG).grid(
    row=3, column=0, columnspan=4, pady=(0, 12))

root.grid_columnconfigure(0, weight=1)
root.mainloop()
