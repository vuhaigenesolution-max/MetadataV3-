import os
import sys
import json
import time
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


# ============= IMPORT BACKEND =============
from backend import run_seed_file

def go_home():
    try:
        home_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "HomePage.py")
        import subprocess
        kwargs = {"creationflags": subprocess.CREATE_NO_WINDOW} if sys.platform == "win32" else {}
        subprocess.Popen([sys.executable, home_path], **kwargs)
        root.destroy()
    except Exception as e:
        messagebox.showerror("Error", f"Không mở được HomePage.py\n{e}")

# =========================================================
# Persist last paths
# =========================================================
def get_exe_dir():
    # Nếu chạy EXE (PyInstaller)
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    # Nếu chạy file .py
    return os.path.dirname(os.path.abspath(__file__))

CONFIG_PATH = os.path.join(get_exe_dir(), "last_paths.json")

DEFAULT_LAST = {
    "source_mode": "file",
    "source_path": "",
    "template_path": "",
    "output_path": ""
}

def load_last_paths():
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            for k, v in DEFAULT_LAST.items():
                data.setdefault(k, v)
            return data
        except Exception:
            return DEFAULT_LAST.copy()
    return DEFAULT_LAST.copy()

def save_last_paths(data: dict):
    existing = load_last_paths()
    existing.update(data)
    folder = os.path.dirname(CONFIG_PATH)
    os.makedirs(folder, exist_ok=True)
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
# GUI
# =========================================================
root = tk.Tk()
root.title("🌟🌟🌟Excel Pool Allocation Tool🌟🌟🌟")
root.geometry("900x520")
root.configure(bg="#0b1d2c")

# Palette & fonts
BG = "#0b1d2c"
CARD_BG = "#0f2536"
PANEL_BG = "#102b42"
ACCENT = "#00c4a7"
ACCENT_DARK = "#0c9c85"
TEXT_MAIN = "#e8f1f8"
TEXT_SUB = "#9fb3c8"

TITLE_FONT = ("Bahnschrift", 20, "bold")
SUBTITLE_FONT = ("Bahnschrift", 12)
LABEL_FONT = ("Bahnschrift", 11, "bold")
ENTRY_FONT = ("Bahnschrift", 11)
BUTTON_FONT = ("Bahnschrift", 11, "bold")

# ttk styling
style = ttk.Style()
style.theme_use("clam")
style.configure("Card.TFrame", background=CARD_BG)
style.configure("Accent.TButton", font=BUTTON_FONT, foreground="#0b1d2c", background=ACCENT)
style.map("Accent.TButton", background=[("active", ACCENT_DARK)], foreground=[("disabled", "#4c5a68")])

style.configure("Outline.TButton", font=BUTTON_FONT, foreground=TEXT_MAIN, background=CARD_BG, relief="solid")
style.map("Outline.TButton", background=[("active", PANEL_BG)])

style.configure(
    "Success.Horizontal.TProgressbar",
    troughcolor=PANEL_BG,
    background=ACCENT,
    bordercolor=PANEL_BG,
    lightcolor=ACCENT,
    darkcolor=ACCENT_DARK
)

# Load last paths
last = load_last_paths()

# Hero header
header_frame = tk.Frame(root, bg=PANEL_BG, padx=20, pady=16)
header_frame.grid(row=0, column=0, columnspan=4, sticky="nsew")
header_row = tk.Frame(header_frame, bg=PANEL_BG)
header_row.pack(anchor="w")

hero_icon = tk.Canvas(header_row, width=30, height=30, bg=PANEL_BG, highlightthickness=0)
hero_icon.create_oval(4, 4, 26, 26, fill=ACCENT, outline=ACCENT)
hero_icon.create_text(15, 15, text="GS", fill=PANEL_BG, font=("Bahnschrift", 9, "bold"))
hero_icon.pack(side="left", padx=(0, 10))

title_lbl = tk.Label(header_row, text="Excel Pool Allocation Tool", font=TITLE_FONT, fg=TEXT_MAIN, bg=PANEL_BG)
title_lbl.pack(side="left", anchor="w")

home_btn = tk.Button(
    header_row,
    text="🏠",
    font=("Bahnschrift", 14),
    bg=PANEL_BG,
    fg=ACCENT,
    relief="flat",
    activebackground=PANEL_BG,
    activeforeground=ACCENT,
    command=go_home,
    cursor="hand2"
)
home_btn.pack(side="right", padx=(20, 0))

subtitle_lbl = tk.Label(header_frame, text="Tạo file mồi", font=SUBTITLE_FONT, fg=TEXT_SUB, bg=PANEL_BG)
subtitle_lbl.pack(anchor="w", pady=(6, 0))

# Main card
card = ttk.Frame(root, style="Card.TFrame", padding=20)
card.grid(row=1, column=0, columnspan=4, padx=24, pady=(14, 10), sticky="nsew")
card.grid_columnconfigure(1, weight=1)

# =========================
# Source mode
# =========================
mode_var = tk.StringVar(value=last.get("source_mode", "file"))

mode_label = tk.Label(card, text="Source Mode", font=LABEL_FONT, fg=TEXT_MAIN, bg=CARD_BG)
mode_label.grid(row=0, column=0, padx=(10, 10), pady=6, sticky="e")

mode_frame = tk.Frame(card, bg=CARD_BG)
mode_frame.grid(row=0, column=1, padx=6, pady=6, sticky="w")

rb_file = tk.Radiobutton(
    mode_frame, text="File", value="file", variable=mode_var,
    bg=CARD_BG, fg=TEXT_MAIN, selectcolor=PANEL_BG,
    activebackground=CARD_BG, activeforeground=TEXT_MAIN
)
rb_folder = tk.Radiobutton(
    mode_frame, text="Folder", value="folder", variable=mode_var,
    bg=CARD_BG, fg=TEXT_MAIN, selectcolor=PANEL_BG,
    activebackground=CARD_BG, activeforeground=TEXT_MAIN
)
rb_file.pack(side="left", padx=(0, 14))
rb_folder.pack(side="left")

# =========================
# Source path row
# =========================
source_label = tk.Label(card, text="Source", font=LABEL_FONT, fg=TEXT_MAIN, bg=CARD_BG)
source_label.grid(row=1, column=0, padx=(10, 10), pady=6, sticky="e")

source_entry = tk.Entry(card, font=ENTRY_FONT, bg="#0d2d44", fg=TEXT_MAIN, relief="flat", insertbackground=TEXT_MAIN)
source_entry.grid(row=1, column=1, padx=6, pady=6, sticky="ew")
source_entry.insert(0, last.get("source_path", ""))

source_hint = tk.Label(card, text="", font=("Bahnschrift", 10), fg=TEXT_SUB, bg=CARD_BG, anchor="w")
source_hint.grid(row=1, column=3, padx=(6, 8), sticky="w")

def browse_source():
    m = mode_var.get()
    if m == "file":
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            source_entry.delete(0, tk.END)
            source_entry.insert(0, path)
            update_path_hint(source_hint, path, "file")
    else:
        path = filedialog.askdirectory()
        if path:
            source_entry.delete(0, tk.END)
            source_entry.insert(0, path)
            update_path_hint(source_hint, path, "folder")

ttk.Button(card, text="👉Browse", width=10, style="Outline.TButton", command=browse_source)\
    .grid(row=1, column=2, padx=(8, 10), pady=6, sticky="w")

update_path_hint(source_hint, source_entry.get(), "file" if mode_var.get() == "file" else "folder")


# =========================
# Output row
# =========================
output_label = tk.Label(card, text="Destination", font=LABEL_FONT, fg=TEXT_MAIN, bg=CARD_BG)
output_label.grid(row=2, column=0, padx=(10, 10), pady=6, sticky="e")

output_entry = tk.Entry(card, font=ENTRY_FONT, bg="#0d2d44", fg=TEXT_MAIN, relief="flat", insertbackground=TEXT_MAIN)
output_entry.grid(row=2, column=1, padx=6, pady=6, sticky="ew")
output_entry.insert(0, last.get("output_path", ""))

output_hint = tk.Label(card, text="", font=("Bahnschrift", 10), fg=TEXT_SUB, bg=CARD_BG, anchor="w")
output_hint.grid(row=2, column=3, padx=(6, 8), sticky="w")
update_path_hint(output_hint, output_entry.get(), "folder")

def browse_output():
    path = filedialog.askdirectory()
    if path:
        output_entry.delete(0, tk.END)
        output_entry.insert(0, path)
        update_path_hint(output_hint, path, "folder")

ttk.Button(card, text="👉Browse", width=10, style="Outline.TButton", command=browse_output)\
    .grid(row=2, column=2, padx=(8, 10), pady=6, sticky="w")

# =========================
# Template row
# =========================
template_label = tk.Label(card, text="Template", font=LABEL_FONT, fg=TEXT_MAIN, bg=CARD_BG)
template_label.grid(row=3, column=0, padx=(10, 10), pady=6, sticky="e")

template_entry = tk.Entry(card, font=ENTRY_FONT, bg="#0d2d44", fg=TEXT_MAIN, relief="flat", insertbackground=TEXT_MAIN)
template_entry.grid(row=3, column=1, padx=6, pady=6, sticky="ew")
template_entry.insert(0, last.get("template_path", ""))

template_hint = tk.Label(card, text="", font=("Bahnschrift", 10), fg=TEXT_SUB, bg=CARD_BG, anchor="w")
template_hint.grid(row=3, column=3, padx=(6, 8), sticky="w")
update_path_hint(template_hint, template_entry.get(), "file")

def browse_template():
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if path:
        template_entry.delete(0, tk.END)
        template_entry.insert(0, path)
        update_path_hint(template_hint, path, "file")

ttk.Button(card, text="👉Browse", width=10, style="Outline.TButton", command=browse_template)\
    .grid(row=3, column=2, padx=(8, 10), pady=6, sticky="w")

# =========================================================
# Progress
# =========================================================
progress_var = tk.DoubleVar(value=0)
progress_start_time = 0.0

progress_frame = tk.Frame(card, bg=CARD_BG)
progress_frame.grid(row=4, column=0, columnspan=4, pady=(16, 4))

progress_bar = ttk.Progressbar(progress_frame, variable=progress_var, maximum=100, length=460, style="Success.Horizontal.TProgressbar")
progress_bar.pack(side="left", padx=(0, 10))

progress_percent_label = tk.Label(progress_frame, text="0%", font=("Bahnschrift", 11, "bold"), fg=TEXT_MAIN, bg=CARD_BG)
progress_percent_label.pack(side="left")

elapsed_time_label = tk.Label(progress_frame, text="0.0s", font=("Bahnschrift", 10), fg=TEXT_SUB, bg=CARD_BG)
elapsed_time_label.pack(side="left", padx=(12, 0))

def ui_set_progress(p):
    progress_var.set(p)
    progress_percent_label.config(text=f"{int(p)}%")
    if progress_start_time:
        elapsed = time.time() - progress_start_time
        elapsed_time_label.config(text=f"{elapsed:.1f}s")
    root.update_idletasks()

def disable_buttons(disabled: bool):
    state = "disabled" if disabled else "normal"
    btn_run.config(state=state)
    btn_open_folder.config(state=state)

# =========================================================
# Open output folder button
# =========================================================
def open_output_folder():
    folder = output_entry.get().strip()
    if os.path.isdir(folder):
        os.startfile(folder)
    else:
        messagebox.showwarning("Warning", "Destination folder không tồn tại!")

btn_open_folder = ttk.Button(card, text="Open Destination Folder", style="Accent.TButton", command=open_output_folder)
btn_open_folder.grid(row=6, column=0, columnspan=4, pady=(10, 0))
btn_open_folder.grid_remove()

# =========================================================
# Run
# =========================================================
btn_frame = tk.Frame(card, bg=CARD_BG)
btn_frame.grid(row=5, column=0, columnspan=4, pady=(10, 0))

btn_run = ttk.Button(btn_frame, text="Run", style="Accent.TButton", width=16)
btn_run.pack(side="left", padx=14, pady=4)

def validate_paths():
    mode = mode_var.get()
    src  = source_entry.get().strip()
    tmpl = template_entry.get().strip()
    out  = output_entry.get().strip()

    if not src or not tmpl or not out:
        return False, "Vui lòng chọn đủ Source, Template và Destination."

    if mode == "file":
        if not os.path.isfile(src):
            return False, "Source mode=File nhưng Source không phải file."
    else:
        if not os.path.isdir(src):
            return False, "Source mode=Folder nhưng Source không phải folder."

    if not os.path.isfile(tmpl):
        return False, "Template phải là file Excel (.xlsx / .xls)."

    if not os.path.isdir(out):
        return False, "Destination phải là folder."
    return True, ""

def on_run():
    ok, msg = validate_paths()
    if not ok:
        messagebox.showerror("Error", msg)
        return

    data = {
        "source_mode":   mode_var.get(),
        "source_path":   source_entry.get().strip(),
        "template_path": template_entry.get().strip(),
        "output_path":   output_entry.get().strip(),
    }
    save_last_paths(data)

    btn_open_folder.grid_remove()
    ui_set_progress(0)
    disable_buttons(True)

    def worker():
        global progress_start_time
        progress_start_time = time.time()

        try:
            def on_progress(current, total):
                pct = int(current / total * 100) if total else 100
                root.after(0, ui_set_progress, pct)

            result = run_seed_file(
                data["source_mode"],
                data["source_path"],
                data["template_path"],
                data["output_path"],
                progress_callback=on_progress,
            )

            msg_done = f"Hoàn tất! Đã xuất {len(result)} file CSV.\nFolder: {data['output_path']}"
            root.after(0, ui_set_progress, 100)
            root.after(0, btn_open_folder.grid)
            root.after(0, lambda: messagebox.showinfo("Success", msg_done))

        except Exception as e:
            root.after(0, lambda: messagebox.showerror("Error", str(e)))
            root.after(0, ui_set_progress, 0)
        finally:
            root.after(0, lambda: disable_buttons(False))

    threading.Thread(target=worker, daemon=True).start()

btn_run.config(command=on_run)

# Footer
# footer = tk.Label(root, text="Gene Solutions • LAB Automation Tool", font=("Bahnschrift", 10), fg=TEXT_SUB, bg=BG)
# footer = tk.Label(root, text="By BI Team", font=("Bahnschrift", 10), fg=TEXT_SUB, bg=BG)
# footer.grid(row=2, column=0, columnspan=4, pady=(8, 12))
footer1 = tk.Label(root, text="Gene Solutions • LAB Automation Tool • By BI Team",
                   font=("Bahnschrift", 10), fg=TEXT_SUB, bg=BG)
footer1.grid(row=2, column=0, columnspan=4, pady=(8, 0))

footer2 = tk.Label(root, text="_____🌟🌟🌟_____",
                   font=("Bahnschrift", 10), fg=TEXT_SUB, bg=BG)
footer2.grid(row=3, column=0, columnspan=4, pady=(0, 12))


root.grid_columnconfigure(0, weight=1)
root.mainloop()
 