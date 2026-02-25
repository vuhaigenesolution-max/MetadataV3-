import tkinter as tk
from tkinter import ttk
import subprocess
import sys
import os

# ===== Colors & Fonts =====
BG = "#0b1d2c"
CARD_BG = "#0f2536"
PANEL_BG = "#102b42"
ACCENT = "#00c4a7"
ACCENT_HOVER = "#2ee6c7"
TEXT_MAIN = "#e8f1f8"
TEXT_SUB = "#9fb3c8"

TITLE_FONT = ("Bahnschrift", 22, "bold")
SUBTITLE_FONT = ("Bahnschrift", 11)
BUTTON_FONT = ("Bahnschrift", 12, "bold")

# Lấy đúng thư mục UI
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ===== Open other pages =====
def open_combine():
    subprocess.Popen([sys.executable, os.path.join(BASE_DIR, "CombinePage.py")])

def open_sample_import():
    subprocess.Popen([sys.executable, os.path.join(BASE_DIR, "SampleImportPage.py")])

def open_seed_file():
    subprocess.Popen([sys.executable, os.path.join(BASE_DIR, "SeedFilePage.py")])

# ===== Window =====
root = tk.Tk()
root.title("Excel Metadata Tool")
root.geometry("900x520")
root.configure(bg=BG)

# ===== Style =====
style = ttk.Style()
style.theme_use("clam")

style.configure(
    "Accent.TButton",
    font=BUTTON_FONT,
    foreground="#0b1d2c",
    background=ACCENT,
    padding=(18, 12)
)
style.map("Accent.TButton", background=[("active", ACCENT_HOVER)])
style.configure("Card.TFrame", background=CARD_BG)

# ===== Header =====
header = tk.Frame(root, bg=PANEL_BG, padx=28, pady=20)
header.pack(fill="x")

logo = tk.Canvas(header, width=36, height=36, bg=PANEL_BG, highlightthickness=0)
logo.create_oval(4, 4, 32, 32, fill=ACCENT, outline=ACCENT)
logo.create_text(18, 18, text="GS", fill=PANEL_BG, font=("Bahnschrift", 9, "bold"))
logo.pack(side="left", padx=(0, 14))

title_block = tk.Frame(header, bg=PANEL_BG)
title_block.pack(side="left")

title_lbl = tk.Label(title_block, text="Excel Metadata Tool", font=TITLE_FONT, fg=TEXT_MAIN, bg=PANEL_BG)
title_lbl.pack(anchor="w")

subtitle_lbl = tk.Label(title_block, text="Chọn chức năng bạn muốn sử dụng", font=SUBTITLE_FONT, fg=TEXT_SUB, bg=PANEL_BG)
subtitle_lbl.pack(anchor="w", pady=(4, 0))

# ===== Main Area =====
main_area = tk.Frame(root, bg=BG)
main_area.pack(expand=True)

card = ttk.Frame(main_area, style="Card.TFrame", padding=40)
card.pack(pady=30)

# ===== Fancy Button =====
def create_big_button(parent, text, command):
    frame = tk.Frame(parent, bg=CARD_BG)
    frame.pack(pady=14)

    btn = tk.Label(
        frame,
        text=text,
        font=BUTTON_FONT,
        fg="#0b1d2c",
        bg=ACCENT,
        width=34,
        height=2,
        cursor="hand2"
    )
    btn.pack()

    btn.bind("<Enter>", lambda e: btn.config(bg=ACCENT_HOVER))
    btn.bind("<Leave>", lambda e: btn.config(bg=ACCENT))
    btn.bind("<Button-1>", lambda e: command())

# 3 nút
create_big_button(card, "📂   Combine File", open_combine)
create_big_button(card, "🧾   Tạo file SampleImport & Manifest", open_sample_import)
create_big_button(card, "🌱   Tạo file mồi", open_seed_file)

# ===== Footer =====
footer = tk.Label(root, text="Gene Solutions • Automation", font=("Bahnschrift", 10), fg=TEXT_SUB, bg=BG)
footer.pack(side="bottom", pady=12)

root.mainloop()