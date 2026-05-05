"""Shared theme, fonts, and helpers for all UI pages."""
import os
import sys
import json
import tkinter as tk
from tkinter import ttk


# ── Color palette ─────────────────────────────────────────
BG          = "#0b1d2c"
CARD_BG     = "#0f2536"
PANEL_BG    = "#102b42"
ACCENT      = "#00c4a7"
ACCENT_DARK = "#0c9c85"
ACCENT_HOVER = "#2ee6c7"
TEXT_MAIN   = "#e8f1f8"
TEXT_SUB    = "#9fb3c8"
ENTRY_BG    = "#0d2d44"

# ── Fonts ─────────────────────────────────────────────────
TITLE_FONT    = ("Bahnschrift", 20, "bold")
SUBTITLE_FONT = ("Bahnschrift", 12)
LABEL_FONT    = ("Bahnschrift", 11, "bold")
ENTRY_FONT    = ("Bahnschrift", 11)
BUTTON_FONT   = ("Bahnschrift", 11, "bold")
HINT_FONT     = ("Bahnschrift", 10)


def apply_ttk_styles():
    """Áp dụng ttk styles dùng chung. Gọi 1 lần khi App khởi tạo."""
    style = ttk.Style()
    style.theme_use("clam")

    style.configure("Card.TFrame", background=CARD_BG)
    style.configure("Panel.TFrame", background=PANEL_BG)

    style.configure("Accent.TButton",
                    font=BUTTON_FONT, foreground=BG, background=ACCENT)
    style.map("Accent.TButton",
              background=[("active", ACCENT_DARK)],
              foreground=[("disabled", "#4c5a68")])

    style.configure("Outline.TButton",
                    font=BUTTON_FONT, foreground=TEXT_MAIN,
                    background=CARD_BG, relief="solid")
    style.map("Outline.TButton",
              background=[("active", PANEL_BG)])

    style.configure("Success.Horizontal.TProgressbar",
                    troughcolor=PANEL_BG,
                    background=ACCENT,
                    bordercolor=PANEL_BG,
                    lightcolor=ACCENT,
                    darkcolor=ACCENT_DARK)


# ── last_paths.json helpers ──────────────────────────────
def get_exe_dir() -> str:
    """Thư mục chứa exe (PyInstaller) hoặc file .py hiện tại."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


CONFIG_PATH = os.path.join(get_exe_dir(), "last_paths.json")


def load_last_paths() -> dict:
    """Đọc toàn bộ last_paths.json. Trả về {} nếu lỗi/không tồn tại."""
    if not os.path.exists(CONFIG_PATH):
        return {}
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError):
        return {}


def save_last_paths(updates: dict):
    """Merge updates vào last_paths.json (giữ nguyên keys cũ)."""
    existing = load_last_paths()
    existing.update(updates)
    folder = os.path.dirname(CONFIG_PATH)
    os.makedirs(folder, exist_ok=True)
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)


# ── UI helpers ────────────────────────────────────────────
def update_path_hint(label_widget: tk.Label, path: str, kind: str = "file"):
    """Cập nhật label hiển thị đường dẫn ngắn gọn (basename) với icon."""
    if not path:
        label_widget.config(text="Chưa chọn", fg="#7f8fa6")
        return
    base = os.path.basename(path.rstrip("/\\"))
    prefix = "📄" if kind == "file" else "📁"
    label_widget.config(text=f"{prefix} {base}", fg="#dcdde1")
