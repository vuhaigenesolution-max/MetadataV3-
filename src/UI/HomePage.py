"""HomePage — landing screen với 3 nút điều hướng."""
import tkinter as tk
from tkinter import ttk

from _theme import (
    BG, CARD_BG, PANEL_BG, ACCENT, ACCENT_HOVER,
    TEXT_MAIN, TEXT_SUB, BUTTON_FONT,
)


TITLE_FONT_HOME    = ("Bahnschrift", 22, "bold")
SUBTITLE_FONT_HOME = ("Bahnschrift", 11)


class HomePage(tk.Frame):
    def __init__(self, parent, controller, **kwargs):
        super().__init__(parent, bg=BG, **kwargs)
        self.controller = controller
        self._build_ui()

    def _build_ui(self):
        # Header
        header = tk.Frame(self, bg=PANEL_BG, padx=28, pady=20)
        header.pack(fill="x")

        logo = tk.Canvas(header, width=36, height=36, bg=PANEL_BG, highlightthickness=0)
        logo.create_oval(4, 4, 32, 32, fill=ACCENT, outline=ACCENT)
        logo.create_text(18, 18, text="GS", fill=PANEL_BG, font=("Bahnschrift", 9, "bold"))
        logo.pack(side="left", padx=(0, 14))

        title_block = tk.Frame(header, bg=PANEL_BG)
        title_block.pack(side="left")
        tk.Label(title_block, text="Excel Metadata Tool",
                 font=TITLE_FONT_HOME, fg=TEXT_MAIN, bg=PANEL_BG).pack(anchor="w")
        tk.Label(title_block, text="Chọn chức năng bạn muốn sử dụng",
                 font=SUBTITLE_FONT_HOME, fg=TEXT_SUB, bg=PANEL_BG).pack(anchor="w", pady=(4, 0))

        # Main area
        main_area = tk.Frame(self, bg=BG)
        main_area.pack(expand=True)
        card = ttk.Frame(main_area, style="Card.TFrame", padding=40)
        card.pack(pady=30)

        self._create_big_button(card, "📂   Combine File",
                                self._open_combine)
        self._create_big_button(card, "🧾   Tạo file SampleImport & Manifest",
                                self._open_sample_import)
        self._create_big_button(card, "🌱   Check Sample Import",
                                self._open_seed_file)

        # Footer
        tk.Label(self, text="Gene Solutions • Automation",
                 font=("Bahnschrift", 10), fg=TEXT_SUB, bg=BG).pack(side="bottom", pady=12)

    def _create_big_button(self, parent, text, command):
        frame = tk.Frame(parent, bg=CARD_BG)
        frame.pack(pady=14)
        btn = tk.Label(
            frame, text=text, font=BUTTON_FONT,
            fg=BG, bg=ACCENT, width=34, height=2, cursor="hand2",
        )
        btn.pack()
        btn.bind("<Enter>", lambda e: btn.config(bg=ACCENT_HOVER))
        btn.bind("<Leave>", lambda e: btn.config(bg=ACCENT))
        btn.bind("<Button-1>", lambda e: command())

    # ── Navigation ────────────────────────────────────────
    def _open_combine(self):
        from CombinePage import CombinePage
        self.controller.show_page(CombinePage)

    def _open_sample_import(self):
        from SampleImportPage import SampleImportPage
        self.controller.show_page(SampleImportPage)

    def _open_seed_file(self):
        from SeedFilePage import SeedFilePage
        self.controller.show_page(SeedFilePage)
