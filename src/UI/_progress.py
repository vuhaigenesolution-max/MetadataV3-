"""Progress bar widget với smooth animation + status label."""
import tkinter as tk
from tkinter import ttk
from typing import Optional

from _theme import (
    BG, CARD_BG, PANEL_BG, TEXT_MAIN, TEXT_SUB,
    HINT_FONT, LABEL_FONT,
)


class SmoothProgress(tk.Frame):
    """Progress bar tween mượt giữa current → target percentage,
    có status label phía dưới mô tả phase đang chạy."""

    def __init__(self, parent, length: int = 460, bg: str = CARD_BG):
        super().__init__(parent, bg=bg)
        self._bg = bg
        self._var = tk.DoubleVar(value=0)
        self._current = 0.0     # giá trị hiển thị thực tại
        self._target  = 0.0     # giá trị muốn tween tới
        self._anim_id = None

        # row 1: bar + percent
        row1 = tk.Frame(self, bg=bg)
        row1.pack(fill="x")

        self.bar = ttk.Progressbar(
            row1, variable=self._var, maximum=100,
            length=length, style="Success.Horizontal.TProgressbar",
        )
        self.bar.pack(side="left", padx=(0, 10))

        self.lbl_pct = tk.Label(
            row1, text="0%", font=LABEL_FONT, fg=TEXT_MAIN, bg=bg,
        )
        self.lbl_pct.pack(side="left")

        # row 2: status text
        self.lbl_status = tk.Label(
            self, text="", font=HINT_FONT, fg=TEXT_SUB, bg=bg, anchor="w",
        )
        self.lbl_status.pack(fill="x", pady=(4, 0))

    # ── Public API ────────────────────────────────────────

    def reset(self):
        """Set về 0%."""
        if self._anim_id:
            self.after_cancel(self._anim_id)
            self._anim_id = None
        self._current = 0.0
        self._target  = 0.0
        self._var.set(0)
        self.lbl_pct.config(text="0%")
        self.lbl_status.config(text="")

    def set_target(self, pct: float, status: Optional[str] = None):
        """Set target percentage; widget sẽ tween mượt từ current → target.
        Optional: cập nhật status text."""
        self._target = max(0.0, min(100.0, float(pct)))
        if status is not None:
            self.lbl_status.config(text=status)
        self._start_animation()

    def set_status(self, status: str):
        """Chỉ cập nhật status text, không đụng đến percentage."""
        self.lbl_status.config(text=status)

    def finish(self, status: str = "✓ Hoàn tất"):
        """Đặt về 100% và dừng tween."""
        self.set_target(100, status)

    # ── Internal: animation loop ──────────────────────────

    def _start_animation(self):
        if self._anim_id is not None:
            return  # đang chạy rồi
        self._animate_step()

    def _animate_step(self):
        self._anim_id = None
        diff = self._target - self._current
        if abs(diff) < 0.5:
            self._current = self._target
            self._var.set(self._current)
            self.lbl_pct.config(text=f"{int(self._current)}%")
            return

        # Easing: tăng 18% khoảng cách mỗi tick → mượt, không giật
        self._current += diff * 0.18
        self._var.set(self._current)
        self.lbl_pct.config(text=f"{int(self._current)}%")
        self._anim_id = self.after(20, self._animate_step)
