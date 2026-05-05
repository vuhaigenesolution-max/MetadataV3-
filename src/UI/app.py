"""Single-root App — controller switches frames in-place (no subprocess)."""
import os
import sys
import tkinter as tk

# ── Đảm bảo Logic có thể import được ──────────────────────
_SRC_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if _SRC_DIR not in sys.path:
    sys.path.insert(0, _SRC_DIR)

from _theme import BG, apply_ttk_styles


class App(tk.Tk):
    """Container chính — quản lý các Page (Frame) và chuyển qua lại nhanh."""

    def __init__(self):
        super().__init__()
        self.title("Excel Metadata Tool")
        self.geometry("900x650")
        self.configure(bg=BG)

        apply_ttk_styles()

        # Container giữ tất cả page (overlap nhau, chỉ hiện 1 cái)
        self.container = tk.Frame(self, bg=BG)
        self.container.pack(fill="both", expand=True)

        # Lazy-init: tạo frame khi page được show lần đầu
        self._frames: dict = {}
        self._current = None

    def show_page(self, page_class, **kwargs):
        """Chuyển sang page (Frame class). Tạo nếu lần đầu, ngược lại reuse."""
        # Hide current page
        if self._current is not None:
            self._current.pack_forget()

        # Lazy create
        if page_class not in self._frames:
            frame = page_class(self.container, controller=self, **kwargs)
            self._frames[page_class] = frame

        frame = self._frames[page_class]
        frame.pack(fill="both", expand=True)
        # Cho phép page tự refresh khi quay lại
        if hasattr(frame, "on_show"):
            frame.on_show()
        self._current = frame

    def reset_page(self, page_class):
        """Xóa cache 1 page → lần show kế tiếp tạo mới."""
        if page_class in self._frames:
            self._frames[page_class].destroy()
            del self._frames[page_class]


if __name__ == "__main__":
    from HomePage import HomePage  # noqa: E402

    app = App()
    app.show_page(HomePage)
    app.mainloop()
