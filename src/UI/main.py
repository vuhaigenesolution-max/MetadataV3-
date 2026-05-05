"""Entry point — launch single-root app."""
import os
import sys

_UI_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "src", "UI")
if _UI_DIR not in sys.path:
    sys.path.insert(0, _UI_DIR)

from app import App
from HomePage import HomePage


def main():
    app = App()
    app.show_page(HomePage)
    app.mainloop()


if __name__ == "__main__":
    main()
