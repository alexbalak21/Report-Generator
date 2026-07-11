import os
import sys

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

from app.updater import check_for_updates
from app.gui.windows.main_window import MainWindow


def main():
    app = MainWindow()
    check_for_updates(app)  # pass the window so the dialog can attach to it
    app.mainloop()


if __name__ == "__main__":
    main()
