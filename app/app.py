import os
import sys

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

from app.logger import get_logger
from app.updater import check_for_updates

log = get_logger(__name__)
from app.gui.windows.main_window import MainWindow


def main():
    log.info("Launching MainWindow")
    app = MainWindow()
    check_for_updates(app)
    app.mainloop()
    log.info("Application closed")


if __name__ == "__main__":
    main()
