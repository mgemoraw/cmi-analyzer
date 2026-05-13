# app/main.py

import sys

from PySide6.QtWidgets import QApplication

import qdarktheme

from gui.main_window import MainWindow


def main():

    app = QApplication(sys.argv)

    # ==========================================
    # DARK THEME
    # ==========================================

    qdarktheme.setup_theme(
        "dark"
    )

    # ==========================================
    # MAIN WINDOW
    # ==========================================

    window = MainWindow()

    window.show()

    # ==========================================
    # START APP
    # ==========================================

    sys.exit(app.exec())


if __name__ == "__main__":

    main()