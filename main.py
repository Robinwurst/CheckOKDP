# main.py
import sys
from PySide6.QtWidgets import QApplication
from ui import MainWindow, apply_app_theme


def main() -> None:
    app = QApplication(sys.argv)
    apply_app_theme(app)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
