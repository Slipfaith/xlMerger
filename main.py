import sys
from PySide6.QtWidgets import QApplication
from main_window import MainWindow
from utils.i18n import i18n

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
