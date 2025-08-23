import sys
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QIcon
from gui.main_window import MainWindow

ICON_PATH = r"E:\PythonProjects\01_icos\xlM2.0.ico"

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon(ICON_PATH))   # иконка для панели задач и диалогов

    window = MainWindow()
    window.setWindowIcon(QIcon(ICON_PATH))  # иконка для титульной строки окна
    window.show()

    sys.exit(app.exec())
