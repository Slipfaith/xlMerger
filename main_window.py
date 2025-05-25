from PySide6.QtWidgets import QMainWindow, QMessageBox, QTabWidget
from PySide6.QtGui import QAction

from gui.file_processor_app import FileProcessorApp
from gui.limits_checker import LimitsChecker

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(self.tr("File Processor Application"))
        self.init_menu()

        # Создаем QTabWidget и добавляем вкладки
        self.tab_widget = QTabWidget()

        # Вкладка для обработки файлов (xlmerger)
        self.file_processor_widget = FileProcessorApp()
        self.tab_widget.addTab(self.file_processor_widget, self.tr("xlMerger"))

        # Вкладка для проверки лимитов
        self.limits_checker_widget = LimitsChecker()
        self.tab_widget.addTab(self.limits_checker_widget, self.tr("Лимит чек"))

        self.setCentralWidget(self.tab_widget)
        self.show()

    def init_menu(self):
        menubar = self.menuBar()

        file_menu = menubar.addMenu(self.tr("File"))
        help_menu = menubar.addMenu(self.tr("Help"))

        exit_action = QAction(self.tr("Exit"), self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        about_action = QAction(self.tr("About"), self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

    def show_about(self):
        QMessageBox.information(self, self.tr("About"), self.tr("Объединяй и проверяй\nVersion 03.15.2025"))
