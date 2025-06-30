from PySide6.QtWidgets import QMainWindow, QMessageBox, QTabWidget
from PySide6.QtGui import QAction
from utils.i18n import tr, i18n
from utils.updater import check_for_update

from gui.file_processor_app import FileProcessorApp
from gui.limits_checker import LimitsChecker
from gui.split_tab import SplitTab
from PySide6.QtGui import QIcon

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(tr("xlMerger: объединяй и проверяй"))
        self.setWindowIcon(QIcon(r"C:\Users\yanismik\Desktop\PythonProject1\xlM_2.0\xlM2.0.ico"))  # <- иконка здесь

        # Создаем QTabWidget и добавляем вкладки
        self.tab_widget = QTabWidget()

        # Вкладка для обработки файлов (xlmerger)
        self.file_processor_widget = FileProcessorApp()
        self.tab_widget.addTab(self.file_processor_widget, tr("xlMerger"))

        # Вкладка для проверки лимитов
        self.limits_checker_widget = LimitsChecker()
        self.tab_widget.addTab(self.limits_checker_widget, tr("Лимит чек"))

        # Вкладка для разделения Excel
        self.split_tab_widget = SplitTab()
        # Tab renamed from "Разделение" to "xlSpliter"
        self.tab_widget.addTab(self.split_tab_widget, tr("xlSpliter"))

        self.setCentralWidget(self.tab_widget)
        self.init_menu()
        self.show()
        # Automatic update check at startup
        self.check_updates(auto=True)

    def retranslate_ui(self):
        self.setWindowTitle(tr("xlMerger: объединяй и проверяй"))
        self.file_menu.setTitle(tr("File"))
        self.help_menu.setTitle(tr("Help"))
        self.exit_action.setText(tr("Exit"))
        self.about_action.setText(tr("About"))
        self.language_menu.setTitle(tr("Language"))
        self.tab_widget.setTabText(0, tr("xlMerger"))
        self.tab_widget.setTabText(1, tr("Лимит чек"))
        self.tab_widget.setTabText(2, tr("xlSpliter"))
        self.update_action.setText(tr("Check for Updates"))

    def init_menu(self):
        menubar = self.menuBar()

        file_menu = menubar.addMenu(tr("File"))
        help_menu = menubar.addMenu(tr("Help"))
        self.file_menu = file_menu
        self.help_menu = help_menu

        exit_action = QAction(tr("Exit"), self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        self.exit_action = exit_action

        about_action = QAction(tr("About"), self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

        update_action = QAction(tr("Check for Updates"), self)
        update_action.triggered.connect(self.check_updates)
        help_menu.addAction(update_action)
        self.update_action = update_action

        language_menu = menubar.addMenu(tr("Language"))
        lang_en = QAction("English", self)
        lang_ru = QAction("Русский", self)
        lang_en.triggered.connect(lambda: i18n.set_language('en'))
        lang_ru.triggered.connect(lambda: i18n.set_language('ru'))
        language_menu.addAction(lang_en)
        language_menu.addAction(lang_ru)
        self.language_menu = language_menu
        self.lang_en_action = lang_en
        self.lang_ru_action = lang_ru

        self.about_action = about_action
        i18n.language_changed.connect(self.retranslate_ui)
        self.retranslate_ui()

    def show_about(self):
        QMessageBox.information(self, tr("About"), tr("Объединяй и проверяй\nVersion 1.0 24.06.2025\nslipfaith"))

    def check_updates(self, auto: bool = False):
        """Check for application updates using PyUpdater."""
        check_for_update(self, auto=auto)
