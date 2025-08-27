from PySide6.QtWidgets import QMainWindow, QMessageBox, QTabWidget
from PySide6.QtGui import QAction, QIcon, QFont
from PySide6.QtCore import Qt
from utils.i18n import tr, i18n
from utils.updater import check_for_update
from __init__ import __version__

from .file_processor_app import FileProcessorApp
from .limits_checker import LimitsChecker
from .split_tab import SplitTab
from .merge_tab import MergeTab


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setup_ui()
        self.init_menu()
        self.show()
        self.check_updates(auto=True)

    def setup_ui(self):
        self.setWindowTitle(tr("xlMerger: объединяй и проверяй"))
        self.setWindowIcon(QIcon(r"C:\Users\yanismik\Desktop\PythonProject1\xlM_2.0\xlM2.0.ico"))
        self.setMinimumSize(650, 450)

        self.tab_widget = QTabWidget()

        self.file_processor_widget = FileProcessorApp()
        self.tab_widget.addTab(self.file_processor_widget, tr("xlMerger"))

        self.limits_checker_widget = LimitsChecker()
        self.tab_widget.addTab(self.limits_checker_widget, tr("Лимит чек"))

        self.split_tab_widget = SplitTab()
        self.tab_widget.addTab(self.split_tab_widget, tr("xlSpliter"))

        self.merge_tab_widget = MergeTab()
        self.tab_widget.addTab(self.merge_tab_widget, tr("Объединить"))

        self.setCentralWidget(self.tab_widget)
        self.apply_modern_style()

    def apply_modern_style(self):
        self.setStyleSheet("""
            QMainWindow {
                background-color: #fafafa;
                color: #2c2c2c;
            }

            QTabWidget::pane {
                border: 1px solid #e0e0e0;
                border-radius: 6px;
                background-color: #ffffff;
                margin-top: 2px;
            }

            QTabWidget::tab-bar {
                alignment: left;
            }

            QTabBar::tab {
                background-color: #f5f5f5;
                color: #666666;
                padding: 12px 24px;
                margin-right: 2px;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                font-weight: 500;
                font-size: 14px;
                min-width: 100px;
                border: 1px solid #e0e0e0;
                border-bottom: none;
            }

            QTabBar::tab:selected {
                background-color: #ffffff;
                color: #2c2c2c;
                border-color: #e0e0e0;
                font-weight: 600;
            }

            QTabBar::tab:hover:!selected {
                background-color: #eeeeee;
                color: #424242;
            }

            QMenuBar {
                background-color: #ffffff;
                color: #2c2c2c;
                border-bottom: 1px solid #e0e0e0;
                padding: 4px 0px;
                font-size: 14px;
            }

            QMenuBar::item {
                background-color: transparent;
                padding: 8px 16px;
                border-radius: 4px;
                margin: 2px 4px;
            }

            QMenuBar::item:selected {
                background-color: #f0f0f0;
                color: #1976d2;
            }

            QMenu {
                background-color: #ffffff;
                border: 1px solid #e0e0e0;
                border-radius: 6px;
                padding: 4px;
                font-size: 14px;
                box-shadow: 0px 4px 12px rgba(0, 0, 0, 0.1);
            }

            QMenu::item {
                padding: 8px 20px;
                border-radius: 4px;
                margin: 1px;
            }

            QMenu::item:selected {
                background-color: #e3f2fd;
                color: #1976d2;
            }

            QMenu::separator {
                height: 1px;
                background-color: #e0e0e0;
                margin: 4px 0px;
            }
        """)

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
        self.tab_widget.setTabText(3, tr("Объединить"))
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
        info_text = (
            f"{tr('Объединяй и проверяй')}\n{tr('Version')} {__version__}\nslipfaith"
        )
        msgbox = QMessageBox(self)
        msgbox.setWindowTitle(tr("About"))
        msgbox.setText(info_text)
        msgbox.setIcon(QMessageBox.Information)
        msgbox.setStyleSheet("""
            QMessageBox {
                background-color: #ffffff;
                color: #2c2c2c;
                font-size: 14px;
            }
            QMessageBox QPushButton {
                background-color: #1976d2;
                color: white;
                border: none;
                padding: 8px 24px;
                border-radius: 4px;
                font-weight: 500;
                min-width: 80px;
            }
            QMessageBox QPushButton:hover {
                background-color: #1565c0;
            }
            QMessageBox QPushButton:pressed {
                background-color: #0d47a1;
            }
        """)
        msgbox.exec()

    def check_updates(self, auto: bool = False):
        check_for_update(self, auto=auto)