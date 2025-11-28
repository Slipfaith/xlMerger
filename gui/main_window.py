from PySide6.QtWidgets import (
    QMainWindow,
    QMessageBox,
    QTabWidget,
    QStackedWidget,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
)
from PySide6.QtGui import QAction, QIcon, QFont, QPixmap, QPainter
from PySide6.QtCore import Qt
from utils.i18n import tr, i18n
from __init__ import __version__

from .file_processor_app import FileProcessorApp
from .limits_checker import LimitsChecker
from .split_tab import SplitTab
from .merge_tab import MergeTab
from .excel_builder_tab import ExcelBuilderTab


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setup_ui()
        self.init_menu()
        self.show()

    def setup_ui(self):
        self.setWindowTitle(tr("xlMerger: объединяй и проверяй"))
        self.setWindowIcon(self._get_app_icon())
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

        self.main_screen = QWidget()
        main_layout = QVBoxLayout(self.main_screen)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.addWidget(self.tab_widget)

        self.excel_builder_widget = ExcelBuilderTab()
        self.builder_container = self._create_builder_page()

        self.stacked_widget = QStackedWidget()
        self.stacked_widget.addWidget(self.main_screen)
        self.stacked_widget.addWidget(self.builder_container)

        self.setCentralWidget(self.stacked_widget)
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
        self.builder_action.setText(tr("xlCraft"))
        self.builder_action.setToolTip(tr("Конструктор Excel"))
        self.back_button.setText(self._get_back_button_text())

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

        self.builder_action = QAction(tr("xlCraft"), self)
        self.builder_action.triggered.connect(self.show_builder_page)
        menubar.addAction(self.builder_action)

        self.about_action = about_action
        i18n.language_changed.connect(self.retranslate_ui)
        self.retranslate_ui()

    def _create_builder_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)

        header = QHBoxLayout()
        self.back_button = QPushButton(self._get_back_button_text())
        self.back_button.clicked.connect(self.show_main_screen)
        header.addWidget(self.back_button)
        header.addStretch()

        layout.addLayout(header)
        layout.addWidget(self.excel_builder_widget)
        return page

    def show_main_screen(self):
        self.stacked_widget.setCurrentWidget(self.main_screen)

    def show_builder_page(self):
        self.stacked_widget.setCurrentWidget(self.builder_container)

    def _get_settings_icon(self) -> QIcon:
        icon = QIcon.fromTheme("settings")
        if not icon.isNull():
            return icon

        pixmap = QPixmap(32, 32)
        pixmap.fill(Qt.transparent)

        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setPen(Qt.NoPen)
        painter.setBrush(Qt.black)
        painter.setFont(QFont(self.font().family(), 20))
        painter.drawText(pixmap.rect(), Qt.AlignCenter, "⚙")
        painter.end()

        return QIcon(pixmap)

    def _get_app_icon(self) -> QIcon:
        icon = QIcon.fromTheme("spreadsheet")
        if not icon.isNull():
            return icon

        pixmap = QPixmap(64, 64)
        pixmap.fill(Qt.transparent)

        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setPen(Qt.NoPen)
        painter.setBrush(Qt.black)
        painter.setFont(QFont(self.font().family(), 28, QFont.Bold))
        painter.drawText(pixmap.rect(), Qt.AlignCenter, "XL")
        painter.end()

        return QIcon(pixmap)

    def _get_back_button_text(self) -> str:
        return "←"

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
