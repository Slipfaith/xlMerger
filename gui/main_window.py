# -*- coding: utf-8 -*-
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
from pathlib import Path
from .style_system import set_button_variant

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
        self.tab_widget.addTab(self.split_tab_widget, "xlSplit")

        self.merge_tab_widget = MergeTab()
        self.tab_widget.addTab(self.merge_tab_widget, "xlCombine")

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
        # Global style is applied on QApplication level.
        return

    def retranslate_ui(self):
        self.setWindowTitle(tr("xlMerger: объединяй и проверяй"))
        self.file_menu.setTitle(tr("File"))
        self.help_menu.setTitle(tr("Help"))
        self.exit_action.setText(tr("Exit"))
        self.about_action.setText(tr("About"))
        self.language_menu.setTitle(tr("Language"))
        self.tab_widget.setTabText(0, tr("xlMerger"))
        self.tab_widget.setTabText(1, tr("Лимит чек"))
        self.tab_widget.setTabText(2, "xlSplit")
        self.tab_widget.setTabText(3, "xlCombine")
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
        set_button_variant(self.back_button, "secondary")
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
        local_icon_path = Path(__file__).resolve().parent.parent / "xlM2.0.ico"
        if local_icon_path.exists():
            local_icon = QIcon(str(local_icon_path))
            if not local_icon.isNull():
                return local_icon

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
        msgbox.exec()
