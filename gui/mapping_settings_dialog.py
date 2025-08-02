from PySide6.QtWidgets import QDialog, QVBoxLayout
from PySide6.QtCore import Signal
from utils.i18n import tr
from gui.pages.match_page import MatchPage


class MappingSettingsDialog(QDialog):
    """Separate window to configure file/folder-column mapping."""

    saveClicked = Signal(dict)
    loadClicked = Signal()

    def __init__(
        self,
        folder_path,
        selected_files,
        selected_sheets,
        columns,
        file_to_column=None,
        folder_to_column=None,
        preserve_formatting=False,
        parent=None,
    ):
        super().__init__(parent)
        self.setWindowTitle(tr("Настройки сопоставления"))
        self.match_page = MatchPage(
            folder_path,
            selected_files,
            selected_sheets,
            columns,
            file_to_column=file_to_column,
            folder_to_column=folder_to_column,
            preserve_formatting=preserve_formatting,
        )

        layout = QVBoxLayout(self)
        layout.addWidget(self.match_page)

        self.match_page.backClicked.connect(self.reject)
        self.match_page.nextClicked.connect(self._on_next)
        self.match_page.saveClicked.connect(self._on_save_clicked)
        self.match_page.loadClicked.connect(lambda: self.loadClicked.emit())

        self._file_to_column = {}
        self._folder_to_column = {}
        self._preserve_formatting = False

        self.setStyleSheet(
            """
            QDialog { background-color: #f0f0f0; }
            QLabel { font-size: 14px; }
            QPushButton { padding: 6px 12px; border-radius: 4px; background-color: #3498db; color: white; }
            QPushButton:hover { background-color: #2980b9; }
            QComboBox { padding: 4px; }
            """
        )

    def _on_save_clicked(self):
        self.saveClicked.emit(self.match_page.get_current_mapping())

    def _on_next(self, file_to_column, folder_to_column, preserve_formatting):
        self._file_to_column = file_to_column
        self._folder_to_column = folder_to_column
        self._preserve_formatting = preserve_formatting
        self.accept()

    def get_mapping(self):
        return self._file_to_column, self._folder_to_column, self._preserve_formatting

    def apply_mapping(self, mapping):
        self.match_page.apply_mapping(mapping, mapping)
