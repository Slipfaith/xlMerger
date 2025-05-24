import os
from PySide6.QtWidgets import QLineEdit, QFileDialog
from PySide6.QtGui import QAction
from PySide6.QtCore import Qt, Signal

class DragDropLineEdit(QLineEdit):
    pathSelected = Signal(str)
    cleared = Signal()

    def __init__(self, mode='file', parent=None):
        super().__init__(parent)
        self.mode = mode  # 'file' or 'folder'
        self.setAcceptDrops(True)
        self.setReadOnly(False)
        self._add_clear_action()
        self._update_placeholder()

    def _add_clear_action(self):
        clear_action = QAction("Очистить", self)
        clear_action.setIconVisibleInMenu(False)
        clear_action.triggered.connect(self.clear)
        self.addAction(clear_action, QLineEdit.TrailingPosition)

    def _update_placeholder(self):
        if self.mode == 'file':
            self.setPlaceholderText("Файл Excel (двойной клик или drag&drop)")
        else:
            self.setPlaceholderText("Папка (двойной клик или drag&drop)")

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if not urls:
            return
        path = urls[0].toLocalFile()
        if self.mode == 'file' and os.path.isfile(path):
            self.setText(path)
            self.pathSelected.emit(path)
        elif self.mode == 'folder' and os.path.isdir(path):
            self.setText(path)
            self.pathSelected.emit(path)

    def mouseDoubleClickEvent(self, event):
        if self.mode == 'file':
            path, _ = QFileDialog.getOpenFileName(self, "Выберите файл Excel", "", "Excel файлы (*.xlsx *.xls)")
            if path:
                self.setText(path)
                self.pathSelected.emit(path)
        else:
            path = QFileDialog.getExistingDirectory(self, "Выберите папку")
            if path:
                self.setText(path)
                self.pathSelected.emit(path)

    def clear(self):
        super().clear()
        self.cleared.emit()
