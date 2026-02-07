# -*- coding: utf-8 -*-
import os
from PySide6.QtWidgets import QLineEdit, QFileDialog
from PySide6.QtCore import Signal

class DragDropLineEdit(QLineEdit):
    filesSelected = Signal(list)       # Для поля "Папка переводов" (эксели)
    folderSelected = Signal(str)       # Для поля "Папка переводов" (папка)
    fileSelected = Signal(str)         # Для поля "Файл Excel" (один эксель)

    def __init__(self, mode='files_or_folder', parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setReadOnly(True)
        self.mode = mode  # 'files_or_folder' или 'file' или 'files'
        self.setStyleSheet(
            """
            QLineEdit {
                border: 2px dashed #aaa;
                border-radius: 6px;
                padding: 6px;
                background: #fafafa;
            }
            QLineEdit:hover {
                background: #f0f0f0;
            }
            """
        )

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        paths = [url.toLocalFile() for url in urls]
        files = [p for p in paths if os.path.isfile(p) and p.lower().endswith(('.xlsx', '.xls'))]
        folders = [p for p in paths if os.path.isdir(p)]

        # Для поля "Файл Excel" — только один файл
        if self.mode == 'file':
            if len(files) == 1 and not folders:
                self.setText(self._short_name(files[0]))
                self.fileSelected.emit(files[0])
            else:
                self.clear()

        # Для поля, принимающего только файлы (несколько)
        elif self.mode == 'files':
            if files and not folders:
                self.setText('; '.join([self._short_name(f) for f in files]))
                self.filesSelected.emit(files)
            else:
                self.clear()

        # Для поля "Папка переводов" — либо папка, либо только эксели
        elif self.mode == 'files_or_folder':
            if files and not folders:
                self.setText('; '.join([self._short_name(f) for f in files]))
                self.filesSelected.emit(files)
            elif folders and not files and len(folders) == 1:
                self.setText(folders[0])
                self.folderSelected.emit(folders[0])
            else:
                self.clear()  # Смешивание файлов/папок — ничего не делаем

    def mouseDoubleClickEvent(self, event):
        if self.mode == 'files_or_folder':
            dlg = QFileDialog(self)
            dlg.setFileMode(QFileDialog.ExistingFiles)
            dlg.setNameFilter("Excel файлы (*.xlsx *.xls)")
            if dlg.exec():
                files = dlg.selectedFiles()
                if files:
                    self.setText('; '.join([self._short_name(f) for f in files]))
                    self.filesSelected.emit(files)
            else:
                folder = QFileDialog.getExistingDirectory(self, "Выбери папку")
                if folder:
                    self.setText(folder)
                    self.folderSelected.emit(folder)
        elif self.mode == 'file':
            file, _ = QFileDialog.getOpenFileName(self, "Выбери файл", '', "Excel файлы (*.xlsx *.xls)")
            if file:
                self.setText(self._short_name(file))
                self.fileSelected.emit(file)
        elif self.mode == 'files':
            dlg = QFileDialog(self)
            dlg.setFileMode(QFileDialog.ExistingFiles)
            dlg.setNameFilter("Excel файлы (*.xlsx *.xls)")
            if dlg.exec():
                files = dlg.selectedFiles()
                if files:
                    self.setText('; '.join([self._short_name(f) for f in files]))
                    self.filesSelected.emit(files)

    def _short_name(self, path, n=5):
        name = os.path.basename(path)
        return name if len(name) <= 2 * n else f"{name[:n]}...{name[-n:]}"
