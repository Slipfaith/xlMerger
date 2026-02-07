# -*- coding: utf-8 -*-
from PySide6.QtCore import QObject, Signal, QSettings
from .translations import TRANSLATIONS

class I18N(QObject):
    language_changed = Signal(str)

    def __init__(self):
        super().__init__()
        self.settings = QSettings("xlMerger", "xlMerger")
        self.language = self.settings.value("language", "ru")
        if self.language not in TRANSLATIONS:
            self.language = "ru"
        self.translations = TRANSLATIONS.get(self.language, {})

    def load(self, lang: str):
        self.translations = TRANSLATIONS.get(lang, {})
        self.language = lang

    def translate(self, text: str) -> str:
        return self.translations.get(text, text)

    def set_language(self, lang: str):
        if lang != self.language:
            self.load(lang)
            self.settings.setValue('language', lang)
            self.language_changed.emit(lang)

i18n = I18N()

def tr(text: str) -> str:
    return i18n.translate(text)
