import json
import os
from PySide6.QtCore import QObject, Signal, QSettings


class I18N(QObject):
    language_changed = Signal(str)

    def __init__(self):
        super().__init__()
        self.settings = QSettings("xlMerger", "xlMerger")
        self.language = self.settings.value("language", "ru")
        self.translations = {}
        self.load(self.language)

    def load(self, lang: str):
        path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'locales', f'{lang}.json')
        if os.path.exists(path):
            with open(path, 'r', encoding='utf-8') as f:
                self.translations = json.load(f)
        else:
            self.translations = {}
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
