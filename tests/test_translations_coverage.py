# -*- coding: utf-8 -*-
import ast
from pathlib import Path

from utils.translations import TRANSLATIONS


def _collect_tr_literals() -> set[str]:
    root = Path(__file__).resolve().parents[1]
    code_paths = list((root / "gui").glob("*.py")) + list((root / "core").glob("*.py")) + [root / "main.py"]
    literals: set[str] = set()

    for path in code_paths:
        tree = ast.parse(path.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            if (
                isinstance(node, ast.Call)
                and isinstance(node.func, ast.Name)
                and node.func.id == "tr"
                and node.args
                and isinstance(node.args[0], ast.Constant)
                and isinstance(node.args[0].value, str)
            ):
                literals.add(node.args[0].value)
    return literals


def test_translations_cover_all_tr_literals_for_ru_and_en():
    literals = _collect_tr_literals()
    missing = {
        lang: sorted(key for key in literals if key not in TRANSLATIONS.get(lang, {}))
        for lang in ("ru", "en")
    }
    assert not missing["ru"], f"Missing RU translations: {missing['ru']}"
    assert not missing["en"], f"Missing EN translations: {missing['en']}"
