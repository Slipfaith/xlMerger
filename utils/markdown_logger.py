"""Markdown logger for detailed copy sessions.

This module collects detailed information about copy operations
and saves them into a Markdown file with rich formatting so that
users can easily trace where each value came from and where it was
placed.
"""

from __future__ import annotations

import datetime
import os
from typing import Any, Dict, List


class MarkdownLogger:
    """Collects copy operations and renders them into Markdown."""

    def __init__(self, markdown_file: str = "copy_report.md") -> None:
        self.markdown_file = markdown_file
        self.entries: List[Dict[str, Any]] = []
        self.session_started = datetime.datetime.now()

    def log_copy(
        self,
        source_file: str,
        source_sheet: str,
        source_cell: str,
        target_file: str,
        target_sheet: str,
        target_cell: str,
        value: Any,
    ) -> None:
        """Store a copy event for later Markdown rendering."""

        self.entries.append(
            {
                "source_file": source_file or "—",
                "source_sheet": source_sheet or "—",
                "source_cell": source_cell or "—",
                "target_file": target_file or "—",
                "target_sheet": target_sheet or "—",
                "target_cell": target_cell or "—",
                "value": self._format_value(value),
            }
        )

    def save(self) -> None:
        """Persist the current session to disk.

        Even если не было успешных копирований, создаём файл отчёта и фиксируем
        пустую сессию, чтобы у пользователя оставался след запуска.
        """

        os.makedirs(os.path.dirname(self.markdown_file) or ".", exist_ok=True)

        session_title = self.session_started.strftime("%Y-%m-%d %H:%M:%S")
        with open(self.markdown_file, "a", encoding="utf-8") as md_file:
            md_file.write(f"## Сессия копирования от {session_title}\n\n")
            if not self.entries:
                md_file.write("Нет успешных копирований за эту сессию.\n\n")
            else:
                md_file.write("| Источник | Приёмник | Текст |\n")
                md_file.write("| --- | --- | --- |\n")
                for entry in self.entries:
                    source = self._format_endpoint(entry, "source")
                    target = self._format_endpoint(entry, "target")
                    text_block = self._format_text_block(entry["value"])
                    md_file.write(f"| {source} | {target} | {text_block} |\n")
                md_file.write("\n")

        self.entries.clear()
        self.session_started = datetime.datetime.now()

    def _format_value(self, value: Any) -> str:
        if value is None:
            return ""
        return str(value).replace("\r\n", "\n")

    def _format_endpoint(self, entry: Dict[str, Any], prefix: str) -> str:
        file_key = f"{prefix}_file"
        sheet_key = f"{prefix}_sheet"
        cell_key = f"{prefix}_cell"
        file_path = entry.get(file_key, "—")
        sheet = entry.get(sheet_key, "—")
        cell = entry.get(cell_key, "—")
        return f"`{file_path}`<br>`{sheet}!{cell}`"

    def _format_text_block(self, text: str) -> str:
        safe_text = (text or "").replace("|", "\\|")
        return f"```\n{safe_text}\n```"
