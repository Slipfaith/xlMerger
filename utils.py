def excel_column_to_index(column: str) -> int:
    """
    Конвертация буквы столбца Excel в индекс столбца (A → 1, B → 2, ..., Z → 26, AA → 27).

    Args:
        column (str): Буквенное обозначение столбца.

    Returns:
        int: Индекс столбца (1-based).
    """
    index = 0
    for char in column.upper():
        index = index * 26 + (ord(char) - ord('A')) + 1
    return index
