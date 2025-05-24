def excel_column_to_index(column):
    """Конвертация буквы столбца Excel в индекс столбца."""
    index = 0
    for char in column.upper():
        index = index * 26 + (ord(char) - ord('A')) + 1
    return index
