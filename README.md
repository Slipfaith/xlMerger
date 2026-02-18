# xlMerger

Desktop utility for working with translation Excel files: merge, split, limit checks, and batch editing.

## English

### What the app can do

- `xlMerger` tab:
  - copy values from source files/folders into selected columns of a target workbook;
  - select multiple sheets and set header row per sheet;
  - map source files/folders to target columns;
  - map mismatched sheet names between source and target files;
  - choose copy mode:
    - sequential fill (skip empty source values),
    - row-by-row copy by row number;
  - optional style copy (preserve source formatting);
  - save/load mapping settings as JSON;
  - preview source files before processing;
  - write output as `<target>_out.xlsx` and keep copy logs.

- `Лимит чек` tab:
  - check text length limits in automatic or manual mode;
  - configure mappings `limit column -> text columns`;
  - configure manual checks for selected cells with upper/lower bounds;
  - highlight/report violations and save result as `<file>_checked.xlsx`.

- `xlSplit` tab:
  - split workbook by language columns for one or multiple sheets;
  - set source column, target columns, and extra columns per sheet;
  - preserve column widths/basic styles in output;
  - generate per-language files like `<file>_<src>-<target>.xlsx`.

- `xlCombine` tab:
  - merge multiple source files into one or more target workbooks;
  - configure per-file/per-sheet column mappings;
  - run merge in background with progress and output links;
  - save merged files as `<target>_merged.xlsx`.

- `xlCraft` tab:
  - batch edit many Excel files from a folder or file list;
  - operations:
    - rename headers (by column letter or header text),
    - fill a specific cell (optionally only when empty),
    - rename sheets,
    - clear a column (optionally clear formatting too);
  - apply operations to all files or a specific file scope;
  - preview sheets before execution and export edited copies to `*_upd` folder.

- General:
  - drag-and-drop for files/folders;
  - RU/EN interface switch.

### Run

```bash
pip install -r requirements.txt
python main.py
```

## Русский

### Возможности программы

- Вкладка `xlMerger`:
  - копирование значений из исходных файлов/папок в выбранные колонки целевого Excel;
  - выбор нескольких листов и строки заголовков для каждого листа;
  - сопоставление исходных файлов/папок с целевыми колонками;
  - сопоставление имен листов, если они отличаются;
  - два режима копирования:
    - последовательное заполнение (пропуск пустых значений),
    - копирование по номеру строки;
  - опциональное сохранение форматирования исходных ячеек;
  - сохранение/загрузка настроек сопоставления в JSON;
  - предпросмотр исходных файлов;
  - сохранение результата как `<target>_out.xlsx` и ведение логов копирования.

- Вкладка `Лимит чек`:
  - проверка ограничений длины текста в авто и ручном режимах;
  - настройка сопоставлений `колонка лимита -> текстовые колонки`;
  - ручная проверка выбранных ячеек с верхним/нижним порогом;
  - подсветка/отчет по нарушениям и сохранение `<file>_checked.xlsx`.

- Вкладка `xlSplit`:
  - разбиение книги по языковым колонкам для одного или нескольких листов;
  - настройка source, target и дополнительных колонок по каждому листу;
  - сохранение ширины колонок и базового форматирования;
  - генерация файлов вида `<file>_<src>-<target>.xlsx`.

- Вкладка `xlCombine`:
  - объединение данных из нескольких исходников в один или несколько целевых файлов;
  - настройка соответствий колонок по файлам и листам;
  - фоновая обработка с прогрессом и ссылками на результат;
  - сохранение результата как `<target>_merged.xlsx`.

- Вкладка `xlCraft`:
  - пакетная обработка множества Excel-файлов из папки или списка;
  - операции:
    - переименование заголовков (по букве колонки или по тексту),
    - заполнение конкретной ячейки (в т.ч. только пустых),
    - переименование листов,
    - очистка колонки (опционально с очисткой формата);
  - применение операций ко всем файлам или к выбранному файлу;
  - предпросмотр листов перед запуском и вывод копий в папку `*_upd`.

- Общее:
  - drag-and-drop файлов и папок;
  - переключение языка интерфейса RU/EN.

### Запуск

```bash
pip install -r requirements.txt
python main.py
```
