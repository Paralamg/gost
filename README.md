# gost-docx

Генерация документов Word (`.docx`), оформленных по ГОСТ 7.32-2017: Times New Roman 14 pt,
полуторный интервал, красная строка 1,25 см, сквозная нумерация заголовков, рисунков и таблиц,
подписи «Рисунок N — …» и «Таблица N — …», перенос длинных таблиц с надписью «Продолжение таблицы N».

Документ собирается из элементов, а нумерация проставляется автоматически — вручную считать
номера рисунков и таблиц не нужно.

## Установка

```bash
pip install gost-docx
```

Требуется Python 3.12 или новее.

Для автоподбора точек разрыва таблиц (`split_after="auto"`) нужен Windows с установленным Word:

```bash
pip install "gost-docx[autosplit]"
```

## Быстрый старт

```python
from pathlib import Path

from gost import WordBuilder
from gost.element_factory import ElementFactory

wb = WordBuilder()
factory = ElementFactory()

wb.add_element(factory.create_head(True, "Введение", 1))
wb.add_element(factory.create_text("Обычный абзац с **жирным** и *курсивным* фрагментом."))
wb.add_element(factory.create_table(
    {"Имя": ["Алиса", "Боб"], "Возраст": [25, 30]},
    title="Пример **таблицы**",
))

wb.save(Path("report.docx"))
```

## Элементы документа

Все элементы создаются через `ElementFactory` — он же ведёт счётчики номеров.

| Метод | Что добавляет |
| --- | --- |
| `create_text(text)` | Абзац. Поддерживает [инлайн-разметку](#инлайн-разметка). |
| `create_head(use_numbers, text, level)` | Заголовок уровня 0–4. Уровень 0 — по центру, прописными, без номера. |
| `create_image(path, alt)` | Рисунок с подписью «Рисунок N — alt». `alt` поддерживает [инлайн-разметку](#инлайн-разметка). |
| `create_table(data, title, ...)` | Таблица с подписью «Таблица N — title». `title` поддерживает [инлайн-разметку](#инлайн-разметка). |
| `create_page_break(break_type)` | Разрыв страницы или раздела. |

### Инлайн-разметка

В тексте абзаца и в подписях таблиц и рисунков (`title`, `alt`) работает облегчённый
markdown-синтаксис. Он разбирается на лету и превращается в оформленные фрагменты Word.

| Разметка в исходнике | Результат |
| --- | --- |
| `**текст**` | **жирный** |
| `*текст*` или `_текст_` | *курсив* |
| `***текст***` | ***жирный курсив*** |
| `__текст__` | подчёркнутый |
| `~` (одиночная тильда) | неразрывный пробел (`\xa0`) |

```python
wb.add_element(factory.create_text(
    "Значение **важно**: до *5* штук, см. рис.~1 и табл.~2."
))
wb.add_element(factory.create_table(
    data, title="Показатели за ***2024*** год",
))
```

Ограничения: экранирования нет — любая `~` становится неразрывным пробелом; одиночные `*`/`_`
без пары остаются как есть, но текст с двумя такими символами (`a * b * c`) может быть ошибочно
принят за курсив. Зачёркивание (`~~…~~`) не поддерживается — оно конфликтует с `~`.

### Нумерация

Режим нумерации задаётся при создании `IndexManager` и передаётся в фабрику:

```python
from gost.element_factory import ElementFactory
from gost.index.index_manager import IndexManager
from gost.index.index_type import IndexType

# Сквозная нумерация по всему документу: Таблица 1, Таблица 2, ...
factory = ElementFactory(IndexManager(index_type=IndexType.CONTINUOUS))

# Нумерация внутри главы: Таблица 1.1, Таблица 1.2, Таблица 2.1, ...
factory = ElementFactory(IndexManager(index_type=IndexType.CHAPTER_RELATIVE))

# С буквенным префиксом (приложения): Таблица А.1, Таблица А.2, ...
factory = ElementFactory(IndexManager(character="А"))
```

### Таблицы

```python
wb.add_element(factory.create_table(
    data,                        # Mapping[str, Sequence[Any]]: заголовок столбца -> значения
    title="Себестоимость",
    show_header=True,            # строка с именами столбцов
    show_row_numbers=True,       # столбец «№ п/п»
    show_column_numbers=True,    # строка с номерами столбцов
    split_after=[16],            # после каких строк переносить таблицу на новую страницу
))
```

Значения приводятся к строке через `str()`, поэтому числа форматируйте заранее:
`[f"{v:.2f}" for v in values]`.

`split_after="auto"` подбирает точки разрыва измерением: `save()` рендерит документ, спрашивает
у Word, на какой странице оказалась каждая строка, и ставит разрыв ровно по краю страницы.
Это работает только на Windows с установленным Word, требует `gost-docx[autosplit]` и заметно
медленнее обычной сборки.

## Стили

Оформление настраивается на двух уровнях: **глобально** (значения по умолчанию для всех
элементов) и **локально** (переопределение конкретного элемента). По умолчанию действует
`StyleSheet.gost()` — Times New Roman, полуторный интервал, красная строка 1,25 см и т.д.

Глобальный уровень задаётся при создании фабрики, как и `IndexManager`:

```python
from gost import WordBuilder, StyleSheet, Pt
from gost.element_factory import ElementFactory

style = StyleSheet.gost()
style.normal.font_size = Pt(13)          # обычный текст
style.table_text.font_name = "Arial"     # текст внутри таблиц
style.headings[0].font_size = Pt(16)     # заголовок уровня 0 (Heading 1)

factory = ElementFactory(style=style)    # копии стилей стемпятся в каждый элемент
```

Локально меняется свойство уже созданного элемента — на другие элементы это не влияет.
У таблицы два стиля: `caption_style` (подпись) и `text_style` (текст ячеек):

```python
table = factory.create_table(data, title="Смета")
table.text_style.font_size = Pt(10)      # только ячейки этой таблицы
table.caption_style.bold = True          # только подпись этой таблицы
wb.add_element(table)

text = factory.create_text("Абзац")
text.style.alignment = WD_ALIGN_PARAGRAPH.CENTER   # только этот абзац
```

Стиль каждого элемента — объект `ParagraphStyle` со свойствами: `font_name`, `font_size`,
`bold`, `italic`, `all_caps`, `color`, `alignment`, `line_spacing`, `first_line_indent`,
`space_before`, `space_after`. Свойства `StyleSheet`: `normal`, `headings` (список,
индекс = уровень заголовка 0–4), `table_caption`, `table_text`, `image_caption`. Типы
для значений (`Pt`, `Cm`, `RGBColor`, `WD_ALIGN_PARAGRAPH`, `WD_LINE_SPACING`)
реэкспортируются из `gost` — импортировать из `docx` не нужно.

> Фабрика копирует стили в момент `create_*`, поэтому глобальные правки в `StyleSheet`
> вносите **до** создания элементов.

## Примеры

В каталоге `examples/` лежат запускаемые скрипты со всеми стилями:

```bash
python examples/example.py            # все элементы и стили -> examples/example.docx
python examples/autosplit_example.py  # split_after="auto" (нужен Windows + Word)
```

## Разработка

Проект использует [uv](https://docs.astral.sh/uv/).

```bash
git clone git@github.com:Paralamg/gost.git
cd gost
uv sync
uv run python examples/example.py
```

## Публикация на PyPI

Сборка выполняется бэкендом `uv_build`, публикация — командой `uv publish`.

### 1. Подготовка

Перед выпуском поднимите версию в `pyproject.toml` (поле `version`) — PyPI не разрешает
повторно загрузить уже опубликованную версию, даже после удаления файла.

Получите API-токен: [pypi.org/manage/account/token](https://pypi.org/manage/account/token/)
(и отдельный — на [test.pypi.org](https://test.pypi.org/manage/account/token/) для тестовой площадки).
Токен начинается с `pypi-`. Передавайте его через переменную окружения, а не аргументом командной
строки — так он не попадёт в историю оболочки:

```powershell
# PowerShell
$env:UV_PUBLISH_TOKEN = "pypi-AgEIcHl..."
```

```bash
# bash
export UV_PUBLISH_TOKEN="pypi-AgEIcHl..."
```

### 2. Сборка

```bash
uv build
```

Артефакты появятся в `dist/`: `gost_docx-<версия>-py3-none-any.whl` и `gost_docx-<версия>.tar.gz`.
Если в `dist/` остались файлы прошлого выпуска, очистите каталог — `uv publish` загружает всё,
что там лежит:

```powershell
Remove-Item -Recurse -Force dist   # PowerShell
```

```bash
rm -rf dist                        # bash
```

### 3. Проверка на TestPyPI (рекомендуется)

```bash
uv publish --publish-url https://test.pypi.org/legacy/
```

Установка из TestPyPI — зависимости при этом берутся с основного PyPI:

```bash
pip install --index-url https://test.pypi.org/simple/ --extra-index-url https://pypi.org/simple/ gost-docx
```

### 4. Публикация

```bash
uv publish
```

Проверка установки из PyPI в чистом окружении:

```bash
uv run --with gost-docx --no-project -- python -c "from gost import WordBuilder; print(WordBuilder())"
```

### 5. Тег релиза

```bash
git tag v0.1.0
git push origin v0.1.0
```

## Лицензия

MIT — см. [LICENSE.txt](LICENSE.txt).
