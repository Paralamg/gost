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
wb.add_element(factory.create_text("Обычный абзац с **жирным** фрагментом."))
wb.add_element(factory.create_table(
    {"Имя": ["Алиса", "Боб"], "Возраст": [25, 30]},
    title="Пример таблицы",
))

wb.save(Path("report.docx"))
```

## Элементы документа

Все элементы создаются через `ElementFactory` — он же ведёт счётчики номеров.

| Метод | Что добавляет |
| --- | --- |
| `create_text(text)` | Абзац. Фрагменты в `**звёздочках**` становятся жирными. |
| `create_head(use_numbers, text, level)` | Заголовок уровня 0–4. Уровень 0 — по центру, прописными, без номера. |
| `create_image(path, alt)` | Рисунок с подписью «Рисунок N — alt». |
| `create_table(data, title, ...)` | Таблица с подписью «Таблица N — title». |
| `create_page_break(break_type)` | Разрыв страницы или раздела. |

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
