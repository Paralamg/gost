"""Измерение реальной вёрстки документа."""

import logging
import time
from pathlib import Path
from typing import Protocol

from gost.timing import logged_duration

logger = logging.getLogger(__name__)

WD_ACTIVE_END_PAGE_NUMBER = 3
CONFIRM_CONVERSIONS = False

# Документ, открытый только на чтение, Word верстает не так, как печатает:
# на 60 рисунках он насчитывает 20 страниц вместо 30, на тексте — 40 вместо 37.
# Ошибка в обе стороны, и Repaginate() её не лечит; сверка с экспортом в PDF —
# то есть с тем, что реально увидит читатель, — сходится только при открытии на
# запись. Документ мы всё равно закрываем без сохранения.
OPEN_READ_ONLY = False


class PageMeasurer(Protocol):
    """Сообщает, где таблица переходит на следующую страницу."""

    def first_row_on_later_page(self, path: Path, table_index: int, after: int) -> int | None:
        """Первая строка таблицы, оказавшаяся дальше страницы её первой строки.

        Args:
            table_index: Номер таблицы в документе (0-based).
            after: Искать начиная с этой строки (0-based).

        Returns:
            Номер строки (0-based) или None, если таблица уместилась на страницу.
        """
        ...


class WordMeasurer:
    """Измеряет вёрстку через COM реального Word.

    Требует Windows, установленный Word и pywin32 (extra «autosplit»).
    Держит Word открытым между измерениями, поэтому используется только как
    контекстный менеджер — иначе процесс Word останется висеть.
    """

    def __init__(self) -> None:
        self.__word = None

    def __enter__(self) -> "WordMeasurer":
        try:
            import win32com.client
        except ImportError as e:
            raise RuntimeError(
                'Для split_after="auto" нужен pywin32: pip install gost-docx[autosplit]. '
                "Либо задайте точки разрыва явно: split_after=[...]."
            ) from e

        try:
            self.__word = win32com.client.Dispatch("Word.Application")
        except Exception as e:
            raise RuntimeError(
                'Для split_after="auto" нужен установленный Microsoft Word. '
                "Либо задайте точки разрыва явно: split_after=[...]."
            ) from e

        self.__word.Visible = False
        return self

    def __exit__(self, *exc) -> None:
        if self.__word is not None:
            self.__word.Quit()
            self.__word = None

    def first_row_on_later_page(self, path: Path, table_index: int, after: int) -> int | None:
        if self.__word is None:
            raise RuntimeError("WordMeasurer используется вне контекстного менеджера")

        with logged_duration(logger, "Word открыл %s", path.name):
            document = self.__word.Documents.Open(
                str(path.resolve()), CONFIRM_CONVERSIONS, OPEN_READ_ONLY,
            )
        try:
            table = document.Tables(table_index + 1)  # COM-коллекции 1-based
            queries = 0

            def page_of(row: int) -> int:
                nonlocal queries
                queries += 1
                return _page_of(table, row)

            start = time.perf_counter()
            row = _search(page_of, table.Rows.Count, after)
            logger.debug(
                "Таблица #%d в документе (строк %d): граница найдена за %.3f с, "
                "запросов к вёрстке %d, первая строка на следующей странице: %s",
                table_index, table.Rows.Count, time.perf_counter() - start, queries, row,
            )
            return row
        finally:
            document.Close(False)


def _page_of(table, row: int) -> int:
    """Один запрос к вёрстке Word. Дорогой (~30 мс) — отсюда бинарный поиск.

    Первый запрос после открытия документа заставляет Word пересчитать вёрстку
    целиком и стоит на порядок больше: ~1 с на полусотне страниц, дальше растёт
    линейно с их числом. Остальные запросы идут по готовой вёрстке и от размера
    документа не зависят.
    """
    return table.Rows(row + 1).Range.Information(WD_ACTIVE_END_PAGE_NUMBER)


def _search(page_of, row_count: int, after: int) -> int | None:
    """Бинарный поиск первой строки на странице дальше первой.

    Строки верстаются сверху вниз, а cantSplit не даёт строке разорваться между
    страницами, поэтому номера страниц по строкам монотонно не убывают — значит
    поиск границы двоичный, а не перебором.
    """
    first_page = page_of(0)
    if page_of(row_count - 1) == first_page:
        return None  # таблица целиком на одной странице

    lo, hi = after, row_count - 1
    while lo < hi:
        mid = (lo + hi) // 2
        if page_of(mid) == first_page:
            lo = mid + 1
        else:
            hi = mid
    return lo
