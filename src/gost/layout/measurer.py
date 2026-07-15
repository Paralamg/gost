"""Измерение реальной вёрстки документа."""

from pathlib import Path
from typing import Protocol

WD_ACTIVE_END_PAGE_NUMBER = 3


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

        document = self.__word.Documents.Open(str(path.resolve()), False, True)
        try:
            table = document.Tables(table_index + 1)  # COM-коллекции 1-based
            return _search(lambda row: _page_of(table, row), table.Rows.Count, after)
        finally:
            document.Close(False)


def _page_of(table, row: int) -> int:
    """Один запрос к вёрстке Word. Дорогой (~25 мс) — отсюда бинарный поиск."""
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
