"""Измерение реальной вёрстки документа.

Документ держится открытым в Word и достраивается кусками. Word пересчитывает
вёрстку только от места вставки, поэтому измерение не дорожает с ростом
документа: ~0.1 с и на третьей странице, и на сто шестидесятой. Альтернатива —
сохранять документ и открывать заново на каждое измерение — стоит O(числа
страниц): около секунды на полусотне страниц и дальше линейно.
"""

import logging
import time
from pathlib import Path
from typing import Protocol

from docx import Document as read_document
from docx.oxml.ns import qn

from gost.timing import logged_duration

logger = logging.getLogger(__name__)

WD_ACTIVE_END_PAGE_NUMBER = 3
CONFIRM_CONVERSIONS = False

# Пустой документ — это один пустой абзац; он же и весь хвост зеркала на старте.
_TRAILING_IN_EMPTY_DOCUMENT = 1

# Документ, открытый только на чтение, Word верстает не так, как печатает: на
# 60 рисунках насчитывает 20 страниц вместо 30, на тексте — 40 вместо 37.
# Ошибка в обе стороны, и Repaginate() её не лечит. Сверка с экспортом в PDF —
# то есть с тем, что реально увидит читатель, — сходится только при открытии на
# запись. Документ мы всё равно закрываем без сохранения.
OPEN_READ_ONLY = False


class LayoutMirror(Protocol):
    """Зеркало собираемого документа: в него дописывают куски и меряют вёрстку.

    Зеркало повторяет документ элемент в элемент, поэтому его вёрстка — это
    вёрстка того, что получит читатель.
    """

    def append(self, chunk: Path) -> None:
        """Дописывает содержимое chunk в конец зеркала."""
        ...

    def undo(self) -> None:
        """Отменяет последний append, возвращая зеркало ровно в прежнее состояние."""
        ...

    def first_row_on_later_page(self, after: int) -> int | None:
        """Первая строка последней таблицы, ушедшая дальше страницы её первой строки.

        Args:
            after: Искать начиная с этой строки (0-based).

        Returns:
            Номер строки (0-based) или None, если таблица уместилась на страницу.
        """
        ...


class WordMirror:
    """Зеркало в реальном Word через COM.

    Требует Windows, установленный Word и pywin32 (extra «autosplit»).
    Держит Word и документ открытыми, поэтому используется только как
    контекстный менеджер — иначе процесс Word останется висеть.
    """

    def __init__(self, base: Path) -> None:
        """Args: base: Пустой документ с нужными стилями — с него начинается зеркало."""
        self.__base = base
        self.__word = None
        self.__document = None
        # Сколько пустых абзацев в хвосте зеркала. Считать их дешевле, чем
        # искать: перебор абзацев Word делает проходом от начала документа, и на
        # каждую вставку это стоило бы O(размера зеркала).
        self.__trailing = _TRAILING_IN_EMPTY_DOCUMENT
        self.__trailing_before_append = _TRAILING_IN_EMPTY_DOCUMENT

    def __enter__(self) -> "WordMirror":
        try:
            import win32com.client
        except ImportError as e:
            raise RuntimeError(
                'Для split_after="auto" нужен pywin32: pip install gost-docx[autosplit]. '
                "Либо задайте точки разрыва явно: split_after=[...]."
            ) from e

        try:
            # DispatchEx, а не Dispatch: Dispatch подключается к уже запущенному
            # Word пользователя, и тогда Quit() на выходе закрыл бы его вместе с
            # открытыми в нём документами. DispatchEx поднимает отдельный
            # процесс, который принадлежит только нам.
            self.__word = win32com.client.DispatchEx("Word.Application")
        except Exception as e:
            raise RuntimeError(
                'Для split_after="auto" нужен установленный Microsoft Word. '
                "Либо задайте точки разрыва явно: split_after=[...]."
            ) from e

        try:
            self.__word.Visible = False
            self.__document = self.__word.Documents.Open(
                str(self.__base.resolve()), CONFIRM_CONVERSIONS, OPEN_READ_ONLY,
            )
        except BaseException:
            # Упавший __enter__ не приводит к __exit__ — Word остался бы висеть.
            self.__close()
            raise
        return self

    def __exit__(self, *exc) -> None:
        self.__close()

    def __close(self) -> None:
        if self.__document is not None:
            self.__document.Close(False)
            self.__document = None
        if self.__word is not None:
            self.__word.Quit()
            self.__word = None

    def append(self, chunk: Path) -> None:
        """Дописывает кусок в конец зеркала.

        Word не даёт документу заканчиваться таблицей и приписывает за ней
        пустой абзац — и зеркалу, и каждому куску, который ей заканчивается. В
        документе таких абзацев нет, а в зеркале они вставали бы между кусками,
        по одному на таблицу. Всё, что ниже, съезжало бы вниз, и таблица мерялась
        бы не там, где окажется у читателя.

        Поэтому вставляем в первый из них: кусок встаёт выше, а вся их цепочка
        вытесняется вниз — под последнюю таблицу, где мерить уже нечего. Удалять
        их нельзя: без абзаца вставке не за что зацепиться, и Word сливает первый
        абзац куска с последним абзацем зеркала.
        """
        with logged_duration(logger, "Зеркало: дописан кусок %s", chunk.name):
            start = self.__live.Content.End - self.__trailing
            point = self.__live.Range(start, start)

            self.__trailing_before_append = self.__trailing
            point.InsertFile(str(chunk.resolve()))
            if _ends_with_table(chunk):
                self.__trailing += 1

    def undo(self) -> None:
        # Собственный Undo Word отменяет вставку точно. Считать позиции в Range
        # и удалять диапазон — нельзя: Word бережёт последний абзац документа и
        # оставляет его, зеркало разъезжается с документом, и вёрстка едет
        # вместе с ним.
        if not self.__live.Undo():
            raise RuntimeError("Word не смог отменить вставку в зеркало")
        self.__trailing = self.__trailing_before_append

    def first_row_on_later_page(self, after: int) -> int | None:
        table = self.__live.Tables(self.__live.Tables.Count)  # последняя — та, что мерим
        queries = 0

        def page_of(row: int) -> int:
            nonlocal queries
            queries += 1
            return table.Rows(row + 1).Range.Information(WD_ACTIVE_END_PAGE_NUMBER)

        start = time.perf_counter()
        row = _search(page_of, table.Rows.Count, after)
        logger.debug(
            "Зеркало: в таблице %d строк, граница найдена за %.3f с, запросов к "
            "вёрстке %d, первая строка на следующей странице: %s",
            table.Rows.Count, time.perf_counter() - start, queries, row,
        )
        return row

    @property
    def __live(self):
        if self.__document is None:
            raise RuntimeError("WordMirror используется вне контекстного менеджера")
        return self.__document


def _ends_with_table(chunk: Path) -> bool:
    """Заканчивается ли кусок таблицей — тогда Word припишет ему пустой абзац.

    Читаем сам файл: он размером с элемент, а спросить о том же Word значило бы
    пройти по всему зеркалу.
    """
    body = read_document(str(chunk)).element.body
    blocks = [child for child in body.iterchildren()
              if child.tag in (qn("w:p"), qn("w:tbl"))]
    return bool(blocks) and blocks[-1].tag == qn("w:tbl")


def _search(page_of, row_count: int, after: int) -> int | None:
    """Бинарный поиск первой строки на странице дальше первой.

    Строки верстаются сверху вниз, а cantSplit не даёт строке разорваться между
    страницами, поэтому номера страниц по строкам монотонно не убывают — значит
    поиск границы двоичный, а не перебором. Один запрос к вёрстке стоит ~30 мс и
    от размера документа не зависит.
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
