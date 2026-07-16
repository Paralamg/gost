"""Подбор точек разрыва таблиц по измеренной вёрстке.

Точку разрыва нельзя вычислить — высота строки зависит от переноса слов,
а он от метрик шрифта, ширины ячейки и позиции таблицы на странице. Поэтому
здесь вёрстка измеряется: документ рендерится, Word сообщает, на какой
странице оказалась каждая строка, и разрыв ставится ровно туда.

Разрывы ставятся по одному. Разрыв сдвигает вниз всё, что за ним следует,
поэтому за один проход можно доверять только первому измеренному разрыву —
остальные измерены по вёрстке, которой уже не будет.
"""

import logging
from collections import deque
from collections.abc import Callable, Sequence
from pathlib import Path

from docx.document import Document

from gost.elements.element import ElementBase
from gost.elements.table import Table
from gost.layout.measurer import PageMeasurer
from gost.timing import logged_duration

logger = logging.getLogger(__name__)

Build = Callable[[], tuple[Document, list[range]]]

Pending = deque[tuple[int, Table]]


def resolve_auto_splits(
        elements: Sequence[ElementBase],
        build: Build,
        measurer: PageMeasurer,
        probe: Path,
) -> Document:
    """Проставляет точки разрыва таблицам с split_after="auto".

    Args:
        elements: Элементы документа в порядке вывода.
        build: Собирает документ заново и возвращает его вместе с диапазонами
            индексов таблиц, которые сгенерировал каждый элемент.
        measurer: Чем измерять вёрстку.
        probe: Куда сохранять промежуточный документ для измерения.

    Returns:
        Документ, соответствующий подобранным точкам разрыва.
    """
    pending: Pending = deque(
        (i, element) for i, element in enumerate(elements)
        if isinstance(element, Table) and element.auto_split
    )
    max_passes = _max_passes(pending)
    logger.debug("Автоподбор разрывов: таблиц %d, предел проходов %d",
                 len(pending), max_passes)

    for pass_no in range(1, max_passes + 1):
        document, spans = build()
        if not pending:
            return document

        with logged_duration(logger, "Проход %d: пробник сохранён", pass_no):
            document.save(str(probe))

        if not _place_next_split(pending, spans, measurer, probe):
            logger.info("Автоподбор разрывов завершён за %d проходов", pass_no)
            return document

    raise RuntimeError(
        f"Не удалось подобрать точки разрыва за {max_passes} проходов. "
        "Задайте split_after явно."
    )


def _max_passes(pending: Pending) -> int:
    """Предел числа проходов — страховка от зацикливания, а не бюджет.

    За проход ставится не больше одного разрыва, а разрывов в таблице не больше,
    чем строк в её теле, — дальше дробить нечего. Лимит поэтому щедрый: упереться
    в него можно только из-за ошибки. Считать его от числа таблиц нельзя — на
    документе с сотней таблиц фиксированный лимит срабатывал бы на исправной
    вёрстке.
    """
    return sum(len(table.grid.body) for _, table in pending) + 1


def _place_next_split(
        pending: Pending,
        spans: list[range],
        measurer: PageMeasurer,
        probe: Path,
) -> bool:
    """Ставит один разрыв первой таблице, которой он нужен. True — если поставил.

    Разобранные таблицы выбрасываются из pending и больше не измеряются. Это
    корректно: таблицы идут по документу сверху вниз, а разрыв всегда ставится в
    самую верхнюю из нуждающихся, поэтому всё, что могло бы сдвинуть уже
    уместившуюся таблицу, находится ниже неё. Без этого каждый проход заново
    пересчитывал бы вёрстку ради готовых таблиц — а пересчёт стоит O(размера
    документа), и на сотне таблиц подбор становится квадратичным.
    """
    while pending:
        i, table = pending[0]
        head_count, offset = _part_geometry(table)
        with logged_duration(logger, "Таблица %s: измерена часть %d",
                             table.index, len(table.split_after) + 1):
            row = measurer.first_row_on_later_page(
                probe,
                spans[i][-1],  # последняя, ещё не разбитая часть
                head_count,
            )
        if row is None:
            logger.debug("Таблица %s: подбор закончен, частей %d, точки разрыва %s",
                         table.index, len(table.split_after) + 1, table.split_after)
            pending.popleft()  # уместилась целиком
            continue

        split = offset + (row - head_count)
        if split <= offset:
            # Не помещается даже первая строка части: дробить дальше некуда,
            # иначе зациклимся на пустой части.
            logger.warning(
                "Таблица %s: строка %d не помещается на страницу целиком, "
                "разрыв невозможен — Word разорвёт часть сам, без подписи "
                "«Продолжение таблицы»",
                table.index, offset + 1,
            )
            pending.popleft()
            continue

        table.split_after.append(split)
        logger.debug("Таблица %s: разрыв после строки %d", table.index, split)
        return True
    return False


def _part_geometry(table: Table) -> tuple[int, int]:
    """Сколько строк занимает шапка последней части и сколько строк тела до неё."""
    offset = table.split_after[-1] if table.split_after else 0
    return len(table.grid.head_rows), offset
