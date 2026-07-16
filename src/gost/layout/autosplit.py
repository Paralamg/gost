"""Подбор точек разрыва таблиц по измеренной вёрстке.

Точку разрыва нельзя вычислить — высота строки зависит от переноса слов,
а он от метрик шрифта, ширины ячейки и позиции таблицы на странице. Поэтому
здесь вёрстка измеряется: документ сохраняется, Word сообщает, на какой
странице оказалась каждая строка, и разрыв ставится ровно туда.

Документ строится один раз, сверху вниз. Дойдя до таблицы с «auto», подбираем
ей разрывы на месте: рисуем, измеряем, и если не уместилась — убираем
нарисованное и рисуем заново с разрывом. Всё, что выше таблицы, к этому моменту
уже окончательно, поэтому измерение достоверно, а переделывать приходится
только саму таблицу — не документ целиком.
"""

import logging
from collections.abc import Sequence
from pathlib import Path

from docx.document import Document
from docx.oxml.ns import qn
from docx.oxml.xmlchemy import BaseOxmlElement

from gost.elements.element import ElementBase
from gost.elements.table import Table
from gost.layout.measurer import PageMeasurer
from gost.timing import logged_duration

logger = logging.getLogger(__name__)


def render_with_auto_splits(
        elements: Sequence[ElementBase],
        document: Document,
        measurer: PageMeasurer,
        probe: Path,
) -> None:
    """Выводит элементы в документ, подбирая разрывы таблицам со split_after="auto".

    Args:
        elements: Элементы документа в порядке вывода.
        document: Куда выводить. Стили уже должны быть применены.
        measurer: Чем измерять вёрстку.
        probe: Куда сохранять промежуточный документ для измерения.
    """
    auto = sum(1 for e in elements if isinstance(e, Table) and e.auto_split)
    logger.debug("Автоподбор разрывов: таблиц с «auto» %d из %d элементов",
                 auto, len(elements))

    for element in elements:
        if isinstance(element, Table) and element.auto_split:
            _fit_table(element, document, measurer, probe)
        else:
            element.render(document)


def _fit_table(
        table: Table,
        document: Document,
        measurer: PageMeasurer,
        probe: Path,
) -> None:
    """Подбирает таблице точки разрыва и оставляет её в документе.

    Каждая попытка добавляет ровно один разрыв, а разрывов в таблице не больше,
    чем строк в её теле, — дальше дробить нечего. Предел попыток поэтому просто
    страховка от зацикливания, а не бюджет.
    """
    for _ in range(len(table.grid.body) + 1):
        blocks = _render(table, document)
        split = _overflow_split(table, document, measurer, probe)
        if split is None:
            logger.debug("Таблица %s: подбор закончен, частей %d, точки разрыва %s",
                         table.index, len(table.split_after) + 1, table.split_after)
            return

        # Разрыв меняет вёрстку самой таблицы, поэтому её надо перерисовать —
        # но только её: всё, что выше, разрыв не задевает.
        _remove(document, blocks)
        table.split_after.append(split)
        logger.debug("Таблица %s: разрыв после строки %d", table.index, split)

    raise RuntimeError(
        f"Таблица {table.index}: не удалось подобрать точки разрыва. "
        "Задайте split_after явно."
    )


def _overflow_split(
        table: Table,
        document: Document,
        measurer: PageMeasurer,
        probe: Path,
) -> int | None:
    """Строка тела, после которой последней части нужен разрыв.

    Returns:
        Номер строки (1-based) или None, если разрыв не нужен либо невозможен.
    """
    part = len(table.split_after) + 1
    head_count = len(table.grid.head_rows)
    offset = table.split_after[-1] if table.split_after else 0

    with logged_duration(logger, "Таблица %s: пробник с частью %d сохранён",
                         table.index, part):
        document.save(str(probe))

    with logged_duration(logger, "Таблица %s: часть %d измерена", table.index, part):
        row = measurer.first_row_on_later_page(probe, _last_table_index(document),
                                               head_count)
    if row is None:
        return None  # уместилась

    split = offset + (row - head_count)
    if split <= offset:
        # Не помещается даже первая строка части: дробить дальше некуда, иначе
        # зациклимся на пустой части.
        logger.warning(
            "Таблица %s: строка %d не помещается на страницу целиком, разрыв "
            "невозможен — Word разорвёт часть сам, без подписи «Продолжение таблицы»",
            table.index, offset + 1,
        )
        return None
    return split


def _last_table_index(document: Document) -> int:
    """Индекс последней таблицы документа — это последняя часть той, что мы мерим."""
    return len(document.element.body.findall(qn("w:tbl"))) - 1


def _render(element: ElementBase, document: Document) -> list[BaseOxmlElement]:
    """Выводит элемент и возвращает блоки, которые он добавил в документ."""
    body = document.element.body
    # Список обязателен: lxml создаёт обёртки узлов на лету и уничтожает сразу,
    # как только на них нет ссылки, а id() уничтоженной переиспользуется под
    # другой узел. Без живых ссылок сравнение по id() перепутает узлы и снесёт
    # чужие — например w:sectPr.
    before = list(body)
    known = {id(block) for block in before}
    element.render(document)
    return [block for block in body if id(block) not in known]


def _remove(document: Document, blocks: list[BaseOxmlElement]) -> None:
    """Убирает блоки неудачной попытки. Таблица добавляет только их — следов не остаётся."""
    body = document.element.body
    for block in blocks:
        body.remove(block)
