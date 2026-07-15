"""Подбор точек разрыва таблиц по измеренной вёрстке.

Точку разрыва нельзя вычислить — высота строки зависит от переноса слов,
а он от метрик шрифта, ширины ячейки и позиции таблицы на странице. Поэтому
здесь вёрстка измеряется: документ рендерится, Word сообщает, на какой
странице оказалась каждая строка, и разрыв ставится ровно туда.

Разрывы ставятся по одному. Разрыв сдвигает вниз всё, что за ним следует,
поэтому за один проход можно доверять только первому измеренному разрыву —
остальные измерены по вёрстке, которой уже не будет.
"""

from collections.abc import Callable, Sequence
from pathlib import Path

from docx.document import Document

from gost.elements.element import ElementBase
from gost.elements.table import Table
from gost.layout.measurer import PageMeasurer

Build = Callable[[], tuple[Document, list[range]]]

MAX_PASSES = 200


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
    auto = [i for i, element in enumerate(elements)
            if isinstance(element, Table) and element.auto_split]

    for _ in range(MAX_PASSES):
        document, spans = build()
        if not auto:
            return document

        document.save(str(probe))
        if not _place_next_split(elements, auto, spans, measurer, probe):
            return document

    raise RuntimeError(
        f"Не удалось подобрать точки разрыва за {MAX_PASSES} проходов. "
        "Задайте split_after явно."
    )


def _place_next_split(
        elements: Sequence[ElementBase],
        auto: list[int],
        spans: list[range],
        measurer: PageMeasurer,
        probe: Path,
) -> bool:
    """Ставит один разрыв первой таблице, которой он нужен. True — если поставил."""
    for i in auto:
        table = elements[i]
        head_count, offset = _part_geometry(table)
        row = measurer.first_row_on_later_page(
            probe,
            spans[i][-1],  # последняя, ещё не разбитая часть
            head_count,
        )
        if row is None:
            continue

        split = offset + (row - head_count)
        # split == offset — не помещается даже первая строка части: дробить дальше
        # некуда, иначе зациклимся на пустой части.
        if split > offset:
            table.split_after.append(split)
            return True
    return False


def _part_geometry(table: Table) -> tuple[int, int]:
    """Сколько строк занимает шапка последней части и сколько строк тела до неё."""
    if table.split_after:
        return len(table.grid.repeated_rows), table.split_after[-1]
    return len(table.grid.head_rows), 0
