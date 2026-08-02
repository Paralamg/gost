"""Подбор точек разрыва таблиц по измеренной вёрстке.

Точку разрыва нельзя вычислить — высота строки зависит от переноса слов, а он
от метрик шрифта, ширины ячейки и позиции таблицы на странице. Поэтому вёрстка
измеряется: Word сообщает, на какой странице оказалась каждая строка, и разрыв
ставится ровно туда.

Документ строится один раз, сверху вниз, и параллельно повторяется в зеркале —
том же документе, открытом в Word. Дойдя до таблицы с «auto», досылаем в
зеркало всё, что накопилось перед ней, и подбираем разрывы: дописываем таблицу,
меряем, а если не уместилась — отменяем вставку и дописываем заново с разрывом.

Всё, что выше таблицы, к этому моменту уже окончательно, поэтому измерение
достоверно. А раз зеркало только достраивается с конца, Word пересчитывает
вёрстку от места вставки, а не целиком.
"""

import logging
from collections.abc import Callable, Sequence
from pathlib import Path

from docx.document import Document

from gost.elements import ElementBase
from gost.elements import Table
from gost.layout.measurer import LayoutMirror
from gost.progress import tracked
from gost.timing import logged_duration

logger = logging.getLogger(__name__)

NewDocument = Callable[[], Document]


def render_with_auto_splits(
        elements: Sequence[ElementBase],
        document: Document,
        mirror: LayoutMirror,
        new_document: NewDocument,
        workdir: Path,
) -> None:
    """Выводит элементы в документ, подбирая разрывы таблицам со split_after="auto".

    Args:
        elements: Элементы документа в порядке вывода.
        document: Куда выводить. Стили уже должны быть применены.
        mirror: Зеркало документа, по которому меряется вёрстка.
        new_document: Создаёт пустой документ с теми же стилями — для кусков.
        workdir: Куда складывать куски для зеркала.
    """
    auto = sum(1 for e in elements if isinstance(e, Table) and e.auto_split)
    logger.debug("Автоподбор разрывов: таблиц с «auto» %d из %d элементов",
                 auto, len(elements))

    chunks = _Chunks(new_document, workdir)
    pending: list[ElementBase] = []

    for element in tracked(elements, "Сборка документа"):
        if isinstance(element, Table) and element.auto_split:
            # Зеркало должно догнать документ до этой таблицы, иначе она встанет
            # не на своё место и померится не там.
            if pending:
                mirror.append(chunks.of(pending))
                pending.clear()
            _fit_table(element, mirror, chunks)
            element.render(document)  # уже с подобранными разрывами
        else:
            element.render(document)
            pending.append(element)

    # Хвост после последней auto-таблицы в зеркало досылать незачем: мерить нечего.


def _fit_table(table: Table, mirror: LayoutMirror, chunks: "_Chunks") -> None:
    """Подбирает таблице точки разрыва и оставляет её в зеркале.

    Каждая попытка добавляет ровно один разрыв, а разрывов в таблице не больше,
    чем строк в её теле, — дальше дробить нечего. Предел попыток поэтому просто
    страховка от зацикливания, а не бюджет.
    """
    for _ in range(len(table.grid.body) + 1):
        mirror.append(chunks.of([table]))
        split = _overflow_split(table, mirror)
        if split is None:
            logger.debug("Таблица %s: подбор закончен, частей %d, точки разрыва %s",
                         table.index, len(table.split_after) + 1, table.split_after)
            return  # таблица остаётся в зеркале — она префикс для следующих

        mirror.undo()
        table.split_after.append(split)
        logger.debug("Таблица %s: разрыв после строки %d", table.index, split)

    raise RuntimeError(
        f"Таблица {table.index}: не удалось подобрать точки разрыва. "
        "Задайте split_after явно."
    )


def _overflow_split(table: Table, mirror: LayoutMirror) -> int | None:
    """Строка тела, после которой последней части нужен разрыв.

    Returns:
        Номер строки (1-based) или None, если разрыв не нужен либо невозможен.
    """
    part = len(table.split_after) + 1
    head_count = len(table.grid.head_rows)
    offset = table.split_after[-1] if table.split_after else 0

    with logged_duration(logger, "Таблица %s: часть %d измерена", table.index, part):
        row = mirror.first_row_on_later_page(head_count)
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


class _Chunks:
    """Куски документа для зеркала: элементы, отрисованные в отдельный файл.

    Элементы приходится рисовать дважды — в документ и в кусок, — потому что
    перенести готовую разметку между документами python-docx не позволяет:
    картинки живут отдельными частями пакета и ссылки на них не переедут.
    Вторая отрисовка стоит O(элемента), измерение по-старому стоило бы
    O(документа) — размен выгодный.
    """

    def __init__(self, new_document: NewDocument, workdir: Path) -> None:
        self.__new_document = new_document
        self.__workdir = workdir
        self.__seq = 0

    def of(self, elements: Sequence[ElementBase]) -> Path:
        """Рисует элементы в отдельный файл и возвращает путь к нему."""
        self.__seq += 1
        path = self.__workdir / f"chunk_{self.__seq}.docx"

        document = self.__new_document()
        for element in elements:
            element.render(document)
        document.save(str(path))
        return path
