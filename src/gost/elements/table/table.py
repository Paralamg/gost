"""Элемент «таблица»: подпись, шапка и разбиение на части."""

import logging
from collections.abc import Mapping, Sequence
from typing import Any, Literal

from docx.document import Document

from gost.elements.element import NumberedElement
from gost.elements.table.table_grid import build_grid
from gost.elements.table.table_writer import write_caption, write_part
from gost.styles import ParagraphStyle
from gost.timing import logged_duration

logger = logging.getLogger(__name__)

CAPTION = "Таблица {index} – {title}"
CONTINUATION = "Продолжение таблицы {index}"

AUTO = "auto"

# Точки разрыва таблицы: номера строк, «auto» (подобрать измерением) или None
# (не разбивать самим — оставить это Word).
SplitAfter = Sequence[int] | Literal["auto"] | None


class Table(NumberedElement):
    """Таблица по ГОСТ 7.32-2017 с подписью «Таблица N – title».

    Длинная таблица разбивается на части: каждая выводится отдельной таблицей
    Word со своей шапкой и подписью «Продолжение таблицы N». Точки разрыва
    задаются номерами строк либо подбираются измерением вёрстки — см.
    *split_after*.
    """

    def __init__(
            self,
            index: str,
            data: Mapping[str, Sequence[Any]],
            title: str,
            *,
            caption_style: ParagraphStyle,
            text_style: ParagraphStyle,
            show_header: bool = True,
            show_row_numbers: bool = True,
            show_column_numbers: bool = True,
            split_after: SplitAfter = None,
    ) -> None:
        """Раскладывает данные в сетку и проверяет их — до вывода в документ.

        Args:
            index: Номер таблицы.
            data: Данные по столбцам: {«Показатель А»: [1.0, 2.5], ...}.
                Совместимо с df.to_dict("list"). Индексы строк не учитываются.
                Значения приводятся к строке через str() — числа форматируйте
                до передачи, иначе 0.1 + 0.2 попадет в документ как
                «0.30000000000000004».
            title: Заголовок таблицы. Обязателен: ГОСТ 7.32-2017 (6.6.2) требует
                наименование у каждой таблицы.
            caption_style: Оформление подписи «Таблица N – …».
            text_style: Оформление текста внутри ячеек таблицы.
            show_header: Выводить строку с именами столбцов.
            show_row_numbers: Добавлять слева столбец «№ п/п» с нумерацией строк.
            show_column_numbers: Добавлять строку с номерами столбцов 1..N.
            split_after: Номера строк тела (1-based), после которых таблица
                разрывается: [20] или range(20, row_count, 20). Каждая часть
                получает подпись «Продолжение таблицы N» и повтор шапки.
                «auto» — подобрать точки разрыва по месту на странице; требует
                Word и extra «autosplit», подбор делает WordBuilder.save().
                None — таблицу разбивает сам Word, повторяя шапку на каждой
                странице, но без подписи «Продолжение».

        Raises:
            ValueError: Если заголовок пуст, столбцов нет, они разной длины или
                split_after выходит за пределы тела таблицы.
        """
        super().__init__(index)
        if not title or not title.strip():
            raise ValueError("Таблице нужен заголовок")
        self.grid = build_grid(
            data,
            show_header=show_header,
            show_row_numbers=show_row_numbers,
            show_column_numbers=show_column_numbers,
        )
        self.title = title
        self.caption_style = caption_style
        self.text_style = text_style
        self.auto_split = split_after == AUTO
        self.split_after = (
            [] if self.auto_split
            else _validate_splits(split_after, len(self.grid.body))
        )

    def render(self, document: Document) -> None:
        first, *rest = _split(self.grid.body, self.split_after)

        with logged_duration(
                logger, "Таблица %s выведена: строк %d, столбцов %d, частей %d",
                self.index, len(self.grid.body), self.grid.width, len(rest) + 1,
        ):
            # Подпись есть у каждой таблицы, поэтому она же отделяет её от предыдущей:
            # два w:tbl подряд Word слил бы в одну таблицу.
            write_caption(
                document,
                CAPTION.format(index=self.index, title=self.title),
                self.caption_style,
            )

            # Повтор шапки силами Word — только когда мы не разбиваем таблицу сами,
            # иначе шапка задвоится на странице продолжения.
            write_part(
                document, self.grid, self.grid.head_rows + first,
                self.text_style, repeat_head=not rest,
            )

            for part in rest:
                write_caption(
                    document,
                    CONTINUATION.format(index=self.index),
                    self.caption_style,
                    page_break_before=True,
                )
                write_part(document, self.grid, self.grid.head_rows + part, self.text_style)


def _validate_splits(split_after: Sequence[int] | None, row_count: int) -> list[int]:
    """Приводит точки разрыва к возрастающему списку без повторов.

    Raises:
        ValueError: Если номер строки выходит за пределы тела таблицы. Разрыв
            после последней строки бессмысленен — часть за ним пуста.
    """
    if not split_after:
        return []

    splits = sorted(set(split_after))
    if splits[0] < 1 or splits[-1] >= row_count:
        raise ValueError(
            f"split_after: ожидались номера строк 1..{row_count - 1}, "
            f"получено {list(split_after)}"
        )
    return splits


def _split(body: list[list[str]], splits: list[int]) -> list[list[list[str]]]:
    """Режет тело таблицы на части по точкам разрыва. Без них — одна часть."""
    bounds = [0, *splits, len(body)]
    return [body[start:end] for start, end in zip(bounds, bounds[1:])]
