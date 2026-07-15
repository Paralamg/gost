from collections.abc import Mapping, Sequence
from typing import Any, Literal

from docx.document import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

from gost.elements.element import NumberedElement
from gost.elements.table_grid import build_grid
from gost.elements.table_writer import write_caption, write_part

CAPTION = "Таблица {index} – {title}"
CONTINUATION = "Продолжение табл. {index}"

AUTO = "auto"
SplitAfter = Sequence[int] | Literal["auto"] | None


class Table(NumberedElement):
    def __init__(
            self,
            index: str,
            data: Mapping[str, Sequence[Any]],
            title: str | None = None,
            *,
            show_header: bool = True,
            show_row_numbers: bool = True,
            show_column_numbers: bool = True,
            split_after: SplitAfter = None,
    ) -> None:
        """Создает таблицу по ГОСТ 7.32 с подписью «Таблица N – title».

        Номер таблицы выделяется всегда, даже при title=None: подпись в этом
        случае не выводится, но номер уже занят и появится в «Продолжение
        табл. N» при разрыве.

        Args:
            data: Данные по столбцам: {«Показатель А»: [1.0, 2.5], ...}.
                Совместимо с df.to_dict("list"). Индексы строк не учитываются.
                Значения приводятся к строке через str() — числа форматируйте
                до передачи, иначе 0.1 + 0.2 попадет в документ как
                «0.30000000000000004».
            title: Заголовок таблицы. None — подпись не выводится.
            show_header: Выводить строку с именами столбцов.
            show_row_numbers: Добавлять слева столбец «№ п/п» с нумерацией строк.
            show_column_numbers: Добавлять строку с номерами столбцов 1..N.
            split_after: Номера строк тела (1-based), после которых таблица
                разрывается: [20] или range(20, row_count, 20). Каждая часть
                получает подпись «Продолжение табл. N» и строку номеров
                столбцов.
                «auto» — подобрать точки разрыва по месту на странице; требует
                Word и extra «autosplit», подбор делает WordBuilder.save().
                None — таблицу разбивает сам Word, повторяя шапку на каждой
                странице, но без подписи «Продолжение».

        Raises:
            ValueError: Если столбцов нет, они разной длины или split_after
                выходит за пределы тела таблицы.
        """
        super().__init__(index)
        self.grid = build_grid(
            data,
            show_header=show_header,
            show_row_numbers=show_row_numbers,
            show_column_numbers=show_column_numbers,
        )
        self.title = title
        self.auto_split = split_after == AUTO
        self.split_after = (
            [] if self.auto_split
            else _validate_splits(split_after, len(self.grid.body))
        )

    def render(self, document: Document) -> None:
        first, *rest = _split(self.grid.body, self.split_after)

        if self.title is not None:
            write_caption(
                document,
                CAPTION.format(index=self.index, title=self.title),
                align=WD_ALIGN_PARAGRAPH.LEFT,
            )

        # Повтор шапки силами Word — только когда мы не разбиваем таблицу сами,
        # иначе строка номеров задвоится на странице продолжения.
        write_part(document, self.grid, self.grid.head_rows + first, repeat_head=not rest)

        for part in rest:
            write_caption(
                document,
                CONTINUATION.format(index=self.index),
                align=WD_ALIGN_PARAGRAPH.RIGHT,
                page_break_before=True,
            )
            write_part(document, self.grid, self.grid.repeated_rows + part)

        document.add_paragraph()


def _validate_splits(split_after: Sequence[int] | None, row_count: int) -> list[int]:
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
    bounds = [0, *splits, len(body)]
    return [body[start:end] for start, end in zip(bounds, bounds[1:])]
