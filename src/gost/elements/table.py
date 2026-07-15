from collections.abc import Mapping, Sequence
from typing import Any

from docx.document import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

from gost.elements.element import NumberedElement
from gost.elements.table_grid import build_grid
from gost.elements.table_writer import write_caption, write_part

CAPTION = "Таблица {index} – {title}"


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
    ) -> None:
        """Создает таблицу по ГОСТ 7.32 с подписью «Таблица N – title».

        Номер таблицы выделяется всегда, даже при title=None: подпись в этом
        случае не выводится, но номер уже занят.

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

        Raises:
            ValueError: Если столбцов нет или они разной длины.
        """
        super().__init__(index)
        self.grid = build_grid(
            data,
            show_header=show_header,
            show_row_numbers=show_row_numbers,
            show_column_numbers=show_column_numbers,
        )
        self.title = title

    def render(self, document: Document) -> None:
        if self.title is not None:
            write_caption(
                document,
                CAPTION.format(index=self.index, title=self.title),
                align=WD_ALIGN_PARAGRAPH.LEFT,
            )

        write_part(document, self.grid, self.grid.head_rows + self.grid.body)
        document.add_paragraph()
