from typing import Optional

import pandas as pd
from docx.document import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Cm, Pt

from gost.elements.element import NumberedElement
from gost.utils import format_value


class Table(NumberedElement):
    def __init__(
            self,
            number: str,
            df: pd.DataFrame,
            float_format: Optional[str] = ".1f",
            title: Optional[str] = None,
            show_header: bool = True,
            show_index: bool = True,
            bold_header: bool = False,
            bold_index: bool = False,
    ) -> None:
        """Создает таблицу из DataFrame с подписью «Таблица number — title».

        Подпись добавляется только при наличии title, при этом счётчик таблиц увеличивается.

        Args:
            df: Данные для отображения в таблице.
            float_format: Формат вещественных чисел (например, «.1f»). None — без форматирования.
            title: Заголовок таблицы. Если None, подпись и счётчик не добавляются.
            show_header: Добавлять ли строку с именами столбцов.
            show_index: Добавлять ли столбец с индексом DataFrame.
            bold_header: Выделять ли заголовки столбцов жирным.
            bold_index: Выделять ли значения индекса жирным.

        Returns:
            Текущий экземпляр WordBuilder для цепочки вызовов.
        """
        super().__init__(number)
        self.df = df
        self.float_format = float_format
        self.title = title
        self.show_header = show_header
        self.show_index = show_index
        self.bold_header = bold_header
        self.bold_index = bold_index

    def render(self, document: Document) -> None:
        if self.title is not None:
            caption = document.add_paragraph()
            caption.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            caption.paragraph_format.first_line_indent = Cm(0)
            caption.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            caption.paragraph_format.space_after = Pt(6)
            run = caption.add_run(f"Таблица – {self.number}  {self.title}")
            run.font.size = Pt(14)

        header_rows = 1 if self.show_header else 0
        rows = len(self.df) + header_rows
        index_offset = 1 if self.show_index else 0
        cols = len(self.df.columns) + index_offset

        table = document.add_table(rows=rows, cols=cols)
        table.style = "Table Grid"
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        if self.show_header:
            if self.show_index:
                table.cell(0, 0).text = ""
            for j, col_name in enumerate(self.df.columns):
                cell = table.cell(0, j + index_offset)
                cell.text = ""
                run = cell.paragraphs[0].add_run(format_value(col_name, self.float_format))
                run.bold = self.bold_header

        for i, (idx, row) in enumerate(self.df.iterrows()):
            row_idx = i + header_rows
            if self.show_index:
                idx_cell = table.cell(row_idx, 0)
                idx_cell.text = ""
                run = idx_cell.paragraphs[0].add_run(str(idx))
                run.bold = self.bold_index
            for j, val in enumerate(row):
                table.cell(row_idx, j + index_offset).text = format_value(val, self.float_format)

        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    para.style = "Table Text"

        document.add_paragraph()
