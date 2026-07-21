"""Вывод сетки таблицы в документ Word."""

from docx.document import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Emu, Length, Pt
from docx.oxml.xmlchemy import BaseOxmlElement
from docx.table import Table as DocxTable, _Cell, _Row

from gost.elements.table_grid import Grid

# Дети w:tcPr, которые по схеме идут после w:tcBorders.
_AFTER_TC_BORDERS = (
    "w:shd", "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign",
    "w:hideMark", "w:headers", "w:cellIns", "w:cellDel", "w:cellMerge", "w:tcPrChange",
)

TABLE_STYLE = "Table Grid"
CELL_STYLE = "Table Text"
CAPTION_SIZE = Pt(14)
ROW_NUMBER_WIDTH = Cm(1.2)


def write_caption(
        document: Document,
        text: str,
        *,
        page_break_before: bool = False,
) -> None:
    """Пишет подпись таблицы отдельным абзацем.

    Подпись первой части и подписи продолжений оформляются одинаково: ГОСТ 7.32-2017
    (6.6.3) требует писать слева и «Таблица N – ...», и «Продолжение таблицы N».
    """
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    fmt = paragraph.paragraph_format
    fmt.first_line_indent = Cm(0)
    fmt.line_spacing_rule = WD_LINE_SPACING.SINGLE
    fmt.space_after = Pt(6)
    fmt.keep_with_next = True
    if page_break_before:
        fmt.page_break_before = True

    paragraph.add_run(text).font.size = CAPTION_SIZE


def write_part(
        document: Document,
        grid: Grid,
        rows: list[list[str]],
        *,
        repeat_head: bool = False,
) -> None:
    """Пишет непрерывный блок строк как отдельную таблицу.

    Args:
        grid: Сетка целиком — нужна для ширины и раскладки столбцов.
        rows: Строки этой части, включая её шапку.
        repeat_head: Помечать ли шапку как повторяемую на каждой странице
            («w:tblHeader»). Только для неразрывной таблицы: при ручном
            разбиении шапка пишется в каждую часть явно.
    """
    table = document.add_table(rows=len(rows), cols=grid.width)
    table.style = TABLE_STYLE
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    widths = _column_widths(document, grid)
    _fix_layout(table, widths)

    # Идентификатор стиля резолвится один раз на таблицу: paragraph.style = "Table Text"
    # ищет стиль по имени и перебирает все стили документа на каждой ячейке — на
    # таблице в сотню строк это 85% времени сборки.
    style_id = document.styles[CELL_STYLE].style_id

    for row, values in zip(table.rows, rows):
        _forbid_split(row)
        for cell, value, width in zip(row.cells, values, widths):
            cell.width = width  # Word ориентируется на ширину ячейки (tcW), а не на gridCol
            # По горизонтали текст центрует стиль «Table Text»; вертикальное
            # выравнивание — свойство ячейки, в стиль абзаца его не убрать.
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            paragraph = cell.paragraphs[0]
            paragraph._p.get_or_add_pPr().style = style_id
            paragraph.text = value

    for row in table.rows[:len(grid.head_rows)]:
        _keep_with_next(row)
        if repeat_head:
            _repeat_as_header(row)

    if grid.head_rows:
        _double_bottom_border(table.rows[len(grid.head_rows) - 1])


def _column_widths(document: Document, grid: Grid) -> list[Length]:
    section = document.sections[0]
    total = section.page_width - section.left_margin - section.right_margin

    if not grid.has_row_numbers:
        return [Emu(total // grid.width)] * grid.width

    data_columns = grid.width - 1
    share = Emu((total - ROW_NUMBER_WIDTH) // data_columns)
    return [ROW_NUMBER_WIDTH, *[share] * data_columns]


def _fix_layout(table: DocxTable, widths: list[Length]) -> None:
    """Фиксирует ширины столбцов.

    Без этого Word считает autofit для каждой таблицы отдельно, и части
    разорванной таблицы разъезжаются по ширине. Ширины ячеек проставляет
    write_part — Word смотрит на них, а не на gridCol.
    """
    table.autofit = False
    for column, width in zip(table.columns, widths):
        column.width = width


def _forbid_split(row: _Row) -> None:
    """«w:cantSplit» — строка не разрывается между страницами."""
    row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))


def _keep_with_next(row: _Row) -> None:
    """«w:keepNext» — строка не отрывается от следующей.

    Держит строки шапки вместе и не даёт им остаться внизу страницы без единой
    строки тела.
    """
    for cell in row.cells:
        for paragraph in cell.paragraphs:
            paragraph.paragraph_format.keep_with_next = True


def _double_bottom_border(row: _Row) -> None:
    """«w:bottom w:val=double» — двойная линия, отделяющая шапку от тела."""
    for cell in row.cells:
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "double")
        bottom.set(qn("w:sz"), "4")  # толщина каждой из двух линий, 1/8 пункта
        bottom.set(qn("w:space"), "0")
        bottom.set(qn("w:color"), "auto")
        _cell_borders(cell).append(bottom)


def _cell_borders(cell: _Cell) -> BaseOxmlElement:
    """«w:tcBorders» ячейки, создавая его при необходимости.

    python-docx не регистрирует w:tcBorders, а порядок детей w:tcPr задан схемой,
    поэтому элемент ставится перед тем, что по схеме идёт после него, — append
    положил бы его после уже проставленного w:vAlign.
    """
    tcPr = cell._tc.get_or_add_tcPr()
    borders = tcPr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tcPr.insert_element_before(borders, *_AFTER_TC_BORDERS)
    return borders


def _repeat_as_header(row: _Row) -> None:
    """«w:tblHeader» — строка повторяется на каждой странице.

    Word повторяет только непрерывный блок строк от начала таблицы.
    """
    row._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))
