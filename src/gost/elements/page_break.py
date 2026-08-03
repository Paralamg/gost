"""Элемент «разрыв»: страницы, колонки или раздела."""

from enum import Enum, auto

from docx.document import Document
from docx.enum.section import WD_SECTION_START
from docx.enum.text import WD_BREAK

from gost.elements.element import ElementBase


class BreakType(Enum):
    """Вид разрыва — как в меню Word «Макет → Разрывы»."""

    PAGE = auto()
    """«Страница»: следующий текст начинается с новой страницы."""

    COLUMN = auto()
    """«Колонка»: следующий текст переносится в начало следующей колонки."""

    TEXT_WRAPPING = auto()
    """«Обтекание текстом»: конец строки при обтекании объекта."""

    NEXT_PAGE = auto()
    """«Следующая страница»: разрыв раздела, новый раздел с новой страницы."""

    CONTINUOUS = auto()
    """«Без разрыва»: разрыв раздела на той же странице."""

    EVEN_PAGE = auto()
    """«Чётная страница»: разрыв раздела, новый раздел с чётной страницы."""

    ODD_PAGE = auto()
    """«Нечётная страница»: разрыв раздела, новый раздел с нечётной страницы."""


# Разрывы страниц — вставляются как разрыв внутри абзаца.
_RUN_BREAKS = {
    BreakType.PAGE: WD_BREAK.PAGE,
    BreakType.COLUMN: WD_BREAK.COLUMN,
    BreakType.TEXT_WRAPPING: WD_BREAK.TEXT_WRAPPING,
}

# Разрывы разделов — новый раздел со своими полями и колонтитулами.
_SECTION_STARTS = {
    BreakType.NEXT_PAGE: WD_SECTION_START.NEW_PAGE,
    BreakType.CONTINUOUS: WD_SECTION_START.CONTINUOUS,
    BreakType.EVEN_PAGE: WD_SECTION_START.EVEN_PAGE,
    BreakType.ODD_PAGE: WD_SECTION_START.ODD_PAGE,
}


class PageBreak(ElementBase):
    """Разрыв между элементами документа.

    Args:
        break_type: Вид разрыва. По умолчанию «Следующая страница» —
            разрыв раздела, начинающий новый раздел с новой страницы.
    """

    def __init__(self, break_type: BreakType = BreakType.NEXT_PAGE) -> None:
        self.break_type = break_type

    def render(self, document: Document) -> None:
        if self.break_type in _SECTION_STARTS:
            # Новый раздел наследует поля предыдущего: python-docx клонирует
            # его w:sectPr, поэтому вёрстка по ГОСТ сохраняется.
            document.add_section(_SECTION_STARTS[self.break_type])
        else:
            document.add_paragraph().add_run().add_break(_RUN_BREAKS[self.break_type])
