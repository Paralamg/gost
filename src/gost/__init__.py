"""Генерация документов Word (.docx), оформленных по ГОСТ 7.32-2017.

Документ собирается из элементов: заголовков, абзацев, таблиц, рисунков, формул
и разрывов. Элементы создаёт :class:`ElementFactory` — она же ведёт счётчики
номеров и раздаёт элементам оформление, — а :class:`WordBuilder` собирает из них
файл::

    from pathlib import Path
    from gost import ElementFactory, WordBuilder

    wb = WordBuilder()
    factory = ElementFactory()
    wb.add_element(factory.create_head("Введение", False, 0))
    wb.add_element(factory.create_text("Обычный абзац с **жирным** фрагментом."))
    wb.save(Path("report.docx"))

Нумерация и оформление — забота библиотеки: номера рисунков, таблиц и формул
проставляются сами, а стили по умолчанию берутся из :meth:`StyleSheet.gost`.
"""

import logging

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Cm, Pt, RGBColor

from .element_factory import ElementFactory
from .elements.page_break import BreakType
from .index.index_manager import IndexManager
from .index.index_type import IndexType
from .styles import ParagraphStyle, StyleSheet
from .word_builder import WordBuilder

# Библиотека не навязывает приложению настройку логирования: без обработчика
# записи просто отбрасываются. Всё пишется в логгер «gost».
logging.getLogger(__name__).addHandler(logging.NullHandler())

# Стили настраиваются объектами ParagraphStyle/StyleSheet; docx-типы для их полей
# (размеры, цвет, выравнивание, интервал) реэкспортируются, чтобы не импортировать
# из docx напрямую.
__all__ = [
    "WordBuilder",
    "ElementFactory",
    "IndexManager",
    "IndexType",
    "BreakType",
    "StyleSheet",
    "ParagraphStyle",
    "Pt",
    "Cm",
    "RGBColor",
    "WD_ALIGN_PARAGRAPH",
    "WD_LINE_SPACING",
]
