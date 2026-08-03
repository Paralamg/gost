"""Элементы документа: из них WordBuilder собирает файл.

Создавать их напрямую обычно не нужно — номера и оформление раздаёт
:class:`~gost.element_factory.ElementFactory`.
"""

from gost.elements.table.table import Table
from gost.elements.text import Text
from gost.elements.image import Image
from gost.elements.head import Head
from gost.elements.formula import Formula
from gost.elements.page_break import BreakType, PageBreak
from gost.elements.element import ElementBase