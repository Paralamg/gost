import logging

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Cm, Pt, RGBColor

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
    "StyleSheet",
    "ParagraphStyle",
    "Pt",
    "Cm",
    "RGBColor",
    "WD_ALIGN_PARAGRAPH",
    "WD_LINE_SPACING",
]
