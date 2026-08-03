"""Элемент «абзац обычного текста»."""

from docx.document import Document

from gost.elements.element import ElementBase
from gost.inline import render_inline
from gost.styles import ParagraphStyle, apply_paragraph_style


class Text(ElementBase):
    """Абзац обычного текста.

    Args:
        text: Текст абзаца. Понимает инлайн-разметку — см.
            :mod:`gost.inline`.
        style: Оформление абзаца.
    """

    def __init__(self, text: str, style: ParagraphStyle) -> None:
        self.text = text
        self.style = style

    def render(self, document: Document) -> None:
        paragraph = document.add_paragraph()
        render_inline(paragraph, self.text)
        apply_paragraph_style(paragraph, self.style)
