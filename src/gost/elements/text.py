from docx.document import Document

from gost.elements.element import ElementBase
from gost.inline import render_inline


class Text(ElementBase):
    def __init__(self, text: str) -> None:
        self.text = text

    def render(self, document: Document) -> None:
        paragraph = document.add_paragraph()
        render_inline(paragraph, self.text)