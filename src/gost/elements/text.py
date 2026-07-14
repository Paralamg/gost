from docx.document import Document

from gost.elements.element import ElementBase
from gost.utils import render_bold


class Text(ElementBase):
    def __init__(self, text: str) -> None:
        self.text = text

    def render(self, document: Document) -> None:
        paragraph = document.add_paragraph()
        render_bold(paragraph, self.text)