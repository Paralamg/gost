from docx.document import Document

from gost.elements.element import NumberedElement


class Formula(NumberedElement):
    def __init__(self, latex: str):
        pass

    def render(self, document: Document) -> None:
        pass