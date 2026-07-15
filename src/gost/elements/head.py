from docx.document import Document

from gost.elements.element import NumberedElement


class Head(NumberedElement):
    """Заголовок документа.

    Args:
        level: Логический уровень заголовка. 0 — ненумерованный структурный
            элемент («Введение», «Основная часть»), 1..4 — нумерованные
            разделы и подразделы («1», «1.1», «1.1.1», «1.1.1.1»).
    """

    def __init__(
            self,
            index: str,
            use_number: bool,
            text: str,
            level: int = 1,
    ) -> None:
        super().__init__(index)
        self.text = text
        self.level = level
        self.use_number = use_number

    def render(self, document: Document) -> None:
        text = self.__prepare_render_text()
        # Логический уровень 0 занимает Heading 1, поэтому нумерованные
        # разделы начинаются с Heading 2.
        document.add_heading(text, level=self.level + 1)

    def __prepare_render_text(self) -> str:
        if self.use_number and self.index:
            return f"{self.index} {self.text}"
        return f"{self.text}"