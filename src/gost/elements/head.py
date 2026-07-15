from docx.document import Document

from gost.elements.element import NumberedElement


class Head(NumberedElement):
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
        document.add_heading(text, level=self.level)

    def __prepare_render_text(self) -> str:
        if self.use_number:
            text = f"{self.index} {self.text}"
        else:
            text = f"{self.text}"

        if self.level == 1:
            text = text.upper()

        return text