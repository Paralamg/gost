"""Элемент «заголовок раздела»."""

from docx.document import Document

from gost.elements.element import NumberedElement
from gost.styles import ParagraphStyle, apply_paragraph_style


class Head(NumberedElement):
    """Заголовок документа.

    Args:
        index: Номер раздела («1.1»). Пустая строка — у структурного элемента
            номера нет.
        text: Текст заголовка.
        use_number: Выводить ли номер перед текстом. Номер всё равно занят
            фабрикой, поэтому нумерация следующих разделов не сбивается.
        level: Логический уровень заголовка. 0 — ненумерованный структурный
            элемент («Введение», «Основная часть»), 1..4 — нумерованные
            разделы и подразделы («1», «1.1», «1.1.1», «1.1.1.1»).
        style: Оформление заголовка. Применяется поверх базового стиля
            «Heading N», поэтому перекрывает его.
    """

    def __init__(
            self,
            index: str,
            text: str,
            use_number: bool,
            level: int,
            style: ParagraphStyle,
    ) -> None:
        super().__init__(index)
        self.text = text
        self.level = level
        self.use_number = use_number
        self.style = style

    def render(self, document: Document) -> None:
        text = self.__prepare_render_text()
        # Логический уровень 0 занимает Heading 1, поэтому нумерованные
        # разделы начинаются с Heading 2.
        heading = document.add_heading(text, level=self.level + 1)
        apply_paragraph_style(heading, self.style)

    def __prepare_render_text(self) -> str:
        """Текст заголовка вместе с номером, если он есть и нужен."""
        if self.use_number and self.index:
            return f"{self.index} {self.text}"
        return f"{self.text}"
