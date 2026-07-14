from pathlib import Path

from docx.document import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Cm, Pt

from gost.elements.element import NumberedElement


class Image(NumberedElement):
    def __init__(
            self,
            number: str,
            path: str,
            alt: str,
    ) -> None:
        """
        Создает изображение с подписью «Рисунок number — alt».
        """
        super().__init__(number)
        self.alt = alt
        self.path = Path(path)

    def render(self, document: Document) -> None:
        if self.path.exists():
            document.add_picture(str(self.path), width=Cm(16.5))
            last_paragraph = document.paragraphs[-1]
            last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            last_paragraph.paragraph_format.first_line_indent = Cm(0)
        else:
            document.add_paragraph(f"[Изображение не найдено: {self.path}]")

        caption = document.add_paragraph()
        caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
        caption.paragraph_format.first_line_indent = Cm(0)
        caption.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        caption.paragraph_format.space_after = Pt(6)
        run = caption.add_run(f"Рисунок {self.number} – {self.alt}")
        run.font.size = Pt(14)
