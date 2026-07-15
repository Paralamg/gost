from pathlib import Path

from docx.document import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Cm, Pt

from gost.elements.element import NumberedElement


class Image(NumberedElement):
    def __init__(
            self,
            index: str,
            path: str,
            alt: str,
    ) -> None:
        """
        Создает изображение с подписью «Рисунок number — alt».
        """
        super().__init__(index)
        self.alt = alt
        self.path = Path(path)

    def render(self, document: Document) -> None:
        if self.path.exists():
            document.add_picture(str(self.path), width=Cm(16.5))
            picture = document.paragraphs[-1]
            picture.alignment = WD_ALIGN_PARAGRAPH.CENTER
            picture.paragraph_format.first_line_indent = Cm(0)
        else:
            picture = document.add_paragraph(f"[Изображение не найдено: {self.path}]")

        # Подпись рисунка идёт под ним, поэтому «не отрывать от следующего»
        # ставится на сам рисунок — иначе он останется внизу страницы один.
        picture.paragraph_format.keep_with_next = True

        caption = document.add_paragraph()
        caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
        caption.paragraph_format.first_line_indent = Cm(0)
        caption.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        caption.paragraph_format.space_after = Pt(6)
        run = caption.add_run(f"Рисунок {self.index} – {self.alt}")
        run.font.size = Pt(14)
