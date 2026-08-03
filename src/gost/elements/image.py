"""Элемент «рисунок с подписью»."""

import logging
from pathlib import Path

from docx.document import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm

from gost.elements.element import NumberedElement
from gost.inline import render_inline
from gost.styles import ParagraphStyle, apply_paragraph_style
from gost.timing import logged_duration

logger = logging.getLogger(__name__)


class Image(NumberedElement):
    """Рисунок с подписью «Рисунок N – alt».

    Рисунок растягивается по ширине текста (16,5 см) с сохранением пропорций и
    ставится по центру, подпись — под ним. Файл читается при выводе документа,
    а не при создании элемента: если к тому моменту его не окажется на месте,
    вместо рисунка встанет заглушка, а в лог уйдёт предупреждение.

    Args:
        index: Номер рисунка.
        path: Путь к файлу изображения.
        alt: Подпись под рисунком. Понимает инлайн-разметку.
        caption_style: Оформление подписи.
    """

    def __init__(
            self,
            index: str,
            path: str,
            alt: str,
            caption_style: ParagraphStyle,
    ) -> None:
        super().__init__(index)
        self.alt = alt
        self.path = Path(path)
        self.caption_style = caption_style

    def render(self, document: Document) -> None:
        if self.path.exists():
            # Вставка читает и перекодирует файл — на тяжёлых картинках заметно.
            with logged_duration(logger, "Рисунок %s вставлен: %s", self.index, self.path.name):
                document.add_picture(str(self.path), width=Cm(16.5))
            picture = document.paragraphs[-1]
            picture.alignment = WD_ALIGN_PARAGRAPH.CENTER
            picture.paragraph_format.first_line_indent = Cm(0)
        else:
            logger.warning("Рисунок %s: файл не найден, вставлена заглушка: %s",
                           self.index, self.path)
            picture = document.add_paragraph(f"[Изображение не найдено: {self.path}]")

        # Подпись рисунка идёт под ним, поэтому «не отрывать от следующего»
        # ставится на сам рисунок — иначе он останется внизу страницы один.
        picture.paragraph_format.keep_with_next = True

        caption = document.add_paragraph()
        render_inline(caption, f"Рисунок {self.index} – {self.alt}")
        apply_paragraph_style(caption, self.caption_style)
