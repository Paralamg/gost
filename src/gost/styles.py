import hashlib
from dataclasses import dataclass, field, replace

from docx.document import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Length, Pt, RGBColor
from docx.text.paragraph import Paragraph

FONT = "Times New Roman"
NORMAL_STYLE = "Normal"
CELL_STYLE = "Table Text"


@dataclass
class ParagraphStyle:
    """Оформление одного абзаца: шрифт и параметры абзаца.

    Мутабельный — пользователь меняет поля напрямую (``style.font_size = Pt(12)``).
    ``font_name`` и ``font_size`` задаются всегда; остальные поля со значением
    ``None`` не трогаются — абзац наследует их от именованного стиля Word.
    """

    font_name: str = FONT
    font_size: Length = field(default_factory=lambda: Pt(14))
    bold: bool | None = None
    italic: bool | None = None
    all_caps: bool | None = None
    color: RGBColor | None = None
    alignment: WD_ALIGN_PARAGRAPH | None = None
    line_spacing: WD_LINE_SPACING | None = None
    first_line_indent: Length | None = None
    space_before: Length | None = None
    space_after: Length | None = None

    def copy(self) -> "ParagraphStyle":
        # Поля — неизменяемые value-типы (Length/RGBColor/enum), shallow-копии
        # достаточно, чтобы правки элемента не текли в общий StyleSheet.
        return replace(self)


@dataclass
class StyleSheet:
    """Набор стилей документа: глобальные значения по умолчанию для элементов.

    ``headings`` индексируется логическим уровнем заголовка 0..4 (уровень 0 —
    ненумерованный структурный элемент, «Heading 1»).
    """

    normal: ParagraphStyle
    headings: list[ParagraphStyle]
    table_caption: ParagraphStyle
    table_text: ParagraphStyle
    image_caption: ParagraphStyle

    @classmethod
    def gost(cls) -> "StyleSheet":
        """Значения по ГОСТ 7.32-2017 — дефолт библиотеки."""
        return cls(
            normal=ParagraphStyle(
                font_size=Pt(14),
                alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                line_spacing=WD_LINE_SPACING.ONE_POINT_FIVE,
                first_line_indent=Cm(1.25),
                space_before=Pt(0),
                space_after=Pt(0),
            ),
            headings=[_gost_heading(*cfg) for cfg in _HEADING_CONFIGS],
            table_caption=ParagraphStyle(
                font_size=Pt(14),
                alignment=WD_ALIGN_PARAGRAPH.LEFT,
                line_spacing=WD_LINE_SPACING.SINGLE,
                first_line_indent=Cm(0),
                space_after=Pt(6),
            ),
            table_text=ParagraphStyle(
                font_size=Pt(12),
                alignment=WD_ALIGN_PARAGRAPH.CENTER,
                first_line_indent=Cm(0),
                space_before=Pt(0),
                space_after=Pt(0),
            ),
            image_caption=ParagraphStyle(
                font_size=Pt(14),
                alignment=WD_ALIGN_PARAGRAPH.CENTER,
                line_spacing=WD_LINE_SPACING.SINGLE,
                first_line_indent=Cm(0),
                space_after=Pt(6),
            ),
        )

    def copy(self) -> "StyleSheet":
        return StyleSheet(
            normal=self.normal.copy(),
            headings=[h.copy() for h in self.headings],
            table_caption=self.table_caption.copy(),
            table_text=self.table_text.copy(),
            image_caption=self.image_caption.copy(),
        )


# (bold, centered, all_caps) по логическим уровням заголовков 0..4.
# Heading 1 — ненумерованный структурный элемент («Введение»), по центру,
# прописными; Heading 2..5 — нумерованные разделы и подразделы.
_HEADING_CONFIGS = [
    (True,  True,  True),
    (True,  False, False),
    (True,  False, False),
    (False, False, False),
    (False, False, False),
]


def _gost_heading(bold: bool, centered: bool, all_caps: bool) -> ParagraphStyle:
    return ParagraphStyle(
        font_size=Pt(14),
        bold=bold,
        italic=False,
        all_caps=all_caps,
        color=RGBColor(0, 0, 0),
        alignment=WD_ALIGN_PARAGRAPH.CENTER if centered else WD_ALIGN_PARAGRAPH.JUSTIFY,
        line_spacing=WD_LINE_SPACING.ONE_POINT_FIVE,
        first_line_indent=Cm(0) if centered else Cm(1.25),
        space_before=Pt(0),
        space_after=Pt(0),
    )


def apply_paragraph_style(paragraph: Paragraph, style: ParagraphStyle) -> None:
    """Применяет *style* к абзацу прямым форматированием.

    Ставит параметры абзаца и шрифт на каждый run. Не мутирует *style* — один и
    тот же элемент рендерится в несколько документов при автоподборе разрывов.
    """
    _apply_parfmt(paragraph.paragraph_format, style)
    for run in paragraph.runs:
        _apply_font(run.font, style)


def resolve_table_text_style(document: Document, style: ParagraphStyle) -> str:
    """Возвращает style_id именованного стиля для текста ячеек по *style*.

    Стиль ячеек нельзя ставить прямым форматированием на каждую ячейку — это
    85% времени сборки таблицы. Поэтому под каждую отличающуюся конфигурацию
    регистрируется отдельный именованный стиль, а write_part резолвит его id
    один раз на таблицу. Имя детерминировано по содержимому стиля, поэтому
    одинаковые конфигурации переиспользуют один стиль, а дефолт ГОСТ совпадает
    с базовым «Table Text», созданным apply_gost_styles.
    """
    name = _table_text_style_name(style)
    try:
        s = document.styles[name]
    except KeyError:
        s = _create_paragraph_style(document, name, style)
    return s.style_id


def apply_gost_styles(doc: Document) -> None:
    """Configure GOST 7.32-2001 base styles, margins and language on *doc* in-place.

    Настраивает базовые именованные стили (Normal, Table Text, Heading 1..5) —
    к ним элементы применяют своё оформление поверх. Значения берутся из
    StyleSheet.gost(), чтобы не дублировать дефолты.
    """
    sheet = StyleSheet.gost()
    _apply_section(doc)
    _apply_language(doc)
    _configure_style(doc.styles[NORMAL_STYLE], sheet.normal)
    _create_paragraph_style(doc, CELL_STYLE, sheet.table_text)
    _apply_heading_styles(doc, sheet.headings)


def _apply_language(doc: Document) -> None:
    """Объявить язык документа русским.

    Шаблон python-docx задаёт в docDefaults язык en-US, и весь текст его
    наследует — Word проверяет русские слова по английскому словарю и
    подчёркивает их как ошибочные.
    """
    rpr = doc.styles.element.find(
        f"{qn('w:docDefaults')}/{qn('w:rPrDefault')}/{qn('w:rPr')}"
    )
    lang = rpr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rpr.append(lang)
    lang.set(qn("w:val"), "ru-RU")


def _apply_section(doc: Document) -> None:
    # Page margins per ГОСТ 7.32-2001: left ≥ 30mm, right ≥ 10mm, top ≥ 20mm, bottom ≥ 20mm
    section = doc.sections[0]
    section.left_margin = Cm(3.0)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)


def _apply_heading_styles(doc: Document, headings: list[ParagraphStyle]) -> None:
    for level, style in enumerate(headings):
        s = doc.styles[f"Heading {level + 1}"]
        _remove_theme_font_overrides(s)
        _configure_style(s, style)
        _remove_bottom_border(s)


def _create_paragraph_style(document: Document, name: str, style: ParagraphStyle):
    s = document.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    _configure_style(s, style)
    return s


def _configure_style(style_obj, style: ParagraphStyle) -> None:
    """Настраивает именованный стиль Word (font + paragraph_format) по *style*."""
    _apply_font(style_obj.font, style)
    _apply_parfmt(style_obj.paragraph_format, style)


def _apply_font(font, style: ParagraphStyle) -> None:
    font.name = style.font_name
    font.size = style.font_size
    if style.bold is not None:
        font.bold = style.bold
    if style.italic is not None:
        font.italic = style.italic
    if style.all_caps is not None:
        font.all_caps = style.all_caps
    if style.color is not None:
        font.color.rgb = style.color


def _apply_parfmt(parfmt, style: ParagraphStyle) -> None:
    if style.alignment is not None:
        parfmt.alignment = style.alignment
    if style.line_spacing is not None:
        parfmt.line_spacing_rule = style.line_spacing
    if style.first_line_indent is not None:
        parfmt.first_line_indent = style.first_line_indent
    if style.space_before is not None:
        parfmt.space_before = style.space_before
    if style.space_after is not None:
        parfmt.space_after = style.space_after


_GOST_TABLE_TEXT_KEY: tuple = ()  # заполняется ниже, после определения _style_key


def _style_key(style: ParagraphStyle) -> tuple:
    """Хешируемый ключ по содержимому стиля — для дедупликации именованных стилей."""
    return (
        style.font_name,
        int(style.font_size) if style.font_size is not None else None,
        style.bold,
        style.italic,
        style.all_caps,
        str(style.color) if style.color is not None else None,
        style.alignment,
        style.line_spacing,
        int(style.first_line_indent) if style.first_line_indent is not None else None,
        int(style.space_before) if style.space_before is not None else None,
        int(style.space_after) if style.space_after is not None else None,
    )


def _table_text_style_name(style: ParagraphStyle) -> str:
    key = _style_key(style)
    if key == _GOST_TABLE_TEXT_KEY:
        return CELL_STYLE  # совпадает с базовым «Table Text»
    digest = hashlib.md5(repr(key).encode()).hexdigest()[:8]
    return f"{CELL_STYLE} {digest}"


def _remove_theme_font_overrides(style) -> None:
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is not None:
        for attr in [qn("w:asciiTheme"), qn("w:hAnsiTheme"), qn("w:eastAsiaTheme"), qn("w:cstheme")]:
            rFonts.attrib.pop(attr, None)


def _remove_bottom_border(style) -> None:
    pPr = style.element.get_or_add_pPr()
    for pBdr in pPr.findall(qn("w:pBdr")):
        pPr.remove(pBdr)


_GOST_TABLE_TEXT_KEY = _style_key(StyleSheet.gost().table_text)
