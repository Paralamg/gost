from docx.document import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


def apply_gost_styles(doc: Document) -> None:
    """Configure GOST 7.32-2001 paragraph and heading styles on *doc* in-place."""
    _apply_section(doc)
    _apply_normal_style(doc)
    _apply_table_text_style(doc)
    _apply_heading_styles(doc)

def _apply_section(doc: Document) -> None:
    # Page margins per ГОСТ 7.32-2001: left ≥ 30mm, right ≥ 10mm, top ≥ 20mm, bottom ≥ 20mm
    section = doc.sections[0]
    section.left_margin = Cm(3.0)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)

def _apply_normal_style(doc: Document) -> None:
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(14)
    normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    normal.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    normal.paragraph_format.first_line_indent = Cm(1.25)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)


def _apply_table_text_style(doc: Document) -> None:
    table_text = doc.styles.add_style("Table Text", WD_STYLE_TYPE.PARAGRAPH)
    table_text.font.name = "Times New Roman"
    table_text.font.size = Pt(12)
    table_text.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    table_text.paragraph_format.first_line_indent = Cm(0)
    table_text.paragraph_format.space_before = Pt(0)
    table_text.paragraph_format.space_after = Pt(0)


def _apply_heading_styles(doc: Document) -> None:
    # (style_name, centered, all_caps, font_size, space_before)
    heading_configs = [
        ("Heading 1", True,  True,  Pt(14), Pt(0)),
        ("Heading 2", False, False, Pt(14), Pt(0)),
        ("Heading 3", False, False, Pt(14), Pt(0)),
        ("Heading 4", False, False, Pt(14), Pt(0)),
    ]
    for style_name, centered, all_caps, font_size, space_before in heading_configs:
        s = doc.styles[style_name]
        s.font.name = "Times New Roman"
        _remove_theme_font_overrides(s)
        s.font.size = font_size
        s.font.bold = True
        s.font.italic = False
        s.font.all_caps = all_caps
        s.font.color.rgb = RGBColor(0, 0, 0)
        s.paragraph_format.alignment = (
            WD_ALIGN_PARAGRAPH.CENTER if centered else WD_ALIGN_PARAGRAPH.LEFT
        )
        s.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        s.paragraph_format.first_line_indent = Cm(0) if centered else Cm(1.25)
        s.paragraph_format.space_before = space_before
        s.paragraph_format.space_after = Pt(0)
        _remove_bottom_border(s)


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