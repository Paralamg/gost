from collections.abc import Mapping

from docx.document import Document
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.shared import Emu

from gost.elements.element import NumberedElement
from gost.elements.formula.omml import latex_to_omml
from gost.inline import render_inline
from gost.styles import ParagraphStyle, apply_paragraph_style

# Разделитель формулы и её номера в линейном формате Word. При вёрстке не
# показывается — Word использует его как точку разгона строки.
SEPARATOR = " #"

WHERE = "где "
# Тире то же, что в подписях таблиц и рисунков.
DASH = "–"


class Formula(NumberedElement):
    """Формула по ГОСТ 7.32-2017 (6.11).

    Формула выводится отдельной строкой по центру, номер — в круглых скобках
    у правого края. Ниже при необходимости идёт расшифровка обозначений.

    Формула задаётся на LaTeX и переводится в OMML — собственный формат формул
    Word. В документе это полноценный объект «Уравнение»: его можно править в
    Word, он масштабируется и печатается без потери качества.

    Args:
        index: Номер формулы. Пустая строка — формула не нумеруется.
        latex: Формула на LaTeX, например ``r"E = \\frac{mv^2}{2}"``.
        where: Расшифровка обозначений в порядке их появления в формуле.
            Ключи и значения понимают инлайн-разметку.
        formula_style: Оформление строки с формулой.
        note_style: Оформление строк расшифровки.

    Raises:
        ValueError: LaTeX не разобрался.
    """

    def __init__(
            self,
            index: str,
            latex: str,
            *,
            where: Mapping[str, str] | None = None,
            formula_style: ParagraphStyle,
            note_style: ParagraphStyle,
    ) -> None:
        super().__init__(index)
        self.latex = latex
        # Конвертация здесь, а не в render: ошибка в формуле должна всплывать при
        # создании элемента, а не при сохранении документа. Заодно не повторяется
        # на каждый повторный рендер при автоподборе разрывов таблиц.
        self.omml = latex_to_omml(latex)
        self.where = dict(where) if where else {}
        self.formula_style = formula_style
        self.note_style = note_style

    def render(self, document: Document) -> None:
        paragraph = document.add_paragraph()
        # m:oMathPara — сосед w:r внутри w:p, а не его потомок. Абзац при этом
        # обязан содержать только формулу: увидев рядом обычные руны, Word
        # считает формулу строчной и сжимает её — знак суммы мельчает, пределы
        # съезжают вбок. Поэтому и номер живёт внутри формулы, а не рядом.
        paragraph._p.append(self.__build_math())
        apply_paragraph_style(paragraph, self.formula_style)

        if self.where:
            # Расшифровка продолжает предложение формулы, поэтому отбивка снизу
            # переносится на её последнюю строку.
            paragraph.paragraph_format.space_after = Emu(0)
            paragraph.paragraph_format.keep_with_next = True
            self.__render_where(document)

    def __build_math(self):
        """Собирает m:oMathPara — выключную формулу, при необходимости с номером.

        Так же кодирует нумерованную формулу сам Word, когда в неё дописывают
        «#(N)»: формула и номер лежат в одной строке m:eqArr с maxDist, который
        разгоняет их к противоположным краям. Обёртка m:oMathPara включает
        выключной режим — крупный знак суммы и пределы над и под ним.
        """
        math = parse_xml(self.omml)
        if self.index:
            math.append(self.__numbered_row(math))
        para = OxmlElement("m:oMathPara")
        para.append(math)
        return para

    def __numbered_row(self, math):
        row = OxmlElement("m:e")
        # Формула переезжает из корня в строку — там её и ждёт m:eqArr.
        for node in list(math):
            row.append(node)
        row.append(_run(SEPARATOR))
        row.append(_number(self.index))

        properties = OxmlElement("m:eqArrPr")
        properties.append(_val("m:maxDist", "1"))
        equations = OxmlElement("m:eqArr")
        equations.append(properties)
        equations.append(row)
        return equations

    def __render_where(self, document: Document) -> None:
        items = list(self.where.items())
        for position, (symbol, description) in enumerate(items):
            prefix = WHERE if position == 0 else ""
            ending = "." if position == len(items) - 1 else ";"
            paragraph = document.add_paragraph()
            render_inline(paragraph, f"{prefix}{symbol} {DASH} {description}{ending}")
            apply_paragraph_style(paragraph, self.note_style)


def _run(text: str, *, literal: bool = False):
    node = OxmlElement("m:r")
    if literal:
        properties = OxmlElement("m:rPr")
        properties.append(OxmlElement("m:nor"))
        node.append(properties)
    text_node = OxmlElement("m:t")
    text_node.set(qn("xml:space"), "preserve")
    text_node.text = text
    node.append(text_node)
    return node


def _number(index: str):
    """Номер формулы — объект m:d, круглые скобки у него по умолчанию.

    Номер помечается обычным текстом: иначе математический движок сочтёт точку
    в «А.1» знаком операции и разгонит номер пробелами — «(А. 1)».
    """
    content = OxmlElement("m:e")
    content.append(_run(index, literal=True))
    node = OxmlElement("m:d")
    node.append(content)
    return node


def _val(tag: str, value: str):
    node = OxmlElement(tag)
    node.set(qn("m:val"), value)
    return node
