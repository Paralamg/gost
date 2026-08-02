"""Конвертер LaTeX → OMML."""

import pytest
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from lxml import etree

from gost.elements.formula.omml import latex_to_omml
from tests.corpus import FORMULAS


def convert(latex: str):
    """Переводит формулу и отдаёт разобранное дерево OMML."""
    return parse_xml(latex_to_omml(latex))


def tags(node) -> set[str]:
    return {etree.QName(child).localname for child in node.iter()}


@pytest.mark.parametrize("latex", FORMULAS)
def test_corpus_parses_as_omml(latex):
    """Каждая формула набора даёт корректный XML, который примет Word.

    Обе готовые библиотеки заваливали именно этот тест: mathml2omml отдавал
    битый XML на \\vec и \\underbrace, docx-equation падал на \\sqrt{\\frac{a}{b}}.
    """
    root = convert(latex)
    assert root.tag == qn("m:oMath")
    assert len(root) > 0, "формула не должна быть пустой"


@pytest.mark.parametrize("latex", FORMULAS)
def test_corpus_declares_math_namespace(latex):
    """Без объявления xmlns:m Word не распознает формулу."""
    assert 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"' \
           in latex_to_omml(latex)


def test_fraction():
    root = convert(r"\frac{a}{b}")
    assert root.find(qn("m:f")) is not None
    assert root.find(f"{qn('m:f')}/{qn('m:num')}") is not None
    assert root.find(f"{qn('m:f')}/{qn('m:den')}") is not None


def test_binomial_has_no_bar():
    """\\binom — дробь без черты, в растягивающихся скобках."""
    root = convert(r"\binom{n}{k}")
    kind = root.find(f"{qn('m:d')}/{qn('m:e')}/{qn('m:f')}/{qn('m:fPr')}/{qn('m:type')}")
    assert kind.get(qn("m:val")) == "noBar"


def test_square_root_hides_degree():
    root = convert(r"\sqrt{x}")
    hide = root.find(f"{qn('m:rad')}/{qn('m:radPr')}/{qn('m:degHide')}")
    assert hide.get(qn("m:val")) == "on"


def test_root_keeps_degree():
    root = convert(r"\sqrt[3]{x}")
    degree = root.find(f"{qn('m:rad')}/{qn('m:deg')}")
    assert "".join(degree.itertext()) == "3"


def test_sum_becomes_nary_not_subscript():
    """Знак суммы с пределами — m:nary, а не обычные индексы."""
    root = convert(r"\sum_{i=1}^{n} a_i")
    nary = root.find(qn("m:nary"))
    assert nary is not None
    assert nary.find(f"{qn('m:naryPr')}/{qn('m:chr')}").get(qn("m:val")) == "∑"
    assert qn("m:sSubSup") not in [child.tag for child in root]


def limit_location(latex: str) -> str:
    return convert(latex).find(
        f".//{qn('m:nary')}/{qn('m:naryPr')}/{qn('m:limLoc')}"
    ).get(qn("m:val"))


def test_sum_limits_go_above_and_below():
    """У суммы пределы стоят над и под знаком."""
    assert limit_location(r"\sum_{i=1}^{n} a_i") == "undOvr"
    assert limit_location(r"\prod_{i=1}^{n} x_i") == "undOvr"


def test_integral_limits_go_beside():
    """У интеграла пределы, наоборот, сбоку — так принято и так делает Word."""
    assert limit_location(r"\int_{0}^{1} f") == "subSup"


def test_limit_location_survives_nesting():
    """Внутри дроби latex2mathml переключается на строчный стиль.

    Расположение пределов задаётся знаком, а не режимом разбора, иначе сумма в
    числителе дроби получила бы пределы сбоку.
    """
    assert limit_location(r"\frac{\sum_{i=1}^{n} a_i}{n}") == "undOvr"


def test_nary_takes_following_expression_as_operand():
    """Выражение за знаком суммы переезжает внутрь m:e.

    MathML держит операнд соседом. Если оставить m:e пустым, Word всё равно
    отводит под него место и между знаком и выражением остаётся провал.
    """
    root = convert(r"\sum_{i=1}^{n} a_i b_i")
    operand = root.find(f"{qn('m:nary')}/{qn('m:e')}")
    assert len(operand) > 0
    assert "".join(operand.itertext()) == "aibi"
    # Операнд не должен остаться ещё и снаружи.
    assert root.find(qn("m:sSub")) is None


def test_integral_becomes_nary():
    root = convert(r"\int_{0}^{1} f")
    assert root.find(f"{qn('m:nary')}/{qn('m:naryPr')}/{qn('m:chr')}").get(qn("m:val")) == "∫"


def test_plain_subscript_stays_subscript():
    """Обычная переменная с индексом не должна превращаться в n-арный оператор."""
    root = convert(r"a_i")
    assert root.find(qn("m:sSub")) is not None
    assert root.find(qn("m:nary")) is None


def test_fences_become_delimiter():
    """Парные скобки сворачиваются в m:d, иначе они не растянутся по высоте дроби."""
    root = convert(r"\left( \frac{a}{b} \right)")
    delimiter = root.find(qn("m:d"))
    assert delimiter is not None
    properties = delimiter.find(qn("m:dPr"))
    assert properties.find(qn("m:begChr")).get(qn("m:val")) == "("
    assert properties.find(qn("m:endChr")).get(qn("m:val")) == ")"
    assert delimiter.find(f"{qn('m:e')}/{qn('m:f')}") is not None


def test_matrix_brackets_become_delimiter():
    """У pmatrix скобки приходят голыми <mo> вокруг таблицы — их тоже надо свернуть."""
    root = convert(r"\begin{pmatrix} a & b \\ c & d \end{pmatrix}")
    delimiter = root.find(qn("m:d"))
    assert delimiter is not None
    assert delimiter.find(f"{qn('m:e')}/{qn('m:m')}") is not None


def test_power_applies_to_bracket_group():
    """В «(a+b)^2» степень стоит на скобках целиком, а не на закрывающей скобке.

    latex2mathml прячет закрывающую скобку в базу <msup>, поэтому пара скобок
    находится не сразу: без обработки они остаются обычным текстом, степень
    повисает на «)», и формула читается неверно.
    """
    root = convert(r"(a + b)^2")
    power = root.find(qn("m:sSup"))
    assert power is not None
    assert power.find(f"{qn('m:e')}/{qn('m:d')}") is not None
    assert "".join(power.find(qn("m:sup")).itertext()) == "2"
    # Закрывающая скобка ушла в атрибут m:endChr, отдельным текстом её быть не должно.
    assert ")" not in "".join(root.itertext())


def test_subscript_applies_to_bracket_group():
    root = convert(r"(a + b)_n")
    assert root.find(f"{qn('m:sSub')}/{qn('m:e')}/{qn('m:d')}") is not None


def test_cases_opens_without_closing_brace():
    """\\begin{cases} даёт открывающую фигурную скобку без пары."""
    root = convert(r"\begin{cases} x, & x > 0 \\ 0, & x \leq 0 \end{cases}")
    properties = root.find(f"{qn('m:d')}/{qn('m:dPr')}")
    assert properties.find(qn("m:begChr")).get(qn("m:val")) == "{"
    assert properties.find(qn("m:endChr")).get(qn("m:val")) == ""


def test_matrix_rows_and_cells():
    root = convert(r"\begin{bmatrix} 1 & 0 \\ 0 & 1 \end{bmatrix}")
    matrix = root.find(f"{qn('m:d')}/{qn('m:e')}/{qn('m:m')}")
    rows = matrix.findall(qn("m:mr"))
    assert len(rows) == 2
    assert all(len(row.findall(qn("m:e"))) == 2 for row in rows)


def test_accent_uses_combining_character():
    """Акцент задаётся комбинирующим знаком: только он встаёт НАД буквой.

    latex2mathml отдаёт отдельно стоящие «→» и «^»; если записать их в m:chr
    как есть, Word нарисует стрелку и крышку рядом с переменной, а не над ней.
    """
    accent = convert(r"\vec{n}").find(qn("m:acc"))
    assert accent.find(f"{qn('m:accPr')}/{qn('m:chr')}").get(qn("m:val")) == "⃗"

    accent = convert(r"\hat{y}").find(qn("m:acc"))
    assert accent.find(f"{qn('m:accPr')}/{qn('m:chr')}").get(qn("m:val")) == "̂"


def test_overline_is_bar_not_accent():
    """\\overline тянется по всей базе, поэтому m:bar, а не акцент над буквой."""
    root = convert(r"\overline{AB}")
    bar = root.find(qn("m:bar"))
    assert bar.find(f"{qn('m:barPr')}/{qn('m:pos')}").get(qn("m:val")) == "top"
    assert "AB" in "".join(bar.itertext())


def test_group_char():
    """\\underbrace — m:groupChr с подписью снизу."""
    root = convert(r"\underbrace{a + b}_{c}")
    group = root.find(f"{qn('m:limLow')}/{qn('m:e')}/{qn('m:groupChr')}")
    properties = group.find(qn("m:groupChrPr"))
    assert properties.find(qn("m:chr")).get(qn("m:val")) == "⏟"
    assert properties.find(qn("m:vertJc")).get(qn("m:val")) == "bot"


def test_limit_under():
    root = convert(r"\lim_{x \to 0} f")
    assert "sSub" in tags(root) or "limLow" in tags(root)


def test_function_name_is_upright():
    """sin, log, lim — прямым шрифтом, иначе Word прочтёт их как произведение букв."""
    root = convert(r"\sin x")
    styles = root.findall(f".//{qn('m:r')}/{qn('m:rPr')}/{qn('m:sty')}")
    assert any(style.get(qn("m:val")) == "p" for style in styles)


def test_single_letter_variable_stays_italic():
    """У переменной начертание по умолчанию — курсив, m:sty не нужен."""
    root = convert(r"x")
    assert root.find(f"{qn('m:r')}/{qn('m:rPr')}") is None


def test_text_is_not_math():
    """\\text{...} — обычный текст, а не набор переменных."""
    root = convert(r"K_{\text{сум}}")
    assert root.find(f".//{qn('m:rPr')}/{qn('m:nor')}") is not None
    assert "сум" in "".join(root.itertext())


def test_cyrillic_survives():
    root = convert(r"P_{вх}")
    assert "вх" in "".join(root.itertext())


def test_nested_root_does_not_recurse():
    """docx-equation уходил здесь в RecursionError."""
    root = convert(r"\sqrt{\frac{a}{b}}")
    assert root.find(f"{qn('m:rad')}/{qn('m:e')}/{qn('m:f')}") is not None


@pytest.mark.parametrize("latex", ["", "   "])
def test_empty_formula_rejected(latex):
    with pytest.raises(ValueError, match="пуст"):
        latex_to_omml(latex)


def test_broken_latex_rejected():
    with pytest.raises(ValueError):
        latex_to_omml(r"\frac{")
