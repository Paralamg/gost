"""Сверка конвертера с эталоном Microsoft — MML2OMML.XSL.

Этот XSLT переводит MathML в OMML и входит в поставку Office. Использовать его в
самой библиотеке нельзя: файл проприетарный, прав на распространение Microsoft не
давала, и он есть только там, где установлен Office. Но как оракул при разработке
он бесценен — сверяет наш маппинг с тем, что считает правильным сам Word.

Тест пропускается, если Office не установлен, поэтому в CI он не мешает. Путь
можно задать переменной окружения GOST_MML2OMML_XSL.

Сверяется состав формулы, а не дерево: наш конвертер местами намеренно лучше
эталона — сворачивает парные скобки в растягивающийся m:d и ставит \\vec акцентом
m:acc, тогда как XSL оставляет и то, и другое обычным текстом.
"""

import os
from pathlib import Path

import latex2mathml.converter
import pytest
from docx.oxml.ns import qn
from lxml import etree

from gost.elements.formula.omml import latex_to_omml
from gost.elements.formula.operators import (
    ACCENT_CHARS,
    BAR_CHARS,
    CLOSING_CHARS,
    GROUP_CHARS,
    NARY_LIMITS,
    OPENING_CHARS,
)
from tests.corpus import FORMULAS

_CANDIDATES = [
    "/Applications/Microsoft Word.app/Contents/Resources/mathml2omml.xsl",
    r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL",
    r"C:\Program Files\Microsoft Office\Office16\MML2OMML.XSL",
]

# Служебные глифы: скобки, акценты, черты, обхваты, знаки сумм и интегралов.
# Мы храним их в атрибутах (m:begChr, m:chr) и берём комбинирующие формы, эталон
# же часто оставляет их обычным текстом в спейсинговой форме. Оба варианта
# правильны, поэтому из сверки они исключаются — сверяются операнды.
_STRUCTURAL = (
    OPENING_CHARS | CLOSING_CHARS | set(NARY_LIMITS) | BAR_CHARS
    | set(GROUP_CHARS) | set(ACCENT_CHARS) | set(ACCENT_CHARS.values())
)


def _find_reference() -> Path | None:
    override = os.environ.get("GOST_MML2OMML_XSL")
    candidates = [override, *_CANDIDATES] if override else _CANDIDATES
    for candidate in candidates:
        path = Path(candidate)
        if path.is_file():
            return path
    return None


_REFERENCE = _find_reference()

pytestmark = pytest.mark.skipif(
    _REFERENCE is None,
    reason="MML2OMML.XSL не найден — нужен установленный Microsoft Office",
)


@pytest.fixture(scope="module")
def reference():
    return etree.XSLT(etree.parse(str(_REFERENCE)))


def _operands(root) -> str:
    """Содержательные символы формулы: переменные, числа, знаки операций."""
    text = "".join(
        node.text or "" for node in root.iter() if node.tag == qn("m:t")
    )
    return "".join(sorted(
        char for char in text if not char.isspace() and char not in _STRUCTURAL
    ))


@pytest.mark.parametrize("latex", FORMULAS)
def test_content_matches_reference(latex, reference):
    """Ни один символ формулы не должен потеряться при переводе.

    Главный тихий отказ конвертера — молча проглоченный операнд, предел или
    ячейка матрицы: XML остаётся валидным, а формула в документе неверна.
    """
    mathml = latex2mathml.converter.convert(latex)
    expected = _operands(reference(etree.fromstring(mathml)).getroot())
    actual = _operands(etree.fromstring(latex_to_omml(latex).encode()))
    assert actual == expected


@pytest.mark.parametrize("latex", FORMULAS)
def test_reference_itself_converts(latex, reference):
    """Страховка: расхождение выше — наша вина, а не сломанный эталон."""
    mathml = latex2mathml.converter.convert(latex)
    assert reference(etree.fromstring(mathml)).getroot() is not None
