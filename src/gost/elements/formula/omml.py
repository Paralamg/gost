"""Перевод формулы из LaTeX в OMML — разметку формул Word.

Word не понимает LaTeX на уровне формата: внутри ``.docx`` формула хранится только
как OMML (``<m:oMath>``). Режим «Уравнение → LaTeX» в интерфейсе конвертирует ввод
в OMML в момент набора, в файл ложится уже OMML.

Путь конвертации: LaTeX → MathML (``latex2mathml``) → OMML (этот модуль). Разбор
LaTeX отдан библиотеке — там накоплены сотни краевых случаев; MathML → OMML это
механический маппинг тегов, и он здесь.

MathML не различает роли символов: знак суммы, скобка и знак вектора приходят
одним тегом ``<mo>``. Роль определяется самим символом — таблицы в
:mod:`gost.elements.formula.operators`.
"""

import logging

import latex2mathml.converter
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree

from gost.elements.formula.operators import (
    ACCENT_CHARS,
    BAR_CHARS,
    BRACKET_PAIRS,
    GROUP_CHARS,
    NARY_LIMITS,
    OPENING_CHARS,
)

logger = logging.getLogger(__name__)

MATHML_NS = "http://www.w3.org/1998/Math/MathML"

# Теги, не несущие оформления: разворачиваются в своих детей.
_TRANSPARENT = frozenset({"mrow", "mstyle", "mpadded", "mphantom"})

# Теги без визуального вклада.
_IGNORED = frozenset({"mspace", "none"})

# Теги, навешивающие на базу индекс и степень.
_SCRIPTS = frozenset({"msub", "msup", "msubsup"})


def latex_to_omml(latex: str) -> str:
    """Переводит формулу из LaTeX в строку OMML.

    Возвращается строка, а не элемент lxml: один и тот же элемент нельзя вставить
    в два дерева, а ``Formula.render`` вызывается несколько раз — автоподбор
    разрывов таблиц рендерит элементы в черновые документы.

    Raises:
        ValueError: LaTeX не разобрался.
    """
    if not latex or not latex.strip():
        raise ValueError("Формула пустая")

    try:
        # display="block" — формула выключная, на отдельной строке. От этого
        # зависит вёрстка пределов: у выключной формулы они встают над и под
        # знаком суммы, а не сбоку от него, как в формуле внутри строки.
        mathml = latex2mathml.converter.convert(latex, display="block")
    except Exception as error:
        raise ValueError(f"Не удалось разобрать формулу LaTeX: {latex!r}") from error

    try:
        root = etree.fromstring(mathml.encode())
    except etree.XMLSyntaxError as error:
        raise ValueError(f"latex2mathml вернул некорректный MathML для {latex!r}") from error

    math = OxmlElement("m:oMath")
    for node in _convert_sequence(list(root)):
        math.append(node)
    # OxmlElement объявляет xmlns:m на корне, поэтому строку можно сразу отдать
    # в parse_xml при вставке в документ.
    return etree.tostring(math, encoding="unicode")


def _convert_sequence(elements: list) -> list:
    """Переводит подряд идущие узлы MathML, сворачивая парные скобки в m:d.

    latex2mathml отдаёт скобки соседями (``<mo>(</mo> … <mo>)</mo>``), а не
    оборачивающим элементом, поэтому пары ищутся здесь. Без этого скобки остались
    бы обычным текстом и не растягивались бы по высоте дроби или матрицы.
    """
    result = []
    index = 0
    while index < len(elements):
        element = elements[index]
        char = _operator_char(element)
        if char in OPENING_CHARS:
            closing_index = _find_closing(elements, index, char)
            if closing_index is not None:
                inner = _convert_sequence(elements[index + 1:closing_index])
                delimiter = _delimiter(char, BRACKET_PAIRS[char], inner)
                # У «(a+b)^2» степень висит на закрывающей скобке — после сворачивания
                # она должна перейти на скобки целиком.
                result.append(_apply_script(elements[closing_index], delimiter))
                index = closing_index + 1
                continue
            if _is_fence(element):
                # \begin{cases} даёт открывающую фигурную скобку без закрывающей.
                inner = _convert_sequence(elements[index + 1:])
                result.append(_delimiter(char, "", inner))
                break
        nodes = _convert(element)
        result.extend(nodes)
        index += 1
        # MathML не вкладывает операнд в знак суммы — тот идёт следом. Если
        # оставить m:e пустым, Word всё равно отведёт под него место, и между
        # знаком и выражением зияет пустота. Поэтому остаток выражения
        # переезжает в операнд — так же собирает сумму и сам редактор Word.
        operand = _empty_nary_operand(nodes)
        if operand is not None:
            for node in _convert_sequence(elements[index:]):
                operand.append(node)
            break
    return result


def _empty_nary_operand(nodes: list):
    """m:e n-арного оператора, если он один и его операнд ещё пуст; иначе None."""
    if len(nodes) != 1 or nodes[0].tag != qn("m:nary"):
        return None
    operand = nodes[0].find(qn("m:e"))
    return operand if len(operand) == 0 else None


def _find_closing(elements: list, start: int, opening: str) -> int | None:
    """Индекс парной закрывающей скобки с учётом вложенности; None — пары нет."""
    closing = BRACKET_PAIRS[opening]
    depth = 0
    for index in range(start + 1, len(elements)):
        char = _bracket_char(elements[index])
        if char == closing and depth == 0:
            return index
        # У ‖ и | открывающая и закрывающая совпадают — вложенность не считаем.
        if opening != closing:
            if char == opening:
                depth += 1
            elif char == closing:
                depth -= 1
    return None


def _bracket_char(element) -> str:
    """Символ скобки, в том числе спрятанной под индекс или степень.

    В «(a+b)^2» latex2mathml отдаёт закрывающую скобку не соседом, а базой
    ``<msup>``. Без этого пара не находилась бы и скобки оставались обычным
    текстом, не растягиваясь по высоте содержимого.
    """
    char = _operator_char(element)
    if char:
        return char
    if etree.QName(element).localname in _SCRIPTS:
        children = list(element)
        if children:
            return _operator_char(children[0])
    return ""


def _apply_script(element, node):
    """Переносит индекс и степень закрывающей скобки на свёрнутый m:d."""
    tag = etree.QName(element).localname
    if tag not in _SCRIPTS:
        return node
    children = list(element)[1:]
    scripts = [_convert(children[i]) if i < len(children) else [] for i in range(2)]
    if tag == "msub":
        return _composite("m:sSub", ("m:e", [node]), ("m:sub", scripts[0]))
    if tag == "msup":
        return _composite("m:sSup", ("m:e", [node]), ("m:sup", scripts[0]))
    return _composite(
        "m:sSubSup", ("m:e", [node]), ("m:sub", scripts[0]), ("m:sup", scripts[1])
    )


def _convert(element) -> list:
    """Переводит один узел MathML в список узлов OMML."""
    tag = etree.QName(element).localname

    if tag in _IGNORED:
        return []
    if tag in _TRANSPARENT:
        return _convert_sequence(list(element))

    handler = _HANDLERS.get(tag)
    if handler is None:
        logger.warning("Формула: тег MathML %s не поддержан, выведен как текст", tag)
        return _text_fallback(element)
    return handler(element)


# --- листья -----------------------------------------------------------------

def _identifier(element) -> list:
    """Переменная или имя функции."""
    text = _text_of(element)
    # Имя функции (sin, log, lim) — несколько букв подряд; курсив превратил бы его
    # в произведение переменных, поэтому прямой шрифт.
    return [_run(text, upright=len(text) > 1)]


def _number(element) -> list:
    """Число: прямое начертание у него и так по умолчанию."""
    return [_run(_text_of(element))]


def _operator(element) -> list:
    """Знак операции, не распознанный как скобка, акцент или n-арный оператор."""
    return [_run(_text_of(element))]


def _literal_text(element) -> list:
    """\\text{...} — обычный текст внутри формулы, без математического начертания."""
    return [_run(_text_of(element), literal=True)]


# --- составные конструкции ---------------------------------------------------

def _fraction(element) -> list:
    """Дробь m:f — с чертой или, у \\binom, без неё."""
    parts = _args(element, 2)
    node = OxmlElement("m:f")
    if element.get("linethickness") == "0":
        # \binom и \choose: дробь без черты.
        node.append(_properties("m:fPr", _val("m:type", "noBar")))
    node.append(_arg("m:num", parts[0]))
    node.append(_arg("m:den", parts[1]))
    return [node]


def _square_root(element) -> list:
    """Квадратный корень: тот же m:rad, но с пустым и скрытым показателем."""
    node = OxmlElement("m:rad")
    node.append(_properties("m:radPr", _val("m:degHide", "on")))
    node.append(OxmlElement("m:deg"))
    node.append(_arg("m:e", _convert_sequence(list(element))))
    return [node]


def _root(element) -> list:
    """Корень с показателем: \\sqrt[3]{x}."""
    parts = _args(element, 2)
    node = OxmlElement("m:rad")
    node.append(_properties("m:radPr", _val("m:degHide", "off")))
    node.append(_arg("m:deg", parts[1]))
    node.append(_arg("m:e", parts[0]))
    return [node]


def _subscript(element) -> list:
    """Нижний индекс — либо нижний предел, если база это знак суммы."""
    parts = _args(element, 2)
    nary = _nary_from(element, sub=parts[1], sup=None)
    if nary is not None:
        return nary
    return [_composite("m:sSub", ("m:e", parts[0]), ("m:sub", parts[1]))]


def _superscript(element) -> list:
    """Верхний индекс — либо верхний предел, если база это знак суммы."""
    parts = _args(element, 2)
    nary = _nary_from(element, sub=None, sup=parts[1])
    if nary is not None:
        return nary
    return [_composite("m:sSup", ("m:e", parts[0]), ("m:sup", parts[1]))]


def _sub_superscript(element) -> list:
    """Оба индекса сразу — либо оба предела у знака суммы или интеграла."""
    parts = _args(element, 3)
    nary = _nary_from(element, sub=parts[1], sup=parts[2])
    if nary is not None:
        return nary
    return [_composite("m:sSubSup", ("m:e", parts[0]), ("m:sub", parts[1]), ("m:sup", parts[2]))]


def _under(element) -> list:
    """Знак под базой: предел \\lim, черта снизу, нижняя скобка-обхват."""
    parts = _args(element, 2)
    return _limit(element, base=parts[0], mark_index=1, above=False)


def _over(element) -> list:
    """Знак над базой: акцент \\vec и \\hat, черта \\overline, верхняя скобка."""
    parts = _args(element, 2)
    return _limit(element, base=parts[0], mark_index=1, above=True)


def _under_over(element) -> list:
    """Знаки и снизу, и сверху: пределы n-арного оператора либо два m:lim."""
    parts = _args(element, 3)
    nary = _nary_from(element, sub=parts[1], sup=parts[2])
    if nary is not None:
        return nary
    inner = _composite("m:limLow", ("m:e", parts[0]), ("m:lim", parts[1]))
    return [_composite("m:limUpp", ("m:e", [inner]), ("m:lim", parts[2]))]


def _limit(element, base: list, mark_index: int, above: bool) -> list:
    """Собирает акцент, скобку-обхват или предел для munder/mover."""
    children = list(element)
    mark = children[mark_index] if mark_index < len(children) else None
    char = _operator_char(mark) if mark is not None else ""

    nary = _nary_from(
        element,
        sub=None if above else _convert(children[mark_index]),
        sup=_convert(children[mark_index]) if above else None,
    )
    if nary is not None:
        return nary

    if char in GROUP_CHARS:
        node = OxmlElement("m:groupChr")
        node.append(_properties(
            "m:groupChrPr", _val("m:chr", char), _val("m:vertJc", GROUP_CHARS[char])
        ))
        node.append(_arg("m:e", base))
        return [node]

    if char in BAR_CHARS:
        # Черта тянется по всей ширине базы — это m:bar, а не акцент над буквой.
        node = OxmlElement("m:bar")
        node.append(_properties("m:barPr", _val("m:pos", "top" if above else "bot")))
        node.append(_arg("m:e", base))
        return [node]

    if char in ACCENT_CHARS or element.get("accent") == "true":
        if above:
            node = OxmlElement("m:acc")
            node.append(_properties("m:accPr", _val("m:chr", ACCENT_CHARS.get(char, char))))
            node.append(_arg("m:e", base))
            return [node]
        # m:acc рисует знак только сверху, поэтому снизу — через m:bar.
        node = OxmlElement("m:bar")
        node.append(_properties("m:barPr", _val("m:pos", "bot")))
        node.append(_arg("m:e", base))
        return [node]

    tag = "m:limUpp" if above else "m:limLow"
    return [_composite(tag, ("m:e", base), ("m:lim", _convert(children[mark_index])))]


def _matrix(element) -> list:
    """Матрица m:m. Скобки вокруг неё сворачивает _convert_sequence — они соседи."""
    rows = [child for child in element if etree.QName(child).localname == "mtr"]
    node = OxmlElement("m:m")
    alignment = _column_alignment(rows)
    if alignment is not None:
        node.append(_matrix_properties(alignment, columns=_column_count(rows)))
    for row in rows:
        row_node = OxmlElement("m:mr")
        for cell in row:
            row_node.append(_arg("m:e", _convert_sequence(list(cell))))
        node.append(row_node)
    return [node]


def _column_alignment(rows: list) -> str | None:
    """Выравнивание столбцов матрицы или None, если оно и так по умолчанию."""
    for row in rows:
        for cell in row:
            align = cell.get("columnalign")
            # По умолчанию Word центрирует; отличие задаёт только \begin{cases}.
            if align and align != "center":
                return align
    return None


def _column_count(rows: list) -> int:
    """Число столбцов матрицы — по самой длинной строке."""
    return max((len(list(row)) for row in rows), default=1)


def _matrix_properties(alignment: str, columns: int):
    """m:mPr: одинаковое выравнивание для всех столбцов матрицы."""
    column = OxmlElement("m:mc")
    column.append(_properties(
        "m:mcPr", _val("m:count", str(columns)), _val("m:mcJc", alignment)
    ))
    columns_node = OxmlElement("m:mcs")
    columns_node.append(column)
    properties = OxmlElement("m:mPr")
    properties.append(columns_node)
    return properties


# --- вспомогательное ---------------------------------------------------------

def _nary_from(element, sub: list | None, sup: list | None) -> list | None:
    """Собирает m:nary, если база — знак суммы, интеграла и т. п.

    Возвращает ``None``, если база обычная: тогда вызывающий строит индексы.
    """
    children = list(element)
    if not children:
        return None
    char = _operator_char(children[0])
    if char not in NARY_LIMITS:
        return None

    node = OxmlElement("m:nary")
    node.append(_properties(
        "m:naryPr",
        _val("m:chr", char),
        _val("m:limLoc", NARY_LIMITS[char]),
        _val("m:subHide", "off" if sub else "on"),
        _val("m:supHide", "off" if sup else "on"),
    ))
    node.append(_arg("m:sub", sub or []))
    node.append(_arg("m:sup", sup or []))
    # Операнд в MathML не вложен в оператор, он идёт следом — m:e остаётся пустым.
    node.append(OxmlElement("m:e"))
    return [node]


def _composite(tag: str, *parts: tuple[str, list]):
    """Конструкция OMML из именованных частей: («m:e», база), («m:sup», степень)."""
    node = OxmlElement(tag)
    for name, children in parts:
        node.append(_arg(name, children))
    return node


def _arg(tag: str, children: list):
    """Одна часть конструкции: тег с вложенными в него узлами."""
    node = OxmlElement(tag)
    for child in children:
        node.append(child)
    return node


def _args(element, count: int) -> list[list]:
    """Переводит первые *count* детей, каждого — в свой список узлов OMML."""
    children = list(element)
    return [_convert(children[i]) if i < len(children) else [] for i in range(count)]


def _run(text: str, *, upright: bool = False, literal: bool = False):
    """Текстовый фрагмент формулы.

    По умолчанию начертание математическое — переменные выходят курсивом.
    ``upright`` даёт прямой шрифт имени функции, ``literal`` — обычный текст.
    """
    node = OxmlElement("m:r")
    if literal:
        # \text{...} — обычный текст абзаца, без математического начертания.
        node.append(_properties("m:rPr", OxmlElement("m:nor")))
    elif upright:
        node.append(_properties("m:rPr", _val("m:sty", "p")))
    text_node = OxmlElement("m:t")
    if text != text.strip():
        text_node.set(qn("xml:space"), "preserve")
    text_node.text = text
    node.append(text_node)
    return node


def _properties(tag: str, *children):
    """Блок настроек конструкции: m:fPr, m:radPr, m:naryPr и подобные."""
    node = OxmlElement(tag)
    for child in children:
        node.append(child)
    return node


def _val(tag: str, value: str):
    """Одна настройка: тег со значением в атрибуте m:val."""
    node = OxmlElement(tag)
    node.set(qn("m:val"), value)
    return node


def _delimiter(opening: str, closing: str, children: list):
    """Содержимое в скобках m:d — они растягиваются по его высоте.

    Пустая *closing* оставляет группу без закрывающей скобки, как \\begin{cases}.
    """
    node = OxmlElement("m:d")
    node.append(_properties(
        "m:dPr", _val("m:begChr", opening), _val("m:endChr", closing)
    ))
    node.append(_arg("m:e", children))
    return node


def _operator_char(element) -> str:
    """Символ, если узел — одиночный оператор ``<mo>``; иначе пустая строка."""
    if element is None or etree.QName(element).localname != "mo":
        return ""
    text = _text_of(element)
    return text if len(text) == 1 else ""


def _is_fence(element) -> bool:
    """Скобка ли это, растягиваемая по содержимому: \\left, \\begin{cases}."""
    return element.get("fence") == "true" or element.get("stretchy") == "true"


def _text_of(element) -> str:
    """Собственный текст узла без окружающих пробелов."""
    return (element.text or "").strip()


def _text_fallback(element) -> list:
    """Запасной перевод неподдержанного тега: весь его текст одним фрагментом."""
    text = "".join(element.itertext()).strip()
    return [_run(text)] if text else []


# Тег MathML -> функция, переводящая его в OMML. Всё, чего здесь нет,
# выводится текстом с предупреждением в лог.
_HANDLERS = {
    "mi": _identifier,
    "mn": _number,
    "mo": _operator,
    "mtext": _literal_text,
    "ms": _literal_text,
    "mfrac": _fraction,
    "msqrt": _square_root,
    "mroot": _root,
    "msub": _subscript,
    "msup": _superscript,
    "msubsup": _sub_superscript,
    "munder": _under,
    "mover": _over,
    "munderover": _under_over,
    "mtable": _matrix,
}
