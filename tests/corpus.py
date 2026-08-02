"""Набор формул, на котором проверяется конвертер LaTeX → OMML.

Покрывает конструкции, встречающиеся в отчётах по ГОСТ, и все места, где
спотыкались готовые библиотеки: вложения под корнем, парные скобки, n-арные
операторы с пределами, скобки-обхваты, акценты.
"""

FORMULAS = [
    # дроби и степени
    r"E = \frac{mv^2}{2}",
    r"e^{i\pi} + 1 = 0",
    r"\binom{n}{k}",
    # радикалы, в том числе с вложениями — на них падал docx-equation
    r"\sqrt{x}",
    r"\sqrt{\frac{a}{b}}",
    r"\sqrt{\bar{x}}",
    r"\sqrt[3]{x}",
    r"\sigma = \sqrt{\frac{\sum_{i=1}^{n}(x_i - \bar{x})^2}{n - 1}}",
    # n-арные операторы с пределами
    r"S = \sum_{i=1}^{n} a_i b_i",
    r"\prod_{i=1}^{n} x_i",
    r"\int_{0}^{\infty} e^{-x^2} dx",
    r"\iint_D f\,dx\,dy",
    # матрицы и системы
    r"A = \begin{pmatrix} a & b \\ c & d \end{pmatrix}",
    r"\begin{bmatrix} 1 & 0 \\ 0 & 1 \end{bmatrix}",
    r"\begin{vmatrix} a & b \\ c & d \end{vmatrix}",
    r"f(x) = \begin{cases} x, & x > 0 \\ 0, & x \leq 0 \end{cases}",
    # скобки, в том числе со степенью на всей группе
    r"\left( \frac{a}{b} \right)",
    r"(a + b)^2",
    r"(a + b)_n",
    r"\left[ \frac{x}{y} \right]^{n+1}",
    # акценты и обхваты — на них падал mathml2omml
    r"\overline{AB} \perp \vec{n}",
    r"\underbrace{a + b}_{c}",
    r"\hat{y} = \tilde{a} + \dot{b}",
    # функции, пределы, производные
    r"\lim_{x \to 0} \frac{\sin x}{x} = 1",
    r"\log_{2} n",
    r"\frac{\partial^2 u}{\partial x^2}",
    r"\partial_x u",
    # символы и кириллица в индексах
    r"\alpha + \beta \geq \gamma \cdot \Delta",
    r"x \approx y \neq z \pm \mp w",
    r"K_{\text{сум}} = \frac{P_{вх}}{P_{вых}}",
]
