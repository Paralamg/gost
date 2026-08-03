"""
Генерирует example.docx для визуальной проверки всех стилей WordBuilder.
Запуск: python examples/example.py
"""
import logging
from pathlib import Path

from gost import BreakType, ElementFactory, IndexManager, IndexType, Pt, WordBuilder

OUTPUT = Path(__file__).parent / "example.docx"
ASSETS = Path(__file__).parent


def main() -> None:
    wb = WordBuilder()
    factory = ElementFactory(IndexManager(character="А", index_type=IndexType.CONTINUOUS))

    # --- Заголовки ---
    wb.add_element(factory.create_head("Основная часть (уровень 0, Heading 1, по центру, UPPER, без номера)", False, 0))
    wb.add_element(factory.create_head("Заголовок первого уровня (уровень 1, Heading 2, с номером)", True, 1))
    wb.add_element(factory.create_head("Заголовок второго уровня (уровень 2, Heading 3, с номером)", True, 2))
    wb.add_element(factory.create_head("Заголовок третьего уровня (уровень 3, Heading 4, с номером)", True, 3))
    wb.add_element(factory.create_head("Заголовок четвертого уровня (уровень 4, Heading 5, с номером)", True, 4))

    # --- Обычный текст ---
    wb.add_element(factory.create_text(
        "Обычный абзац: Times New Roman 14pt, выравнивание по ширине, "
        "межстрочный интервал 1,5, красная строка 1,25 см. "
        "Длинный текст нужен, чтобы проверить выравнивание по ширине на нескольких строках — "
        "вот ещё немного слов для этого. "
    ))
    # Инлайн-разметка: жирный/курсив/жирный курсив/подчёркнутый и неразрывный пробел (~).
    wb.add_element(factory.create_text(
        "Инлайн-разметка: **жирный**, *курсив*, _тоже курсив_, "
        "***жирный курсив***, __подчёркнутый__. "
        "Неразрывный пробел склеивает «рис.~1» и «10~кг», чтобы их не разорвал перенос строки."
    ))

    # Пример ссылки на изображение. Подпись (alt) тоже понимает инлайн-разметку.
    image = factory.create_image(
        str(ASSETS / "test_image.png"),
        "Пример подписи к рисунку с *курсивом*. Пример подписи к рисунку. "
        "Пример подписи к рисунку. Пример подписи к рисунку. Пример подписи к рисунку",
    )
    wb.add_element(factory.create_text(f"Тут демонстрируется пример ссылки на рисунок {image.index}. "
                                       f"Также можно ссылаться на таблицы и формулы."))

    # А после ссылки добавляется рисунок
    wb.add_element(image)

    # --- Таблица по умолчанию: «№ п/п» и строка номеров столбцов ---
    # Числа форматируются до передачи — Table приводит значения через str().

    base_table = factory.create_table(
        {
            "Показатель А": [f"{v:.2f}" for v in (1.0, 2.5, 3.333)],
            "Показатель Б": [f"{v:.2f}" for v in (100.0, 200.0, 300.0)],
            "Текст": ["alpha", "beta", "gamma"],
        },
        title="Таблица по умолчанию: со столбцом «№ п/п» и строкой номеров столбцов",
    )

    wb.add_element(factory.create_text(f"На таблице {base_table.index} можно увидеть пример таблицы по умолчанию: "
                                       f"автоматически добавляется нумерация столбцов и строк, "
                                       f"после нумерации строк идет двойная линия."))

    # --- Таблица без автонумерации; подпись с инлайн-разметкой ---
    wb.add_element(factory.create_table(
        {"Имя": ["Алиса", "Боб", "Вера"], "Возраст": [25, 30, 22]},
        title="Подпись с **жирным** и ***жирным курсивом*** (без «№ п/п» и номеров столбцов)",
        show_row_numbers=False,
        show_column_numbers=False,
    ))

    # --- Таблица без строки с именами столбцов ---
    wb.add_element(factory.create_table(
        {"Col1": [10, 20], "Col2": [30, 40]},
        title="Без строки-заголовка (show_header=False)",
        show_header=False,
    ))

    # --- Две таблицы подряд: подписи не дают Word слить их в одну ---
    wb.add_element(factory.create_table(
        {"X": [0.1, 0.2], "Y": [0.3, 0.4]},
        title="Таблица сразу за другой таблицей",
    ))

    # --- Разрыв на новую страницу: «Продолжение таблицы N» ---
    # Строка 20 — последняя, которая помещается на страницу при этой вёрстке.
    # Значение зависит от всего, что выше по документу, поэтому подбирать его
    # руками не нужно: см. autosplit_example.py, где split_after="auto"
    # находит его измерением.
    wb.add_element(factory.create_table(
        _cost_data(50),
        title="Себестоимость с переносом на следующую страницу",
        split_after=[15],
    ))

    # --- Формулы ---
    wb.add_element(factory.create_page_break())
    wb.add_element(factory.create_head("Формулы", False, 0))

    # Нумерованная формула с расшифровкой обозначений: формула по центру,
    # номер у правого поля, ниже блок «где ...».
    kinetic = factory.create_formula(
        r"E = \frac{mv^2}{2}",
        where={"m": "масса тела, кг", "v": "скорость тела, м/с"},
    )
    wb.add_element(factory.create_text(
        f"Кинетическая энергия вычисляется по формуле ({kinetic.index})."
    ))
    wb.add_element(kinetic)

    # Формула без ссылок в тексте — ГОСТ разрешает не нумеровать.
    wb.add_element(factory.create_text("Ненумерованная формула идёт просто по центру:"))
    wb.add_element(factory.create_formula(r"a^2 + b^2 = c^2", numbered=False))

    # Тяжёлая формула: вложенные радикал, дробь и сумма с пределами — проверка
    # того, что скобки и знак корня растягиваются по высоте содержимого.
    wb.add_element(factory.create_text(
        "Среднеквадратическое отклонение — вложенные радикал, дробь и сумма:"
    ))
    wb.add_element(factory.create_formula(
        r"\sigma = \sqrt{\frac{\sum_{i=1}^{n}(x_i - \bar{x})^2}{n - 1}}",
        # Ключи блока «где» — обычный текст с инлайн-разметкой, не LaTeX.
        where={"*x*": "значение отдельного измерения", "n": "объём выборки"},
    ))

    # Матрица, система и кириллица в индексах.
    wb.add_element(factory.create_text("Матрица, система уравнений и индексы кириллицей:"))
    wb.add_element(factory.create_formula(
        r"A = \begin{pmatrix} a_{11} & a_{12} \\ a_{21} & a_{22} \end{pmatrix}"
    ))
    wb.add_element(factory.create_formula(
        r"f(x) = \begin{cases} x, & x > 0 \\ 0, & x \leq 0 \end{cases}"
    ))
    wb.add_element(factory.create_formula(
        r"K_{\text{сум}} = \frac{P_{вх}}{P_{вых}} \cdot 100\%"
    ))

    # --- Разрыв: по умолчанию «Следующая страница» ---
    # Разрыв раздела: новый раздел начинается с новой страницы и наследует
    # поля предыдущего.
    wb.add_element(factory.create_page_break())
    wb.add_element(factory.create_text("Этот абзац открывает новый раздел на новой странице."))

    # --- Разрыв: обычная «Страница» ---
    # Разрыв внутри раздела — новой страницы достаточно, отдельный раздел не нужен.
    wb.add_element(factory.create_page_break(BreakType.PAGE))
    wb.add_element(factory.create_text("Этот абзац начинается с новой страницы того же раздела."))

    # --- Настройка стилей: локально и глобально ---
    wb.add_element(factory.create_head("Настройка стилей", False, 0))

    # Локально: правим свойства стиля конкретной таблицы — на другие не влияет.
    local_table = factory.create_table(
        {"Параметр": ["Шрифт ячеек", "Подпись"], "Значение": ["16 pt", "жирная"]},
        title="Локальный стиль этой таблицы",
    )
    local_table.text_style.font_size = Pt(16)   # только текст ячеек этой таблицы
    local_table.caption_style.bold = True        # только подпись этой таблицы
    wb.add_element(local_table)

    # Глобально: меняем дефолт фабрики — влияет на все последующие элементы.
    factory.style.normal.font_size = Pt(10)
    factory.style.table_text.font_name = "Courier New"
    wb.add_element(factory.create_text(
        "После правки factory.style последующий обычный текст идёт шрифтом 10 pt."
    ))
    wb.add_element(factory.create_table(
        {"Код": ["A1", "B2"], "Описание": ["первый", "второй"]},
        title="Глобальный стиль: моноширинный шрифт в ячейках",
    ))

    wb.save(OUTPUT)  # об успешном сохранении сообщит сам WordBuilder, на уровне INFO


def _cost_data(rows: int) -> dict[str, list[str]]:
    positions = ["Бетон", "ПГС", "Щебень", "ЖБИ", "Цемент", "Песок", "ПАВ"]
    return {
        "Позиция": [positions[i % len(positions)] for i in range(rows)],
        "Выручка от реализации, тыс.руб.": [f"{8741.0 - i * 100:.1f}" for i in range(rows)],
        "Материальные затраты, план": [f"{322.2 - i * 10:.2f}" for i in range(rows)],
        "Материальные затраты, факт": [f"{151.8 + i * 10:.2f}" for i in range(rows)],
        "Расходы на оплату труда, план": [f"{221.82 - i * 5:.2f}" for i in range(rows)],
        "Расходы на оплату труда, факт": [f"{223.6 + i * 5:.2f}" for i in range(rows)],
    }


if __name__ == "__main__":
    # DEBUG вместо INFO покажет время вывода каждой таблицы и рисунка.
    logging.basicConfig(level=logging.INFO, format="%(levelname)-7s %(name)s: %(message)s")
    main()
