"""
Генерирует example.docx для визуальной проверки всех стилей WordBuilder.
Запуск: python examples/example.py
"""
from pathlib import Path

from gost import WordBuilder
from gost.element_factory import ElementFactory
from gost.index.index_manager import IndexManager
from gost.index.index_type import IndexType

OUTPUT = Path(__file__).parent / "example.docx"
ASSETS = Path(__file__).parent


def main() -> None:
    wb = WordBuilder()
    factory = ElementFactory(IndexManager(character="А", index_type=IndexType.CONTINUOUS))

    # --- Заголовки ---
    wb.add_element(factory.create_head(True, "Заголовок первого уровня (Heading 1, по центру, UPPER)", 1))
    wb.add_element(factory.create_head(True, "Заголовок второго уровня (Heading 2, слева, с номером)", 2))
    wb.add_element(factory.create_head(True, "Заголовок третьего уровня (Heading 3, слева, с номером)", 3))
    wb.add_element(factory.create_head(True, "Заголовок четвертого уровня (Heading 4, слева, с номером)", 4))

    # --- Обычный текст ---
    wb.add_element(factory.create_text(
        "Обычный абзац: Times New Roman 14pt, выравнивание по ширине, "
        "межстрочный интервал 1,5, красная строка 1,25 см. "
        "Длинный текст нужен, чтобы проверить выравнивание по ширине на нескольких строках — "
        "вот ещё немного слов для этого. "
    ))
    wb.add_element(factory.create_text(""))
    wb.add_element(factory.create_text(
        "Абзац с **жирным текстом** внутри обычного предложения — "
        "проверка парсера **bold**-разметки."
    ))

    # Пример ссылки на изображение
    image = factory.create_image(
        str(ASSETS / "test_image.png"),
        "Пример подписи к рисунку. Пример подписи к рисунку. Пример подписи к рисунку. "
        "Пример подписи к рисунку. Пример подписи к рисунку",
    )
    wb.add_element(factory.create_text(f"На рисунке {image.index}."))

    # --- Рисунок ---
    wb.add_element(image)

    # --- Таблица по умолчанию: «№ п/п» и строка номеров столбцов ---
    # Числа форматируются до передачи — Table приводит значения через str().
    wb.add_element(factory.create_table(
        {
            "Показатель А": [f"{v:.2f}" for v in (1.0, 2.5, 3.333)],
            "Показатель Б": [f"{v:.2f}" for v in (100.0, 200.0, 300.0)],
            "Текст": ["alpha", "beta", "gamma"],
        },
        title="Таблица по умолчанию: со столбцом «№ п/п» и строкой номеров столбцов",
    ))

    # --- Таблица без автонумерации ---
    wb.add_element(factory.create_table(
        {"Имя": ["Алиса", "Боб", "Вера"], "Возраст": [25, 30, 22]},
        title="Без «№ п/п» и без строки номеров столбцов",
        show_row_numbers=False,
        show_column_numbers=False,
    ))

    # --- Таблица без строки с именами столбцов ---
    wb.add_element(factory.create_table(
        {"Col1": [10, 20], "Col2": [30, 40]},
        title="Без строки-заголовка (show_header=False)",
        show_header=False,
    ))

    # --- Таблица без подписи ---
    wb.add_element(factory.create_table({"X": [0.1, 0.2], "Y": [0.3, 0.4]}))

    # --- Разрыв на новую страницу: «Продолжение табл. N» ---
    # Строка 20 — последняя, которая помещается на страницу при этой вёрстке.
    # Значение зависит от всего, что выше по документу, поэтому подбирать его
    # руками не нужно: см. autosplit_example.py, где split_after="auto"
    # находит его измерением.
    wb.add_element(factory.create_table(
        _cost_data(60),
        title="Себестоимость с переносом на следующую страницу",
        split_after=[20],
    ))

    wb.save(OUTPUT)
    print(f"Сохранено: {OUTPUT}")


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
    main()
