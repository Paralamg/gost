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
    factory = ElementFactory(IndexManager(index_type=IndexType.CONTINUOUS))

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
        "вот ещё немного слов для этого."
    ))
    wb.add_element(factory.create_text(
        "Абзац с **жирным текстом** внутри обычного предложения — "
        "проверка парсера **bold**-разметки."
    ))
    wb.add_element(factory.create_text("Ещё один абзац — проверка отступа красной строки."))

    # --- Рисунок ---
    wb.add_element(factory.create_image(
        str(ASSETS / "test_image.png"),
        "Пример подписи к рисунку. Пример подписи к рисунку. Пример подписи к рисунку. "
        "Пример подписи к рисунку. Пример подписи к рисунку",
    ))

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

    wb.save(OUTPUT)
    print(f"Сохранено: {OUTPUT}")


if __name__ == "__main__":
    main()
