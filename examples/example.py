"""
Генерирует style_preview.docx для визуальной проверки всех стилей WordBuilder.
Запуск: python examples/example.py
"""
from pathlib import Path

import pandas as pd

from gost import WordBuilder
from gost.element_factory import ElementFactory
from gost.index.index_manager import IndexManager
from gost.index.index_type import IndexType

OUTPUT = Path(__file__).parent / "example.docx"
ASSETS = Path(__file__).parent  # изображений нет, путь нужен только как base_dir


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

    # --- Рисунок (файл не существует → показывает placeholder) ---
    wb.add_element(factory.create_image(
        str(ASSETS / "test_image.png"),
        "Пример подписи к рисунку. Пример подписи к рисунку. Пример подписи к рисунку. "
        "Пример подписи к рисунку. Пример подписи к рисунку",
    ))

    # --- Таблица с заголовком ---
    df_full = pd.DataFrame(
        {
            "Показатель А": [1.0, 2.5, 3.333],
            "Показатель Б": [100.0, 200.0, 300.0],
            "Текст": ["alpha", "beta", "gamma"],
        },
        index=["Строка 1", "Строка 2", "Строка 3"],
    )
    wb.add_element(factory.create_table(
        df_full,
        float_format=".2f",
        title="Таблица с заголовком и индексом. Таблица с заголовком и индексом. "
              "Таблица с заголовком и индексом. Таблица с заголовком и индексом",
    ))

    # --- Таблица без заголовка ---
    df_no_title = pd.DataFrame(
        {"X": [0.1, 0.2], "Y": [0.3, 0.4]},
        index=["a", "b"],
    )
    wb.add_element(factory.create_table(df_no_title, title=None, bold_header=True, bold_index=True))

    # --- Таблица без строки заголовков столбцов ---
    df_no_header = pd.DataFrame(
        {"Col1": [10, 20], "Col2": [30, 40]},
        index=["i", "ii"],
    )
    wb.add_element(factory.create_table(
        df_no_header,
        title="Таблица без строки-заголовка (show_header=False)",
        show_header=False,
    ))

    # --- Таблица без столбца индекса ---
    df_no_index = pd.DataFrame(
        {"Имя": ["Алиса", "Боб", "Вера"], "Возраст": [25, 30, 22]},
    )
    wb.add_element(factory.create_table(
        df_no_index,
        title="Таблица без индекса (show_index=False)",
        show_index=False,
    ))

    wb.save(OUTPUT)
    print(f"Сохранено: {OUTPUT}")


if __name__ == "__main__":
    main()
