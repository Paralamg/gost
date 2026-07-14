"""
Генерирует style_preview.docx для визуальной проверки всех стилей WordBuilder.
Запуск: python examples/style_preview.py
"""
from pathlib import Path

import pandas as pd

from gost import WordBuilder

OUTPUT = Path(__file__).parent / "style_preview.docx"
ASSETS = Path(__file__).parent  # изображений нет, путь нужен только как base_dir


def main() -> None:
    wb = WordBuilder()

    # --- Заголовки ---
    wb.add_title("Заголовок первого уровня (Title, по центру)", level=1)
    wb.add_title("Заголовок второго уровня (Heading 1, слева)", level=2)
    wb.add_title("Заголовок третьего уровня (Heading 2, слева)", level=3)

    wb.add_separator()

    # --- Обычный текст ---
    wb.add_text(
        "Обычный абзац: Times New Roman 14pt, выравнивание по ширине, "
        "межстрочный интервал 1,5, красная строка 1,25 см. "
        "Длинный текст нужен, чтобы проверить выравнивание по ширине на нескольких строках — "
        "вот ещё немного слов для этого."
    )
    wb.add_text(
        "Абзац с **жирным текстом** внутри обычного предложения — "
        "проверка парсера **bold**-разметки."
    )
    wb.add_text("Ещё один абзац — проверка отступа красной строки.")

    # --- Рисунок (файл не существует → показывает placeholder) ---
    wb.add_image("Пример подписи к рисунку. Пример подписи к рисунку. Пример подписи к рисунку. Пример подписи к рисунку. Пример подписи к рисунку", "test_image.png")

    # --- Таблица с заголовком ---
    df_full = pd.DataFrame(
        {
            "Показатель А": [1.0, 2.5, 3.333],
            "Показатель Б": [100.0, 200.0, 300.0],
            "Текст": ["alpha", "beta", "gamma"],
        },
        index=["Строка 1", "Строка 2", "Строка 3"],
    )
    wb.add_table(df_full, float_format=".2f", title="Таблица с заголовком и индексом. Таблица с заголовком и индексом. Таблица с заголовком и индексом. Таблица с заголовком и индексом")

    # --- Таблица без заголовка ---
    df_no_title = pd.DataFrame(
        {"X": [0.1, 0.2], "Y": [0.3, 0.4]},
        index=["a", "b"],
    )
    wb.add_table(df_no_title, title=None, bold_header=True, bold_index=True)

    # --- Таблица без строки заголовков столбцов ---
    df_no_header = pd.DataFrame(
        {"Col1": [10, 20], "Col2": [30, 40]},
        index=["i", "ii"],
    )
    wb.add_table(df_no_header, title="Таблица без строки-заголовка (show_header=False)", show_header=False)

    # --- Таблица без столбца индекса ---
    df_no_index = pd.DataFrame(
        {"Имя": ["Алиса", "Боб", "Вера"], "Возраст": [25, 30, 22]},
    )
    wb.add_table(df_no_index, title="Таблица без индекса (show_index=False)", show_index=False)

    wb.save(OUTPUT)
    print(f"Сохранено: {OUTPUT}")


if __name__ == "__main__":
    main()
