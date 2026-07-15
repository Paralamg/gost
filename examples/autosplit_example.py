"""
Автоподбор точек разрыва: split_after="auto".

Точку разрыва нельзя вычислить заранее — высота строки зависит от переноса
слов, а он от метрик шрифта и ширины ячейки; плюс таблица может начинаться
с любого места страницы. Поэтому вёрстка измеряется: WordBuilder.save()
рендерит документ, спрашивает у Word, на какой странице оказалась каждая
строка, и ставит разрыв ровно туда. Разрывы подбираются по одному, потому
что каждый сдвигает вниз всё, что за ним следует.

Требует Windows, установленный Word и pywin32:
    pip install gost-docx[autosplit]

Запуск: python examples/autosplit_example.py
"""
from pathlib import Path

from gost import WordBuilder
from gost.element_factory import ElementFactory

OUTPUT = Path(__file__).parent / "autosplit_example.docx"
ROWS = 110


def main() -> None:
    wb = WordBuilder()
    factory = ElementFactory()

    # Текст перед таблицей: таблица начинается не с верха страницы, поэтому
    # число строк на первой странице отличается от последующих.
    wb.add_element(factory.create_text(
        "Таблица ниже начинается не с верха страницы, и разрывы всё равно "
        "попадают точно в край страницы. " * 5
    ))

    table = factory.create_table(
        {
            "Позиция": [f"Материал {i}" for i in range(1, ROWS + 1)],
            "Выручка, тыс.руб.": [f"{1000 + i * 7:.1f}" for i in range(ROWS)],
            "Затраты, план": [f"{300 - i:.2f}" for i in range(ROWS)],
            "Затраты, факт": [f"{150 + i:.2f}" for i in range(ROWS)],
        },
        title="Автоподбор точек разрыва",
        split_after="auto",
    )
    wb.add_element(table)

    wb.save(OUTPUT)  # здесь и происходит подбор — это медленно, Word открывается
    print(f"Сохранено: {OUTPUT}")
    print(f"Подобранные точки разрыва: {table.split_after} (строк в теле: {ROWS})")


if __name__ == "__main__":
    main()
