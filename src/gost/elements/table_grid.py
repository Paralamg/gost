"""Построение сетки таблицы. Не зависит от python-docx."""

from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from typing import Any

ROW_NUMBER_HEADER = "№ п/п"


@dataclass(frozen=True)
class Grid:
    """Ячейки таблицы, приведённые к строкам и разложенные по ролям."""

    header: list[str] | None
    numbers: list[str] | None
    body: list[list[str]]
    has_row_numbers: bool

    @property
    def width(self) -> int:
        for row in (self.header, self.numbers, *self.body):
            if row is not None:
                return len(row)
        return 0

    @property
    def head_rows(self) -> list[list[str]]:
        """Шапка таблицы: имена столбцов и/или их номера."""
        return [row for row in (self.header, self.numbers) if row is not None]


def build_grid(
        data: Mapping[str, Sequence[Any]],
        *,
        show_header: bool = True,
        show_row_numbers: bool = True,
        show_column_numbers: bool = True,
) -> Grid:
    """Раскладывает данные по столбцам в сетку строк.

    Args:
        data: Данные по столбцам: {«Показатель А»: [1.0, 2.5], ...}.
        show_header: Выводить строку с именами столбцов.
        show_row_numbers: Добавлять слева столбец «№ п/п» с нумерацией строк.
        show_column_numbers: Добавлять строку с номерами столбцов 1..N.

    Raises:
        ValueError: Если столбцов нет или они разной длины.
    """
    columns = list(data)
    if not columns:
        raise ValueError("Таблица должна содержать хотя бы один столбец")

    lengths = {len(values) for values in data.values()}
    if len(lengths) > 1:
        sizes = {name: len(values) for name, values in data.items()}
        raise ValueError(f"Столбцы разной длины: {sizes}")
    (row_count,) = lengths

    header = list(columns) if show_header else None
    body = [[_to_text(data[column][i]) for column in columns] for i in range(row_count)]

    if show_row_numbers:
        if header is not None:
            header.insert(0, ROW_NUMBER_HEADER)
        for number, row in enumerate(body, start=1):
            row.insert(0, str(number))

    width = len(columns) + (1 if show_row_numbers else 0)
    numbers = [str(i) for i in range(1, width + 1)] if show_column_numbers else None

    return Grid(header, numbers, body, has_row_numbers=show_row_numbers)


def _to_text(value: Any) -> str:
    return "" if value is None else str(value)
