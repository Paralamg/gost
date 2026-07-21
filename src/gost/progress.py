"""Шкала прогресса сборки документа."""

from collections.abc import Iterator, Sequence
from typing import TypeVar

from tqdm import tqdm

T = TypeVar("T")


def tracked(elements: Sequence[T], description: str) -> Iterator[T]:
    """Перебирает элементы, показывая шкалу прогресса.

    Шкала гаснет сама, когда вывод идёт не в терминал: в файле лога или в CI
    она осталась бы мусором из сотен строк с процентами. Это делает
    «disable=None» — у tqdm это значит «выключить, если stderr не терминал»,
    в отличие от False, которое включает шкалу всегда.

    Args:
        elements: Что перебирать. Длина нужна, чтобы шкала знала свой предел.
        description: Подпись слева от шкалы.
    """
    return tqdm(elements, desc=description, unit="эл.", disable=None)
