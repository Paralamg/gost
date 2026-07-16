import logging
import tempfile
from contextlib import ExitStack
from pathlib import Path

from docx import Document as new_document
from docx.document import Document

from .elements.element import ElementBase
from .elements.table import Table
from .layout.autosplit import resolve_auto_splits
from .layout.measurer import PageMeasurer, WordMeasurer
from .styles import apply_gost_styles
from .timing import logged_duration

logger = logging.getLogger(__name__)


class WordBuilder:
    def __init__(self) -> None:
        self.__elements: list[ElementBase] = []

    def add_element(self, element: ElementBase) -> None:
        self.__elements.append(element)

    def add_elements(self, elements: list[ElementBase]) -> None:
        self.__elements.extend(elements)

    def save(self, path: Path, measurer: PageMeasurer | None = None) -> None:
        """Собирает документ и сохраняет его в файл формата .docx.

        Args:
            path: Путь к итоговому файлу.
            measurer: Чем измерять вёрстку для таблиц со split_after="auto".
                По умолчанию — реальный Word через COM. Без таких таблиц не
                используется.
        """
        if any(isinstance(e, Table) and e.auto_split for e in self.__elements):
            document = self.__build_with_auto_splits(measurer)
        else:
            document, _ = self.__build()

        with logged_duration(logger, "Документ записан на диск"):
            document.save(str(path))
        logger.info("Документ сохранён: %s (элементов: %d, таблиц: %d)",
                    path, len(self.__elements), len(document.tables))

    def __build_with_auto_splits(self, measurer: PageMeasurer | None) -> Document:
        with ExitStack() as stack:
            workdir = stack.enter_context(tempfile.TemporaryDirectory())
            if measurer is None:
                measurer = stack.enter_context(WordMeasurer())
            return resolve_auto_splits(
                self.__elements,
                self.__build,
                measurer,
                Path(workdir) / "probe.docx",
            )

    def __build(self) -> tuple[Document, list[range]]:
        """Собирает документ с нуля.

        Returns:
            Документ и — для каждого элемента — диапазон индексов таблиц,
            которые он сгенерировал. Диапазоны нужны, чтобы сопоставить
            измеренную вёрстку с элементом.
        """
        # Пересобирается целиком на каждом проходе автоподбора, поэтому на больших
        # документах это заметная часть времени.
        with logged_duration(logger, "Документ собран: элементов %d", len(self.__elements)):
            document = new_document()
            apply_gost_styles(document)

            spans: list[range] = []
            for element in self.__elements:
                before = len(document.tables)
                element.render(document)
                spans.append(range(before, len(document.tables)))
        return document, spans
