"""Сборка документа из элементов и запись его в файл."""

import logging
import tempfile
from contextlib import ExitStack
from pathlib import Path

from docx import Document as new_document
from docx.document import Document

from .elements.element import ElementBase
from .elements import Table
from .layout.autosplit import render_with_auto_splits
from .layout.measurer import LayoutMirror, WordMirror
from .progress import tracked
from .styles import apply_gost_styles
from .timing import logged_duration

logger = logging.getLogger(__name__)


class WordBuilder:
    """Накопитель элементов документа: собирает их и сохраняет в .docx.

    Элементы копятся в порядке добавления и выводятся в документ только при
    :meth:`save` — до этого момента их можно править, например менять стили.
    """

    def __init__(self) -> None:
        self.__elements: list[ElementBase] = []

    def add_element(self, element: ElementBase) -> None:
        """Добавляет элемент в конец документа."""
        self.__elements.append(element)

    def add_elements(self, elements: list[ElementBase]) -> None:
        """Добавляет элементы в конец документа, сохраняя их порядок."""
        self.__elements.extend(elements)

    def save(self, path: Path, mirror: LayoutMirror | None = None) -> None:
        """Собирает документ и сохраняет его в файл формата .docx.

        Args:
            path: Путь к итоговому файлу.
            mirror: Чем измерять вёрстку для таблиц со split_after="auto".
                По умолчанию — реальный Word через COM. Без таких таблиц не
                используется.
        """
        if any(isinstance(e, Table) and e.auto_split for e in self.__elements):
            document = self.__build_with_auto_splits(mirror)
        else:
            document = self.__build()

        with logged_duration(logger, "Документ записан на диск"):
            document.save(str(path))
        logger.info("Документ сохранён: %s (элементов: %d, таблиц: %d)",
                    path, len(self.__elements), len(document.tables))

    def __build_with_auto_splits(self, mirror: LayoutMirror | None) -> Document:
        """Собирает документ, подбирая разрывы таблицам со split_after="auto".

        Зеркало и рабочий каталог живут только на время сборки: после неё
        временные файлы удаляются, а Word — если поднимали его мы — закрывается.
        """
        with ExitStack() as stack:
            workdir = Path(stack.enter_context(tempfile.TemporaryDirectory()))
            if mirror is None:
                # Зеркало начинается с пустого документа с теми же стилями и
                # дальше повторяет собираемый кусок за куском.
                base = workdir / "mirror.docx"
                _new_document().save(str(base))
                mirror = stack.enter_context(WordMirror(base))

            with logged_duration(logger, "Документ собран с автоподбором: элементов %d",
                                 len(self.__elements)):
                document = _new_document()
                render_with_auto_splits(
                    self.__elements, document, mirror, _new_document, workdir,
                )
            return document

    def __build(self) -> Document:
        """Собирает документ обычным проходом: элемент за элементом, сверху вниз."""
        with logged_duration(logger, "Документ собран: элементов %d", len(self.__elements)):
            document = _new_document()
            for element in tracked(self.__elements, "Сборка документа"):
                element.render(document)
        return document


def _new_document() -> Document:
    """Пустой документ с настроенными стилями ГОСТ — основа для любой сборки."""
    document = new_document()
    apply_gost_styles(document)
    return document
