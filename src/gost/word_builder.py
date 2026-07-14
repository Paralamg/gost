from pathlib import Path

from docx import Document

from .elements.element import ElementBase
from .styles import apply_gost_styles


class WordBuilder:
    def __init__(self) -> None:
        """
        Инициализирует документ с полями по ГОСТ 7.32-2001 и применяет стили.
        """
        self.__document = Document()
        self.__is_rendered = False
        self.__figure_counter = 0
        self.__table_counter = 0
        self.__elements: list[ElementBase] = []

        apply_gost_styles(self.__document)

    def add_element(self, element: ElementBase) -> None:
        self.__elements.append(element)

    def save(self, path: Path) -> None:
        """Сохраняет документ в файл формата .docx.

        Args:
            path: Путь к итоговому файлу.
        """
        if not self.__is_rendered:
            self.__render_elements()
            self.__is_rendered = True

        self.__document.save(str(path))

    def __render_elements(self):
        for element in self.__elements:
            element.render(self.__document)
