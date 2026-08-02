from collections.abc import Mapping, Sequence
from typing import Any

from gost.elements.head import Head
from gost.elements.image import Image
from gost.elements.page_break import BreakType, PageBreak
from gost.elements.table.table import SplitAfter, Table
from gost.elements.text import Text
from gost.index.index_manager import IndexManager
from gost.styles import StyleSheet


class ElementFactory:
    def __init__(
            self,
            index_manager: IndexManager = IndexManager(),
            style: StyleSheet | None = None,
    ) -> None:
        assert index_manager is not None
        self.__index_manager = index_manager
        # Глобальный уровень: значения по умолчанию для всех создаваемых элементов.
        # Каждому элементу отдаётся своя копия, поэтому локальные правки одного
        # элемента не текут в лист и в другие элементы.
        self.style = style if style is not None else StyleSheet.gost()

    def create_text(self, text: str) -> Text:
        return Text(text, self.style.normal.copy())

    def create_head(self, use_numbers: bool, text: str, level: int) -> Head:
        index = self.__index_manager.get_head_index(level)
        return Head(index, use_numbers, text, level, self.style.headings[level].copy())

    def create_page_break(self, break_type: BreakType = BreakType.NEXT_PAGE) -> PageBreak:
        return PageBreak(break_type)

    def create_image(self, path: str, alt: str) -> Image:
        index = self.__index_manager.get_image_index()
        return Image(index, path, alt, self.style.image_caption.copy())

    def create_table(
            self,
            data: Mapping[str, Sequence[Any]],
            title: str,
            *,
            show_header: bool = True,
            show_row_numbers: bool = True,
            show_column_numbers: bool = True,
            split_after: SplitAfter = None,
    ) -> Table:
        index = self.__index_manager.get_table_index()
        return Table(
            index,
            data,
            title,
            caption_style=self.style.table_caption.copy(),
            text_style=self.style.table_text.copy(),
            show_header=show_header,
            show_row_numbers=show_row_numbers,
            show_column_numbers=show_column_numbers,
            split_after=split_after,
        )
