from collections.abc import Mapping, Sequence
from typing import Any

from gost.elements.head import Head
from gost.elements.image import Image
from gost.elements.table import SplitAfter, Table
from gost.elements.text import Text
from gost.index.index_manager import IndexManager


class ElementFactory:
    def __init__(self, index_manager: IndexManager = IndexManager()) -> None:
        assert index_manager is not None
        self.__index_manager = index_manager

    def create_text(self, text: str) -> Text:
        return Text(text)

    def create_head(self, use_numbers: bool, text: str, level: int) -> Head:
        index = self.__index_manager.get_head_index(level)
        return Head(index, use_numbers, text, level)

    def create_image(self, path: str, alt: str) -> Image:
        index = self.__index_manager.get_image_index()
        return Image(index, path, alt)

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
            show_header=show_header,
            show_row_numbers=show_row_numbers,
            show_column_numbers=show_column_numbers,
            split_after=split_after,
        )
