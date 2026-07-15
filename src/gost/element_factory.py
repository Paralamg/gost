from typing import Optional

import pandas as pd

from gost.elements.head import Head
from gost.elements.image import Image
from gost.elements.table import Table
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
            df: pd.DataFrame,
            float_format: Optional[str] = ".1f",
            title: Optional[str] = None,
            show_header: bool = True,
            show_index: bool = True,
            bold_header: bool = False,
            bold_index: bool = False,
    ) -> Table:
        index = self.__index_manager.get_table_index()
        return Table(index, df, float_format, title, show_header,show_index, bold_header, bold_index)