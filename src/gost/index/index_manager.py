from gost.index.counter import Counter, HeadCounter
from gost.index.index_type import IndexType


class IndexManager:
    def __init__(
            self,
            character: str | None = None,
            index_type: IndexType = IndexType.CONTINUOUS
    ) -> None:
        self.__character = character
        self.__index_type = index_type

        self.__head_counter = HeadCounter()
        self.__image_counter = Counter()
        self.__table_counter = Counter()

    def get_head_index(self, level: int) -> str:
        index = self.__head_counter.get_next(level)

        if self.__character:
            index = self.__add_character(index)

        if level == 0 and self.__index_type == IndexType.CHAPTER_RELATIVE:
            self.__reset_counters()
        return index

    def get_image_index(self):
        return self.__format_index(self.__image_counter.get_next())

    def get_table_index(self):
        return self.__format_index(self.__table_counter.get_next())

    def __format_index(self, raw_index):
        index = str(raw_index)
        if self.__index_type == IndexType.CHAPTER_RELATIVE:
            index = self.__add_chapter(index)

        if self.__character:
            index = self.__add_character(index)

        return index

    def __add_chapter(self, index: str) -> str:
        current_chapter = self.__head_counter.get_current_chapter()
        index = f"{current_chapter}.{index}"
        return index

    def __add_character(self, index: str) -> str:
        return f"{self.__character}.{index}"

    def __reset_counters(self):
        self.__image_counter.reset()
        self.__head_counter.reset()