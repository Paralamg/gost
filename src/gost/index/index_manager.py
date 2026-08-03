"""Выдача номеров элементам документа."""

from gost.index.counter import Counter, HeadCounter
from gost.index.index_type import IndexType


class IndexManager:
    """Счётчики номеров разделов, рисунков, таблиц и формул.

    У каждого вида элементов свой счётчик, поэтому таблицы и рисунки нумеруются
    независимо. Номер выдаётся по запросу и сразу расходуется — порядок вызовов
    задаёт порядок номеров.

    Args:
        character: Префикс приложения: «А» даёт «Таблица А.1», «Рисунок А.1».
            ``None`` — без префикса.
        index_type: Режим нумерации объектов: сквозной по документу или
            относительно номера текущего раздела.
    """

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
        self.__formula_counter = Counter()

    def get_head_index(self, level: int) -> str:
        """Занимает следующий номер раздела на *level* и возвращает его.

        Уровень 0 не нумеруется — для него возвращается пустая строка. Номер
        занимается даже у заголовка, который выводится без номера: иначе
        нумерация следующих разделов сбилась бы.
        """
        index = self.__head_counter.get_next(level)

        if self.__character and index:
            index = self.__add_character(index)

        # Новый раздел начинает нумерацию рисунков, таблиц и формул заново: 1.1, 1.2,
        # затем 2.1, 2.2. Нумерация самих разделов при этом сквозная.
        if level == 1 and self.__index_type == IndexType.CHAPTER_RELATIVE:
            self.__reset_counters()
        return index

    def get_image_index(self):
        """Занимает следующий номер рисунка и возвращает его."""
        return self.__format_index(self.__image_counter.get_next())

    def get_table_index(self):
        """Занимает следующий номер таблицы и возвращает его."""
        return self.__format_index(self.__table_counter.get_next())

    def get_formula_index(self):
        """Занимает следующий номер формулы и возвращает его."""
        return self.__format_index(self.__formula_counter.get_next())

    def __format_index(self, raw_index):
        """Оформляет номер объекта: приписывает раздел и префикс приложения."""
        index = str(raw_index)
        if self.__index_type == IndexType.CHAPTER_RELATIVE:
            index = self.__add_chapter(index)

        if self.__character:
            index = self.__add_character(index)

        return index

    def __add_chapter(self, index: str) -> str:
        """Приписывает номер текущего раздела: «1» → «2.1»."""
        current_chapter = self.__head_counter.get_current_chapter()
        index = f"{current_chapter}.{index}"
        return index

    def __add_character(self, index: str) -> str:
        """Приписывает префикс приложения: «1» → «А.1»."""
        return f"{self.__character}.{index}"

    def __reset_counters(self):
        """Начинает нумерацию объектов заново — с начала нового раздела."""
        self.__image_counter.reset()
        self.__table_counter.reset()
        self.__formula_counter.reset()