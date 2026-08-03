"""Создание элементов документа: нумерация и оформление в одном месте."""

from collections.abc import Mapping, Sequence
from typing import Any

from gost.elements.formula import Formula
from gost.elements.head import Head
from gost.elements.image import Image
from gost.elements.page_break import BreakType, PageBreak
from gost.elements.table.table import SplitAfter, Table
from gost.elements.text import Text
from gost.index.index_manager import IndexManager
from gost.styles import StyleSheet


class ElementFactory:
    """Создаёт элементы документа, проставляя им номера и оформление.

    Единственный способ получить элемент: конструкторы самих элементов требуют
    уже готовые номер и стили, а знает о них фабрика. Номер выдаётся в момент
    вызова ``create_*``, поэтому элементы нужно создавать в том же порядке, в
    котором они пойдут в документ, — на него же можно сослаться в тексте до
    вставки самого элемента::

        table = factory.create_table(data, title="Смета")
        wb.add_element(factory.create_text(f"Данные приведены в таблице {table.index}."))
        wb.add_element(table)

    Оформление берётся из :attr:`style`; каждому элементу отдаётся копия,
    поэтому правки в одном элементе не задевают остальные, а правки в
    :attr:`style` действуют на элементы, созданные после них.

    Args:
        index_manager: Счётчики номеров: режим нумерации и префикс приложения.
        style: Оформление по умолчанию. ``None`` — :meth:`StyleSheet.gost`.
    """

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
        """Создаёт абзац обычного текста.

        Args:
            text: Текст абзаца. Понимает инлайн-разметку: ``**жирный**``,
                ``*курсив*``, ``__подчёркнутый__``, ``~`` — неразрывный пробел.
        """
        return Text(text, self.style.normal.copy())

    def create_head(self, text: str, use_numbers: bool, level: int) -> Head:
        """Создаёт заголовок и занимает под него номер раздела.

        Номер занимается в любом случае, даже при ``use_numbers=False``: разделы
        нумеруются подряд, и пропуск номера сбил бы нумерацию следующих. Не
        нумеруется только уровень 0 — у структурных элементов номера нет.

        Args:
            text: Текст заголовка.
            use_numbers: Выводить ли номер раздела перед текстом.
            level: Логический уровень 0..4. 0 — структурный элемент
                («Введение», «Заключение»): по центру, прописными, без номера;
                1..4 — разделы и подразделы («1», «1.1», «1.1.1», «1.1.1.1»).
        """
        index = self.__index_manager.get_head_index(level)
        return Head(index, text, use_numbers, level, self.style.headings[level].copy())

    def create_page_break(self, break_type: BreakType = BreakType.NEXT_PAGE) -> PageBreak:
        """Создаёт разрыв страницы или раздела.

        Args:
            break_type: Вид разрыва. По умолчанию «Следующая страница» —
                разрыв раздела, начинающий новый раздел с новой страницы.
        """
        return PageBreak(break_type)

    def create_image(self, path: str, alt: str) -> Image:
        """Создаёт рисунок с подписью «Рисунок N – alt».

        Файл читается не сейчас, а при сборке документа: если к тому моменту его
        не окажется на месте, вместо рисунка встанет заглушка с предупреждением
        в лог.

        Args:
            path: Путь к файлу изображения.
            alt: Подпись под рисунком. Понимает инлайн-разметку.
        """
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
        """Создаёт таблицу с подписью «Таблица N – title».

        Параметры совпадают с :class:`~gost.elements.table.table.Table` —
        подробное описание данных, ограничений и точек разрыва там.

        Args:
            data: Данные по столбцам: ``{«Показатель А»: [1.0, 2.5], ...}``.
                Совместимо с ``df.to_dict("list")``. Значения приводятся к
                строке через ``str()``.
            title: Наименование таблицы. Обязательно, понимает инлайн-разметку.
            show_header: Выводить строку с именами столбцов.
            show_row_numbers: Добавлять слева столбец «№ п/п».
            show_column_numbers: Добавлять строку с номерами столбцов 1..N.
            split_after: Номера строк тела, после которых таблицу переносить на
                новую страницу, либо ``"auto"`` — подобрать их измерением.

        Raises:
            ValueError: Если наименование пусто, столбцов нет, они разной длины
                или split_after выходит за пределы тела таблицы.
        """
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

    def create_formula(
            self,
            latex: str,
            *,
            where: Mapping[str, str] | None = None,
            numbered: bool = True,
    ) -> Formula:
        """Создаёт формулу из LaTeX — объектом «Уравнение» Word.

        Args:
            latex: Формула на LaTeX. Пишите её как ``r"..."``, иначе обратные
                слеши придётся удваивать.
            where: Расшифровка обозначений в порядке их появления в формуле —
                блок «где ...». Ключи и значения понимают инлайн-разметку.
            numbered: Нумеровать ли формулу. ``False`` не тратит номер: ГОСТ
                разрешает не нумеровать формулы, на которые нет ссылок.

        Raises:
            ValueError: LaTeX не разобрался. Ошибка всплывает здесь, а не при
                сохранении документа.
        """
        # Ненумерованная формула не тратит счётчик: ГОСТ разрешает не нумеровать
        # формулы, на которые нет ссылок в тексте.
        index = self.__index_manager.get_formula_index() if numbered else ""
        return Formula(
            index,
            latex,
            where=where,
            formula_style=self.style.formula.copy(),
            note_style=self.style.formula_note.copy(),
        )
