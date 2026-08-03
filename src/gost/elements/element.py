"""Базовые типы элементов документа."""

from abc import ABC, abstractmethod

from docx.document import Document


class ElementBase(ABC):
    """Элемент документа — всё, что умеет вывести себя в документ Word."""

    @abstractmethod
    def render(self, document: Document) -> None:
        """Дописывает элемент в конец *document*.

        Вызывается при сборке документа и не обязан быть однократным: при
        автоподборе разрывов таблиц один и тот же элемент выводится ещё и в
        черновые документы. Поэтому render не меняет состояние элемента и не
        расходует ничего одноразового.
        """
        pass


class NumberedElement(ElementBase, ABC):
    """Элемент с номером: таблица, рисунок, формула, заголовок.

    Номер выдаёт :class:`~gost.index.index_manager.IndexManager` при создании
    элемента, поэтому на него можно сослаться в тексте до вставки самого
    элемента. Пустая строка означает, что элемент не нумеруется.
    """

    def __init__(self, index: str) -> None:
        self.index = index
