from abc import ABC, abstractmethod
from typing import Optional

from docx.document import Document


class ElementBase(ABC):
    @abstractmethod
    def render(self, document: Document) -> None:
        pass


class NumberedElement(ElementBase, ABC):
    def __init__(self, index: str) -> None:
        self.index = index
