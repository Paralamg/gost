from gost.index.head_counter import HeadCounter
from gost.index.counter import Counter
from gost.elements.head import Head
from gost.elements.image import Image
from gost.elements.text import Text


class ElementFactory:
    def __init__(self, character: str) -> None:
        self.__image_counter = Counter()
        self.__table_counter = Counter()
        self.__head_counter = HeadCounter()
        self.__character = character

    def create_text(self, text: str) -> Text:
        return Text(text)

    def create_head(self, use_numbers: bool, text: str, level: int = 1) -> Head:
        number = self.__head_counter.get_next(level)
        return Head(number, use_numbers, text, level)

    def create_image(self, path: str, alt: str) -> Image:
        number = self.__image_counter.get_next()
        return Image(number, alt, path)