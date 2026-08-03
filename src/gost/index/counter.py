"""Счётчики номеров: сквозной и многоуровневый."""


class Counter:
    """Счётчик подряд идущих номеров.

    Args:
        initial: Номер, который вернёт первый :meth:`get_next`.
    """

    def __init__(self, initial: int = 1) -> None:
        # Хранится значение «перед первым»: get_next сначала увеличивает, а
        # потом отдаёт, поэтому и в начале, и после reset состояние одинаковое.
        self.__initial = initial - 1
        self.current: int = self.__initial

    def get_next(self) -> int:
        """Занимает следующий номер и возвращает его."""
        self.current = self.current + 1
        return self.current

    def reset(self):
        """Возвращает счётчик в начало: следующий номер снова первый."""
        self.current = self.__initial


class HeadCounter:
    """Многоуровневый счётчик номеров разделов: «1», «1.1», «1.1.1».

    Уровней четыре — раздел, подраздел, пункт и подпункт.
    """

    def __init__(self) -> None:
        self.__levels_counter: dict[int, Counter] = {
            1: Counter(),
            2: Counter(),
            3: Counter(),
            4: Counter(),
        }

    def get_next(self, head_level: int) -> str:
        """Занимает следующий номер на *head_level* и возвращает его целиком.

        Args:
            head_level: Уровень заголовка 0..4. Уровень 0 — структурный
                элемент, он не нумеруется.

        Returns:
            Номер со всеми старшими уровнями («1.2.1») или пустая строка для
            уровня 0.
        """
        # Уровень 0 — структурный элемент, он не нумеруется и не участвует
        # в нумерации разделов.
        if head_level == 0:
            return ""

        self.__levels_counter[head_level].get_next()

        # Появление нового значения на текущем уровне начинает нумерацию всех
        # более низких уровней заново: 1.1, 1.2, затем 2 -> 2.1, а не 2.3.
        for level in self.__levels_counter:
            if level > head_level:
                self.__levels_counter[level].reset()

        result_number = str(self.__levels_counter[1].current)
        if head_level > 1:
            for level in range(2, head_level + 1):
                result_number = result_number + '.' + str(self.__levels_counter[level].current)

        return result_number

    def get_current_chapter(self) -> int:
        """Номер раздела, в котором мы сейчас находимся — старший уровень."""
        return self.__levels_counter[1].current

    def reset(self):
        """Возвращает в начало счётчики всех уровней."""
        for level in self.__levels_counter:
            self.__levels_counter[level].reset()
