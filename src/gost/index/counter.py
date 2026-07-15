class Counter:
    def __init__(self, initial: int = 1) -> None:
        self.__initial = initial - 1
        self.current: int = self.__initial

    def get_next(self) -> int:
        self.current = self.current + 1
        return self.current

    def reset(self):
        self.current = self.__initial


class HeadCounter:
    def __init__(self) -> None:
        self.__levels_counter: dict[int, Counter] = {
            1: Counter(),
            2: Counter(),
            3: Counter(),
            4: Counter(),
        }

    def get_next(self, head_level: int) -> str:
        # Уровень 0 — структурный элемент, он не нумеруется и не участвует
        # в нумерации разделов.
        if head_level == 0:
            return ""

        self.__levels_counter[head_level].get_next()

        result_number = str(self.__levels_counter[1].current)
        if head_level > 1:
            for level in range(2, head_level + 1):
                result_number = result_number + '.' + str(self.__levels_counter[level].current)

        return result_number

    def get_current_chapter(self) -> int:
        return self.__levels_counter[1].current

    def reset(self):
        for level in self.__levels_counter:
            self.__levels_counter[level].reset()

    def __get_currents(self) -> list[int]:
        currents: list[int] = []
        for level in self.__levels_counter:
            currents.append(self.__levels_counter[level].current)

        return currents
