"""Замер длительности операций для debug-логов."""

import logging
import time
from collections.abc import Iterator
from contextlib import contextmanager


@contextmanager
def logged_duration(logger: logging.Logger, message: str, *args: object) -> Iterator[None]:
    """Пишет в debug длительность блока: «<message> за 1.234 с».

    Args:
        message: Шаблон в стиле logging, с %-подстановками под args.
        args: Значения для message. Подставляются лениво, средствами logging, —
            если debug выключен, сообщение не форматируется вовсе.
    """
    start = time.perf_counter()
    try:
        yield
    finally:
        logger.debug(message + " за %.3f с", *args, time.perf_counter() - start)
