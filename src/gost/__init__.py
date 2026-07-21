import logging

from .word_builder import WordBuilder

# Библиотека не навязывает приложению настройку логирования: без обработчика
# записи просто отбрасываются. Всё пишется в логгер «gost».
logging.getLogger(__name__).addHandler(logging.NullHandler())

__all__ = ["WordBuilder"]
