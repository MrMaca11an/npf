"""Базовая абстракция парсера документов."""
from __future__ import annotations

from abc import ABC, abstractmethod

from npf_recon.models import Record
from npf_recon.sources.base import RawDocument


class Parser(ABC):
    """
    Абстрактный парсер: преобразует RawDocument в список Record.

    Каждый конкретный парсер знает структуру одного формата файла и
    возвращает нормализованные записи, одинаковые для всей системы.
    """

    @abstractmethod
    def parse(self, raw: RawDocument, rule: object) -> list[Record]:
        """
        Парсит документ согласно правилу из реестра.

        :param raw:  метаданные документа.
        :param rule: правило FileRule из config/mappings.py.
        :return:     список нормализованных Record.
        """
        ...
