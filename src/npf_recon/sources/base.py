"""
Базовые абстракции слоя источников данных.

Чтобы перейти с файловой системы на API, достаточно реализовать новый класс,
наследующий Source, — весь остальной код менять не нужно.
"""
from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass
from pathlib import Path
from typing import Literal


@dataclass
class RawDocument:
    """
    Ссылка на необработанный документ из произвольного источника.

    Для файловой системы path — реальный путь.
    Для API path может быть виртуальным идентификатором.
    """

    path: Path                    # Полный путь к файлу (или псевдо-путь для API)
    base_name: str                # Имя файла без расширения (в нижнем регистре)
    ext: str                      # Расширение без точки, нижний регистр («xlsx», «json»)
    fmt: Literal["excel", "json"] # Формат содержимого
    side: Literal["БУ", "ПУ"]    # Источник: бухгалтерский или персонифицированный учёт


class Source(ABC):
    """
    Абстрактный источник документов.

    Реализации: FileSystemSource (текущая), будущие ApiSource и т.д.
    """

    @abstractmethod
    def scan(self) -> list[RawDocument]:
        """
        Сканирует источник и возвращает список доступных документов.

        :return: список RawDocument для дальнейшего парсинга.
        """
        ...
