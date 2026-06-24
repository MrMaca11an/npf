"""Слой парсеров: преобразует RawDocument в нормализованные Record."""
from .base import Parser
from .registry import match_rule, get_parser

__all__ = ["Parser", "match_rule", "get_parser"]
