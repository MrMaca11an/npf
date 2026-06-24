"""
Реестр парсеров и сопоставление файлов с правилами.

match_rule   — находит правило FileRule для файла по (base_name, side).
get_parser   — возвращает экземпляр парсера по имени из правила.
"""
from __future__ import annotations

import sys
import os

# Добавляем корень репозитория в path, чтобы импортировать config/mappings
_REPO_ROOT = os.path.join(os.path.dirname(__file__), "..", "..", "..", "..")
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

from config.mappings import ALL_RULES, FileRule
from .base import Parser


def match_rule(base_name: str, side: str) -> FileRule | None:
    """
    Ищет правило для файла по его base_name (имя без расширения, нижний регистр)
    и стороне ('БУ' или 'ПУ').

    Проверяются только правила соответствующей стороны, что исключает
    путаницу (напр., «Страховой резерв» есть и в ПУ, и в БУ).

    :param base_name: имя файла без расширения в нижнем регистре.
    :param side:      'БУ' или 'ПУ'.
    :return:          подходящее FileRule или None.
    """
    rules = ALL_RULES.get(side, [])
    for rule in rules:
        if rule.name_contains.lower() in base_name:
            return rule
    return None


from functools import lru_cache


@lru_cache(maxsize=1)
def _build_registry() -> dict[str, Parser]:
    """
    Строит реестр парсеров один раз (кэшируется на время работы процесса).

    Экземпляры парсеров не хранят состояния между файлами, поэтому их можно
    переиспользовать — это исключает повторное создание объектов на каждый файл.
    """
    # Импорт здесь, чтобы избежать циклических зависимостей
    from .excel_generic import ExcelGenericParser
    from .excel_vyplata import ExcelVyplataParser
    from .json_npo import JsonNpoParser

    return {
        "excel_generic": ExcelGenericParser(),
        "vyplata": ExcelVyplataParser(),
        "json_npo": JsonNpoParser(),
    }


def get_parser(rule: FileRule) -> Parser:
    """
    Возвращает экземпляр парсера по имени из правила.

    :param rule: FileRule из реестра.
    :return:     готовый к работе Parser (переиспользуемый синглтон).
    :raises ValueError: если парсер с таким именем не зарегистрирован.
    """
    registry = _build_registry()
    parser = registry.get(rule.parser_name)
    if parser is None:
        raise ValueError(
            f"Парсер '{rule.parser_name}' не зарегистрирован. "
            f"Доступные парсеры: {list(registry.keys())}"
        )
    return parser
