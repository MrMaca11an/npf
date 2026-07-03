"""
Агрегация нормализованных записей по показателю и стороне.

Принимает плоский список Record, группирует по (indicator, side) и
строит IndicatorSide для каждой группы.
"""
from __future__ import annotations

import logging
from collections import defaultdict

from npf_recon.models import IndicatorSide, Record

logger = logging.getLogger(__name__)

# Тип индекса: (indicator, side) → список Record
GroupKey = tuple[str, str]


def group_records(records: list[Record]) -> dict[GroupKey, list[Record]]:
    """
    Группирует список Record по паре (indicator, side).

    :param records: плоский список нормализованных записей.
    :return: словарь {(indicator, side): [Record, ...]}.
    """
    groups: dict[GroupKey, list[Record]] = defaultdict(list)
    for rec in records:
        groups[(rec.indicator, rec.side)].append(rec)
    return dict(groups)


def build_indicator_side(records: list[Record]) -> IndicatorSide:
    """
    Строит IndicatorSide из записей одного (indicator, side).

    :param records: записи одного показателя и стороны.
    :return: агрегированный IndicatorSide.
    """
    return IndicatorSide.from_records(records)


def aggregate(records: list[Record]) -> dict[GroupKey, IndicatorSide]:
    """
    Полная агрегация: группировка + построение IndicatorSide.

    :param records: все нормализованные записи из всех файлов.
    :return: словарь {(indicator, side): IndicatorSide}.
    """
    groups = group_records(records)
    result: dict[GroupKey, IndicatorSide] = {}
    for key, recs in groups.items():
        result[key] = build_indicator_side(recs)
        logger.debug(
            "Агрегирован показатель [%s] сторона [%s]: итог=%s, дат=%d",
            key[0], key[1], result[key].total, len(result[key].dates)
        )
    return result
