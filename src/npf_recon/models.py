"""
Доменные модели (dataclasses) системы сверки БУ-ПУ.

Все слои бизнес-логики (агрегация, сверка, отчёт) работают только с этими
структурами данных, не зная ничего о формате источника (Excel, JSON, API).
"""
from __future__ import annotations

import datetime
from dataclasses import dataclass, field
from typing import Literal


@dataclass
class Record:
    """
    Нормализованная запись из любого источника данных.

    После парсинга все данные приведены к единообразному виду:
    дата — datetime.date, сумма — float, сторона — «БУ» или «ПУ».
    """

    indicator: str            # Название показателя (из REPORT_INDICATORS)
    date: datetime.date       # Дата операции
    amount: float             # Сумма в рублях
    side: Literal["БУ", "ПУ"]  # Источник: бухгалтерский или персонифицированный учёт
    source_file: str          # Имя файла-источника (для отладки)
    extra: dict = field(default_factory=dict)  # Доп. атрибуты (напр., «account» из JSON)


@dataclass
class IndicatorSide:
    """
    Агрегированные данные по одному показателю и одной стороне (БУ или ПУ).

    Хранит суммы по дням и итог за период.
    """

    daily: dict[datetime.date, float]   # Сумма по каждому дню
    total: float                         # Итог за период (сумма всех daily)
    dates: list[datetime.date]           # Список дат в хронологическом порядке

    @classmethod
    def from_records(cls, records: list[Record]) -> "IndicatorSide":
        """
        Строит IndicatorSide из списка Record для одного показателя и стороны.

        :param records: записи (все должны быть одного indicator и side).
        :return: объект IndicatorSide.
        """
        daily: dict[datetime.date, float] = {}
        for rec in records:
            daily[rec.date] = daily.get(rec.date, 0.0) + rec.amount

        dates = sorted(daily.keys())
        total = round(sum(daily.values()), 2)
        return cls(daily=daily, total=total, dates=dates)


@dataclass
class ReconRow:
    """
    Строка результата сверки по одному показателю.

    Содержит агрегаты БУ и ПУ, а также расхождения.
    """

    indicator: str                        # Название показателя
    bu: IndicatorSide | None              # Данные БУ (None — нет данных)
    pu: IndicatorSide | None              # Данные ПУ (None — нет данных)
    diff_total: float | None              # Разница итогов (БУ − ПУ); None — не сравнивается
    diff_dates: list[datetime.date]       # Даты с расхождениями
    diff_by_date: dict[datetime.date, float]  # Расхождение по каждой дате
