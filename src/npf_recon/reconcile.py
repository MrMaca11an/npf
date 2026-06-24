"""
Сверка БУ и ПУ по каждому показателю.

Для каждого из 14 фиксированных показателей:
  — если есть данные обеих сторон → вычисляем расхождение по итогу и датам;
  — если есть только одна сторона → расхождение не вычисляется (diff_total=None);
  — если нет данных вообще → пустая строка (обе стороны None).

Результат: список ReconRow в том же фиксированном порядке.
"""
from __future__ import annotations

import logging
from datetime import date

from config.mappings import REPORT_INDICATORS
from npf_recon.models import IndicatorSide, ReconRow

logger = logging.getLogger(__name__)

GroupKey = tuple[str, str]


def reconcile(
    aggregated: dict[GroupKey, IndicatorSide],
    eps: float = 0.005,
) -> list[ReconRow]:
    """
    Производит сверку по всем 14 показателям.

    :param aggregated: словарь {(indicator, side): IndicatorSide} из aggregate().
    :param eps:        допуск для сравнения сумм (расхождение |d| <= eps игнорируется).
    :return:           список из 14 ReconRow в порядке REPORT_INDICATORS.
    """
    rows: list[ReconRow] = []

    for indicator in REPORT_INDICATORS:
        bu = aggregated.get((indicator, "БУ"))
        pu = aggregated.get((indicator, "ПУ"))

        if bu is not None and pu is not None:
            # Обе стороны присутствуют — вычисляем расхождения
            diff_total = round(bu.total - pu.total, 2)

            all_dates = sorted(set(bu.dates) | set(pu.dates))
            diff_dates: list[date] = []
            diff_by_date: dict[date, float] = {}

            for d in all_dates:
                bu_val = bu.daily.get(d, 0.0)
                pu_val = pu.daily.get(d, 0.0)
                diff = round(bu_val - pu_val, 2)
                if abs(diff) > eps:
                    diff_dates.append(d)
                    diff_by_date[d] = diff

            logger.info(
                "Показатель '%s': БУ=%s, ПУ=%s, расхождение=%s, дат с расхождением=%d",
                indicator, bu.total, pu.total, diff_total, len(diff_dates)
            )

        else:
            # Одна или обе стороны отсутствуют — расхождение не вычисляется
            diff_total = None
            diff_dates = []
            diff_by_date = {}

            if bu is not None or pu is not None:
                present_side = "БУ" if bu is not None else "ПУ"
                logger.info(
                    "Показатель '%s': присутствует только %s — расхождение не вычисляется",
                    indicator, present_side
                )
            else:
                logger.debug("Показатель '%s': нет данных", indicator)

        rows.append(
            ReconRow(
                indicator=indicator,
                bu=bu,
                pu=pu,
                diff_total=diff_total,
                diff_dates=diff_dates,
                diff_by_date=diff_by_date,
            )
        )

    return rows
