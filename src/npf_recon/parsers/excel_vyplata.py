"""
Парсер Excel-файла «Выплата».

Читает файл с заголовком, ищет колонки по именам заголовков, маппирует
ВидВыплат на показатель через словарь из правила.
"""
from __future__ import annotations

import logging

import pandas as pd

from config.mappings import FileRule
from npf_recon.models import Record
from npf_recon.normalize import parse_date, parse_number
from npf_recon.sources.base import RawDocument
from .base import Parser

logger = logging.getLogger(__name__)


class ExcelVyplataParser(Parser):
    """
    Парсер файла выплат (ПУ).

    Ожидает в params:
      date_col_header   (str) — имя колонки с датой операции.
      type_col_header   (str) — имя колонки с видом выплат.
      amount_col_header (str) — имя колонки с суммой.
      type_map          (dict[str, str]) — маппинг вида выплаты на показатель.
    """

    def parse(self, raw: RawDocument, rule: FileRule) -> list[Record]:
        params = rule.params
        date_col_header: str = params["date_col_header"]
        type_col_header: str = params["type_col_header"]
        amount_col_header: str = params["amount_col_header"]
        type_map: dict[str, str] = params["type_map"]

        try:
            df = pd.read_excel(
                raw.path,
                engine="openpyxl",
                dtype=object,
            )
        except Exception as exc:
            logger.error("Ошибка чтения файла %s: %s", raw.path, exc)
            return []

        # Поиск колонок по имени (сравниваем без учёта пробелов по краям)
        col_map: dict[str, str] = {}
        for col in df.columns:
            col_stripped = str(col).strip()
            for needed in (date_col_header, type_col_header, amount_col_header):
                if col_stripped == needed:
                    col_map[needed] = col

        missing = [h for h in (date_col_header, type_col_header, amount_col_header)
                   if h not in col_map]
        if missing:
            logger.error(
                "Файл %s: не найдены колонки %s (доступны: %s)",
                raw.path.name, missing, list(df.columns)
            )
            return []

        records: list[Record] = []
        for idx, row in df.iterrows():
            raw_date = row[col_map[date_col_header]]
            raw_type = row[col_map[type_col_header]]
            raw_amount = row[col_map[amount_col_header]]

            date = parse_date(raw_date)
            if date is None:
                continue

            type_str = str(raw_type).strip() if raw_type is not None else ""
            indicator = type_map.get(type_str)
            if indicator is None:
                logger.debug(
                    "%s строка %s: вид выплаты '%s' не в маппинге, пропускаем",
                    raw.path.name, idx, type_str
                )
                continue

            amount = parse_number(raw_amount)
            if amount is None:
                logger.debug(
                    "%s строка %s: пустая сумма, пропускаем",
                    raw.path.name, idx
                )
                continue

            records.append(
                Record(
                    indicator=indicator,
                    date=date,
                    amount=amount,
                    side=rule.side,
                    source_file=raw.path.name,
                )
            )

        logger.info(
            "[%s] %s: прочитано %d записей",
            rule.side, raw.path.name, len(records)
        )
        return records
