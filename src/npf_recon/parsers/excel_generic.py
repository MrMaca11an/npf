"""
Универсальный парсер Excel-файлов с позиционными колонками.

Читает файл без заголовка (header=None), берёт дату из колонки date_col
и сумму из колонки amount_col. Строки без валидной даты пропускаются
(заголовки, итоговые строки, пустые строки).
"""
from __future__ import annotations

import logging
from pathlib import Path

import pandas as pd

from config.mappings import FileRule
from npf_recon.models import Record
from npf_recon.normalize import parse_date, parse_number
from npf_recon.sources.base import RawDocument
from .base import Parser

logger = logging.getLogger(__name__)


class ExcelGenericParser(Parser):
    """
    Парсер Excel-файлов с позиционными колонками (без заголовка).

    Ожидает в params:
      date_col   (int) — 0-based индекс колонки с датой.
      amount_col (int) — 0-based индекс колонки с суммой.
      indicator  (str) — название показателя для всех записей файла.
    """

    def parse(self, raw: RawDocument, rule: FileRule) -> list[Record]:
        """
        :param raw:  метаданные документа.
        :param rule: правило с params: date_col, amount_col, indicator.
        :return:     список Record.
        """
        params = rule.params
        date_col: int = params["date_col"]
        amount_col: int = params["amount_col"]
        indicator: str = params["indicator"]

        try:
            df = pd.read_excel(
                raw.path,
                header=None,
                engine="openpyxl",
                dtype=object,  # Все ячейки как object — не преобразовываем автоматически
            )
        except Exception as exc:
            logger.error("Ошибка чтения файла %s: %s", raw.path, exc)
            return []

        records: list[Record] = []
        for idx, row in df.iterrows():
            raw_date = row.iloc[date_col] if date_col < len(row) else None
            raw_amount = row.iloc[amount_col] if amount_col < len(row) else None

            date = parse_date(raw_date)
            if date is None:
                # Пропускаем строку без валидной даты (заголовок, итог, пустая)
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
            "[%s] %s: прочитано %d записей (показатель: %s)",
            rule.side, raw.path.name, len(records), indicator
        )
        return records
