"""
Парсер оборотно-сальдовой ведомости (ОСВ) — БУ «Страховой резерв», счёт 397.03.

Файл выгружается из 1С как ОСВ по счёту: множество служебных строк-заголовков,
группировка по стратегиям/лицевым счетам и строки оборотов по дням вида
«Обороты за DD.MM.YY». Показатель «Страховой резерв» за день и за период равен
   оборот = Кредит − Дебет.

Парсер берёт ТОЛЬКО дневные строки «Обороты за DD.MM.YY» (по ним же агрегацией
получается итог за период). Итоговые строки («Итого», «397.03», «Обороты за
период», подытоги по стратегиям) автоматически пропускаются, так как не
содержат даты после слова «за» — это исключает двойной счёт.

Колонки оборотов (Дебет/Кредит) определяются автоматически по шапке
«Обороты за период»; при неудаче берутся явные значения из правила
(debit_col / credit_col).
"""
from __future__ import annotations

import logging
import re

import pandas as pd
from openpyxl.utils import column_index_from_string

from config.mappings import FileRule
from npf_recon.models import Record
from npf_recon.normalize import parse_date, parse_number
from npf_recon.sources.base import RawDocument
from .base import Parser

logger = logging.getLogger(__name__)

# «Обороты за 10.02.26», «Оборот за 10.02.2026» — дневная строка с датой.
_DAILY_RE = re.compile(r"оборот\w*\s+за\s+(\d{1,2}\.\d{1,2}\.\d{2,4})", re.IGNORECASE)


def _norm(value: object) -> str:
    return re.sub(r"\s+", " ", str(value)).strip().lower()


def _resolve_col(spec: object) -> int:
    """Приводит спецификацию колонки к 0-based индексу (число или буква Excel)."""
    if isinstance(spec, str) and spec.strip().isalpha():
        return column_index_from_string(spec.strip().upper()) - 1
    return int(spec)


class ExcelOsvParser(Parser):
    """
    Парсер ОСВ по счёту (БУ).

    Ожидает в params:
      indicator  (str)            — название показателя для всех записей.
      debit_col  (int|str, опц.)  — колонка «Обороты Дебет» (запасной вариант).
      credit_col (int|str, опц.)  — колонка «Обороты Кредит» (запасной вариант).
    """

    def parse(self, raw: RawDocument, rule: FileRule) -> list[Record]:
        params = rule.params
        indicator: str = params["indicator"]

        try:
            df = pd.read_excel(
                raw.path, header=None, engine="openpyxl", dtype=object
            )
        except Exception as exc:
            logger.error("Ошибка чтения файла %s: %s", raw.path, exc)
            return []

        debit_col, credit_col = self._resolve_columns(df, params)
        logger.debug(
            "%s: колонки оборотов Дебет=%s, Кредит=%s",
            raw.path.name, debit_col, credit_col,
        )

        records: list[Record] = []
        for idx, row in df.iterrows():
            match = self._match_daily(row)
            if match is None:
                continue
            date = parse_date(match)
            if date is None:
                continue

            debit = parse_number(row.iloc[debit_col]) if debit_col < len(row) else None
            credit = parse_number(row.iloc[credit_col]) if credit_col < len(row) else None
            amount = round((credit or 0.0) - (debit or 0.0), 2)
            if amount == 0.0:
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
            rule.side, raw.path.name, len(records), indicator,
        )
        return records

    # --------------------------------------------------------------
    @staticmethod
    def _match_daily(row) -> str | None:
        """Ищет в ячейках строки маркер «Обороты за DD.MM.YY», возвращает дату-строку."""
        for value in row:
            if value is None:
                continue
            m = _DAILY_RE.search(str(value))
            if m:
                return m.group(1)
        return None

    @staticmethod
    def _resolve_columns(df: pd.DataFrame, params: dict) -> tuple[int, int]:
        """
        Определяет 0-based индексы колонок «Обороты Дебет» и «Обороты Кредит».

        Сначала ищет шапку «Обороты за период» и под ней «Дебет»/«Кредит».
        При неудаче — значения из правила (debit_col / credit_col), иначе D/E.
        """
        detected = ExcelOsvParser._detect_oboroty_columns(df)
        if detected is not None:
            return detected

        debit_spec = params.get("debit_col", "D")
        credit_spec = params.get("credit_col", "E")
        return _resolve_col(debit_spec), _resolve_col(credit_spec)

    @staticmethod
    def _detect_oboroty_columns(df: pd.DataFrame) -> tuple[int, int] | None:
        """Авто-определение по шапке «Обороты за период» + «Дебет»/«Кредит»."""
        n_rows = len(df)
        for r in range(min(n_rows, 30)):
            for c in range(df.shape[1]):
                if _norm(df.iat[r, c]) == "обороты за период":
                    # «Дебет»/«Кредит» — в этой же или следующих 1-2 строках,
                    # начиная с колонки c.
                    debit_col = credit_col = None
                    for rr in range(r, min(r + 3, n_rows)):
                        for cc in range(c, min(c + 4, df.shape[1])):
                            label = _norm(df.iat[rr, cc])
                            if label == "дебет" and debit_col is None:
                                debit_col = cc
                            elif label == "кредит" and credit_col is None:
                                credit_col = cc
                    if debit_col is not None and credit_col is not None:
                        return debit_col, credit_col
        return None
