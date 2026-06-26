"""
Парсер Excel-файла «Выплата».

Читает файл с заголовком, ищет колонки по именам заголовков, маппирует
ВидВыплат на показатель через словарь из правила.
"""
from __future__ import annotations

import logging
import re

import pandas as pd

from config.mappings import FileRule
from npf_recon.models import Record
from npf_recon.normalize import parse_date, parse_number
from npf_recon.sources.base import RawDocument
from .base import Parser

logger = logging.getLogger(__name__)


def _norm_header(value: object) -> str:
    """
    Нормализует заголовок колонки для устойчивого сравнения:
    нижний регистр, схлопывание любых пробельных последовательностей,
    удаление пробелов по краям.

    «Дата  операции » и «дата операции» → «дата операции».
    """
    return re.sub(r"\s+", " ", str(value)).strip().lower()


def _as_aliases(value: object) -> list[str]:
    """Приводит значение параметра-заголовка к списку синонимов."""
    if isinstance(value, (list, tuple)):
        return [str(v) for v in value]
    return [str(value)]


def _alnum(value: object) -> str:
    """Агрессивная нормализация: только буквы/цифры в нижнем регистре.

    Снимает различия в пробелах и пунктуации:
      «Выкупная сумма(Расторжение)» и «Выкупная сумма (Расторжение)»
      → «выкупнаясуммарасторжение».
    """
    return re.sub(r"[^0-9a-zа-яё]", "", str(value).lower())


def _build_type_matcher(params: dict):
    """
    Строит функцию сопоставления значения ВидВыплат с показателем, устойчивую к
    опечаткам, пробелам и пунктуации реальных выгрузок.

    Порядок проверки:
      1) точное совпадение по агрессивно-нормализованному значению (снимает
         различия в пробелах/пунктуации, напр. «сумма(» vs «сумма (»);
      2) правила по ключевым словам (переживают опечатки в середине слова,
         напр. «Негосударсвенная пенсия» → ключ «пенси»);
      3) запасной вариант — вхождение по подстроке (выигрывает длинный ключ).
    """
    type_map: dict[str, str] = params.get("type_map", {})
    keyword_rules = params.get("type_keyword_rules") or []

    norm_exact = {_alnum(k): v for k, v in type_map.items()}
    norm_rules = [([_alnum(k) for k in kws], ind) for kws, ind in keyword_rules]
    keys_by_len = sorted((k for k in norm_exact if k), key=len, reverse=True)

    def match(value: object) -> str | None:
        v = _alnum(value)
        if not v:
            return None
        if v in norm_exact:
            return norm_exact[v]
        for kws, indicator in norm_rules:
            if all(kw in v for kw in kws):
                return indicator
        for key in keys_by_len:
            if key in v or v in key:
                return norm_exact[key]
        return None

    return match


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
        # Каждый «логический» столбец может задаваться как строка или список
        # синонимов (aliases) — берётся первый совпавший заголовок.
        date_aliases = _as_aliases(params["date_col_header"])
        type_aliases = _as_aliases(params["type_col_header"])
        amount_aliases = _as_aliases(params["amount_col_header"])
        match_type = _build_type_matcher(params)

        try:
            df = pd.read_excel(
                raw.path,
                engine="openpyxl",
                dtype=object,
            )
        except Exception as exc:
            logger.error("Ошибка чтения файла %s: %s", raw.path, exc)
            return []

        # Индекс заголовков: нормализованное имя → реальное имя колонки
        norm_to_real = {_norm_header(col): col for col in df.columns}

        def _find(aliases: list[str]) -> object | None:
            for alias in aliases:
                real = norm_to_real.get(_norm_header(alias))
                if real is not None:
                    return real
            return None

        date_col = _find(date_aliases)
        type_col = _find(type_aliases)
        amount_col = _find(amount_aliases)

        if date_col is None or type_col is None or amount_col is None:
            missing = []
            if date_col is None:
                missing.append(date_aliases)
            if type_col is None:
                missing.append(type_aliases)
            if amount_col is None:
                missing.append(amount_aliases)
            logger.error(
                "Файл %s: не найдены колонки %s (доступны: %s)",
                raw.path.name, missing, list(df.columns)
            )
            return []

        records: list[Record] = []
        for idx, row in df.iterrows():
            raw_date = row[date_col]
            raw_type = row[type_col]
            raw_amount = row[amount_col]

            date = parse_date(raw_date)
            if date is None:
                continue

            type_str = str(raw_type).strip() if raw_type is not None else ""
            indicator = match_type(type_str)
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
