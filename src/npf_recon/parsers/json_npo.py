"""
Парсер JSON-файла НПО (бухгалтерский учёт).

Поддерживает ДВЕ структуры JSON (автоопределяется по форме данных):

1) «Плоская» (старая):
{
  "report": {
    "items": [
      {
        "date": "2026-02-01T00:00:00",
        "components": [
          {"account": "76.01", "name": "Пенсионные взносы ФЛ по договорам НПО", "amount": "123456.78"}
        ]
      }
    ]
  }
}

2) «Вложенная по разделам» (новая, реальная выгрузка):
{
  "report": {
    "items": [
      {
        "section": "НПО",
        "items": [
          {
            "indicator": "Целевые взносы ФЛ",
            "amount": 961697.12,
            "items": [
              {"date": "2026-02-01T00:00:00", "amount": 15},
              {"date": "2026-02-02T00:00:00", "amount": 77833.19}
            ]
          },
          {"indicator": "Остатки на начало", "date": "2026-02-01T00:00:00", "amount": 126379073074.24}
        ]
      },
      {"section": "СПН", "items": [...]}
    ]
  }
}

Во вложенной структуре «indicator» — это либо показатель с разбивкой по дням
(вложенный «items»: [{date, amount}]), либо разовый остаток на дату (свои
«date» + «amount», без вложенных items). Разделы («section») не входящие в
наш реестр показателей (напр. «СПН» — обязательное пенсионное страхование,
не относится к 14 строкам отчёта) не содержат подходящих индикаторов и
поэтому естественным образом пропускаются (индикатор не находит соответствия).

Сопоставление имени показателя из файла с показателем отчёта — в три уровня
приоритета:
  1) ТОЧНОЕ совпадение с названием строки отчёта (REPORT_INDICATORS) — расчёт
     на то, что источник рано или поздно приведёт ключи к точным названиям
     строк отчёта (в т.ч. для РППО инвестиционный/страховой, начало/конец).
  2) Известные варианты названий показателей/выплат (component_map).
  3) Известные варианты названий возвратов (deduction_map) — вычитаются.
Если ничего не подошло — компонент пропускается (лог на уровне DEBUG).
"""
from __future__ import annotations

import json
import logging
import re
from pathlib import Path

from config.mappings import FileRule, REPORT_INDICATORS
from npf_recon.models import Record
from npf_recon.normalize import parse_date, parse_number
from npf_recon.sources.base import RawDocument
from .base import Parser

logger = logging.getLogger(__name__)


def _norm(value: object) -> str:
    """Нормализует имя показателя/компонента: нижний регистр, схлопывание пробелов."""
    return re.sub(r"\s+", " ", str(value)).strip().lower()


def _build_matcher(name_map: dict[str, str], *, exact_only: bool = False):
    """
    Строит функцию сопоставления имени показателя из файла с показателем отчёта.

    :param name_map:   {имя_в_файле: показатель_отчёта}.
    :param exact_only: если True — только точное совпадение (без подстрок).
                        Используется для сопоставления с REPORT_INDICATORS:
                        подстроки здесь опасны («ИД» — короткий ключ, который
                        иначе совпал бы с любым именем, содержащим буквы «ид»,
                        напр. «Начисление ИД на СПН» из несвязанного раздела).

    Для остальных карт (component_map/deduction_map) допускается сопоставление
    по подстроке — их ключи длинные и специфичные, ложные совпадения исключены.
    """
    norm_map = {_norm(k): v for k, v in name_map.items()}
    keys_by_len = sorted(norm_map, key=len, reverse=True)

    def match(name: str) -> str | None:
        key = _norm(name)
        if key in norm_map:
            return norm_map[key]
        if exact_only:
            return None
        for k in keys_by_len:
            if k in key:
                return norm_map[k]
        return None

    return match


class JsonNpoParser(Parser):
    """
    Парсер JSON-файла с данными НПО (бухгалтерский учёт).

    Ожидает в params:
      component_map (dict[str, str]) — известные названия показателей/выплат.
      deduction_map (dict[str, str]) — известные названия возвратов (вычитаются).
    """

    def parse(self, raw: RawDocument, rule: FileRule) -> list[Record]:
        component_map: dict[str, str] = rule.params["component_map"]
        deduction_map: dict[str, str] = rule.params.get("deduction_map", {})

        match_report_row = _build_matcher(
            {name: name for name in REPORT_INDICATORS}, exact_only=True
        )
        match_add = _build_matcher(component_map)
        match_deduct = _build_matcher(deduction_map)

        def resolve(name: str) -> tuple[str | None, float]:
            """Возвращает (показатель, знак). Знак −1 — для возвратов."""
            indicator = match_report_row(name)
            if indicator is not None:
                return indicator, 1.0
            indicator = match_add(name)
            if indicator is not None:
                return indicator, 1.0
            indicator = match_deduct(name)
            if indicator is not None:
                return indicator, -1.0
            return None, 1.0

        try:
            with open(raw.path, encoding="utf-8") as f:
                data = json.load(f)
        except Exception as exc:
            logger.error("Ошибка чтения JSON %s: %s", raw.path, exc)
            return []

        try:
            items = data["report"]["items"]
        except (KeyError, TypeError) as exc:
            logger.error(
                "Файл %s: неожиданная структура JSON (ожидается report.items): %s",
                raw.path.name, exc
            )
            return []

        if _is_nested_schema(items):
            records = _parse_nested_schema(items, resolve, rule.side, raw.path.name)
        else:
            records = _parse_flat_schema(items, resolve, rule.side, raw.path.name)

        logger.info(
            "[%s] %s: прочитано %d записей",
            rule.side, raw.path.name, len(records)
        )
        return records


# ------------------------------------------------------------------
# Схема 1 (старая): report.items[].{date, components[]}
# ------------------------------------------------------------------

def _parse_flat_schema(items, resolve, side: str, filename: str) -> list[Record]:
    records: list[Record] = []
    for item in items:
        raw_date = item.get("date")
        date = parse_date(raw_date)
        if date is None:
            logger.debug("Пропуск элемента без даты: %s", item)
            continue

        for comp in item.get("components", []):
            comp_name = str(comp.get("name", "")).strip()
            indicator, sign = resolve(comp_name)
            if indicator is None:
                logger.debug("Компонент '%s' не в маппинге, пропускаем", comp_name)
                continue

            amount = parse_number(comp.get("amount"))
            if amount is None:
                logger.debug(
                    "Компонент '%s' дата %s: пустая сумма, пропускаем",
                    comp_name, date
                )
                continue

            amount = -abs(amount) if sign < 0 else amount
            records.append(
                Record(
                    indicator=indicator,
                    date=date,
                    amount=amount,
                    side=side,
                    source_file=filename,
                    extra={"account": comp.get("account", "")},
                )
            )
    return records


# ------------------------------------------------------------------
# Схема 2 (новая): report.items[].{section, items[].{indicator, amount, items[]|date}}
# ------------------------------------------------------------------

def _is_nested_schema(items) -> bool:
    """Определяет схему по форме первых элементов report.items."""
    for it in items:
        if not isinstance(it, dict):
            continue
        if "components" in it:
            return False
        if "section" in it or "indicator" in it:
            return True
    return False


def _parse_nested_schema(items, resolve, side: str, filename: str) -> list[Record]:
    records: list[Record] = []

    def walk(entries) -> None:
        for entry in entries:
            if not isinstance(entry, dict):
                continue

            # Обёртка раздела («section»: «НПО»/«СПН») — спускаемся внутрь.
            if "section" in entry and isinstance(entry.get("items"), list):
                walk(entry["items"])
                continue

            if "indicator" not in entry:
                nested = entry.get("items")
                if isinstance(nested, list):
                    walk(nested)
                continue

            name = str(entry["indicator"]).strip()
            indicator, sign = resolve(name)
            if indicator is None:
                # Показатели из несвязанных разделов (напр. «СПН» — ОПС, не
                # относится к нашим 14 строкам отчёта) закономерно не находят
                # соответствия и пропускаются здесь.
                logger.debug("Показатель '%s' не в маппинге, пропускаем", name)
                continue

            daily_items = entry.get("items")
            if isinstance(daily_items, list) and daily_items:
                # Показатель с разбивкой по дням
                for leaf in daily_items:
                    if not isinstance(leaf, dict):
                        continue
                    date = parse_date(leaf.get("date"))
                    amount = parse_number(leaf.get("amount"))
                    if date is None or amount is None:
                        continue
                    amount = -abs(amount) if sign < 0 else amount
                    records.append(
                        Record(indicator=indicator, date=date, amount=amount,
                               side=side, source_file=filename)
                    )
            elif "date" in entry:
                # Разовый остаток на дату (напр. «Остатки на начало/конец»)
                date = parse_date(entry.get("date"))
                amount = parse_number(entry.get("amount"))
                if date is not None and amount is not None:
                    amount = -abs(amount) if sign < 0 else amount
                    records.append(
                        Record(indicator=indicator, date=date, amount=amount,
                               side=side, source_file=filename)
                    )
            else:
                logger.debug(
                    "Показатель '%s': нет ни дневной разбивки, ни даты — пропускаем",
                    name
                )

    walk(items)
    return records
