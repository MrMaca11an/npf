"""
Парсер JSON-файла НПО (бухгалтерский учёт).

Ожидаемая структура JSON:
{
  "report": {
    "startDate": "...",
    "endDate": "...",
    "items": [
      {
        "date": "2026-02-01T00:00:00",
        "amount": "...",
        "components": [
          {
            "account": "76.01",
            "name": "Пенсионные взносы ФЛ по договорам НПО",
            "amount": "123456.78"
          }
        ]
      }
    ]
  }
}

Каждый компонент маппируется на показатель через словарь из правила.
"""
from __future__ import annotations

import json
import logging
from pathlib import Path

from config.mappings import FileRule
from npf_recon.models import Record
from npf_recon.normalize import parse_date, parse_number
from npf_recon.sources.base import RawDocument
from .base import Parser

logger = logging.getLogger(__name__)


class JsonNpoParser(Parser):
    """
    Парсер JSON-файла с данными НПО.

    Ожидает в params:
      component_map (dict[str, str]) — маппинг имени компонента на показатель.
    """

    def parse(self, raw: RawDocument, rule: FileRule) -> list[Record]:
        component_map: dict[str, str] = rule.params["component_map"]

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

        records: list[Record] = []
        for item in items:
            raw_date = item.get("date")
            date = parse_date(raw_date)
            if date is None:
                logger.debug("Пропуск элемента без даты: %s", item)
                continue

            components = item.get("components", [])
            for comp in components:
                comp_name = str(comp.get("name", "")).strip()
                indicator = component_map.get(comp_name)
                if indicator is None:
                    logger.debug(
                        "Компонент '%s' не в маппинге, пропускаем", comp_name
                    )
                    continue

                amount = parse_number(comp.get("amount"))
                if amount is None:
                    logger.debug(
                        "Компонент '%s' дата %s: пустая сумма, пропускаем",
                        comp_name, date
                    )
                    continue

                records.append(
                    Record(
                        indicator=indicator,
                        date=date,
                        amount=amount,
                        side=rule.side,
                        source_file=raw.path.name,
                        extra={"account": comp.get("account", "")},
                    )
                )

        logger.info(
            "[%s] %s: прочитано %d записей",
            rule.side, raw.path.name, len(records)
        )
        return records
