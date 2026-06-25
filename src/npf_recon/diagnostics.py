"""
Диагностика парсинга — наглядная демонстрация того, что КАЖДЫЙ входной файл
(Excel из ПУ, Excel из БУ и JSON из БУ) корректно распознаётся, разбирается,
а извлечённые данные правильно распределяются по показателям и суммируются.

Используется:
  — из CLI: `python -m npf_recon --diagnose`;
  — из меню (пункт «Диагностика парсинга»);
  — из скрипта scripts/demo_parsing.py.

Выводит три блока:
  1) Пофайловый разбор: файл → сторона → парсер → показатели → кол-во записей и суммы.
  2) Матрица покрытия: по каждому из 14 показателей — какие стороны (БУ/ПУ) заполнены.
  3) Проверка ожидаемых источников: все ли показатели получили данные из нужных файлов.
"""
from __future__ import annotations

import logging
from collections import defaultdict

from config.mappings import REPORT_INDICATORS
from npf_recon.config import Config
from npf_recon.models import Record
from npf_recon.normalize import format_amount
from npf_recon.parsers.registry import get_parser, match_rule
from npf_recon.sources.filesystem import FileSystemSource

logger = logging.getLogger(__name__)

# Ожидаемые источники по показателям (для финальной проверки покрытия).
EXPECTED_SOURCES: dict[str, set[str]] = {
    "Пенсионные взносы ФЛ": {"БУ", "ПУ"},
    "Пенсионные взносы ЮЛ": {"БУ", "ПУ"},
    "Целевые взносы ФЛ": {"БУ", "ПУ"},
    "Целевые взносы ЮЛ": {"БУ", "ПУ"},
    "Страховой резерв": {"БУ", "ПУ"},
    "Выплата пенсии": {"ПУ"},
    "Выплата выкупных сумм": {"ПУ"},
    "Выплата наследуемых сумм": {"ПУ"},
}


def _hr(char: str = "─", width: int = 78) -> str:
    return char * width


def run_diagnostics(config: Config) -> bool:
    """
    Запускает диагностику парсинга и печатает отчёт.

    :param config: конфигурация с путями ПУ/БУ.
    :return: True, если все ожидаемые источники покрыты; иначе False.
    """
    print("\n" + _hr("═"))
    print("  ДИАГНОСТИКА ПАРСИНГА — проверка распознавания и извлечения данных")
    print(_hr("═"))

    source = FileSystemSource(pu_dir=config.pu_dir, bu_dir=config.bu_dir)
    raw_docs = source.scan()

    if not raw_docs:
        print("\n⚠  Файлы не найдены. Поместите файлы в каталоги ПУ и БУ или "
              "запустите генерацию тестовых данных.")
        return False

    # (indicator, side) -> [records]
    coverage: dict[tuple[str, str], list[Record]] = defaultdict(list)
    all_records: list[Record] = []

    # ---------- 1. Пофайловый разбор ----------
    print("\n1) ПОФАЙЛОВЫЙ РАЗБОР")
    print(_hr())
    for raw in raw_docs:
        rule = match_rule(raw.base_name, raw.side)
        if rule is None:
            print(f"  [{raw.side}] {raw.path.name}")
            print(f"        формат: {raw.fmt:5s} | ⚠ нет правила — файл пропущен")
            continue

        parser = get_parser(rule)
        records = parser.parse(raw, rule)
        all_records.extend(records)

        by_ind: dict[str, list[Record]] = defaultdict(list)
        for r in records:
            by_ind[r.indicator].append(r)
            coverage[(r.indicator, r.side)].append(r)

        print(f"  [{raw.side}] {raw.path.name}")
        print(f"        формат: {raw.fmt:5s} | парсер: {rule.parser_name:14s} | "
              f"записей: {len(records)}")
        for ind in sorted(by_ind):
            recs = by_ind[ind]
            total = sum(r.amount for r in recs)
            days = len({r.date for r in recs})
            print(f"          • {ind:32s} строк: {len(recs):3d} | "
                  f"дней: {days:2d} | сумма: {format_amount(total)}")

    # ---------- 2. Матрица покрытия показателей ----------
    print("\n2) МАТРИЦА ПОКРЫТИЯ ПОКАЗАТЕЛЕЙ (источник → показатель)")
    print(_hr())
    print(f"  {'Показатель':42s} {'БУ':>14s}  {'ПУ':>14s}")
    print("  " + _hr("·", 74))
    for ind in REPORT_INDICATORS:
        bu_recs = coverage.get((ind, "БУ"), [])
        pu_recs = coverage.get((ind, "ПУ"), [])
        bu_str = format_amount(sum(r.amount for r in bu_recs)) if bu_recs else "—"
        pu_str = format_amount(sum(r.amount for r in pu_recs)) if pu_recs else "—"
        print(f"  {ind:42s} {bu_str:>14s}  {pu_str:>14s}")

    # ---------- 3. Проверка ожидаемых источников ----------
    print("\n3) ПРОВЕРКА ОЖИДАЕМЫХ ИСТОЧНИКОВ")
    print(_hr())
    all_ok = True
    for ind, expected_sides in EXPECTED_SOURCES.items():
        present = {side for (i, side) in coverage if i == ind}
        missing = expected_sides - present
        if missing:
            all_ok = False
            print(f"  ✗ {ind:42s} отсутствует: {', '.join(sorted(missing))}")
        else:
            print(f"  ✓ {ind:42s} источники: {', '.join(sorted(present))}")

    print("\n" + _hr("═"))
    if all_ok:
        print(f"  ИТОГ: все ожидаемые источники распознаны. "
              f"Всего записей: {len(all_records)}.")
    else:
        print("  ИТОГ: ⚠ часть ожидаемых источников не найдена (см. выше).")
    print(_hr("═") + "\n")

    return all_ok
