"""
Главный конвейер обработки данных.

Этапы:
  1. Проверка наличия директорий (создаёт их при отсутствии).
  2. Сканирование файлов (FileSystemSource).
  3. Разбор файлов (Parser → Record).
  4. Агрегация (aggregate).
  5. Сверка (reconcile).
  6. Отчёт (write_report).
"""
from __future__ import annotations

import logging
import sys
from pathlib import Path

from npf_recon.aggregate import aggregate
from npf_recon.config import Config
from npf_recon.normalize import format_amount
from npf_recon.parsers.registry import get_parser, match_rule
from npf_recon.reconcile import reconcile
from npf_recon.report import write_report
from npf_recon.sources.filesystem import FileSystemSource

logger = logging.getLogger(__name__)


def run(config: Config) -> int:
    """
    Запускает полный конвейер обработки данных.

    :param config: конфигурация с путями и параметрами.
    :return:       код завершения (0 — успех, 1 — ошибка).
    """
    # ------------------------------------------------------------------
    # 1. Проверка/создание директорий
    # ------------------------------------------------------------------
    dirs_created = []
    for label, dir_path in [("ПУ", config.pu_dir), ("БУ", config.bu_dir)]:
        if not dir_path.exists():
            dir_path.mkdir(parents=True, exist_ok=True)
            dirs_created.append((label, dir_path))

    if dirs_created:
        print("\n⚠  Директории данных не существовали и были созданы:")
        for label, path in dirs_created:
            print(f"   [{label}] {path}")
        print(
            "\nПожалуйста, поместите входные файлы в указанные директории "
            "и запустите программу повторно.\n"
        )
        return 0

    # ------------------------------------------------------------------
    # 2. Сканирование файлов
    # ------------------------------------------------------------------
    source = FileSystemSource(pu_dir=config.pu_dir, bu_dir=config.bu_dir)
    raw_docs = source.scan()

    if not raw_docs:
        print(
            "\n⚠  Файлы данных не найдены в директориях:\n"
            f"   ПУ: {config.pu_dir}\n"
            f"   БУ: {config.bu_dir}\n"
            "Поместите файлы и запустите программу повторно.\n"
        )
        return 0

    # ------------------------------------------------------------------
    # 3. Парсинг файлов
    # ------------------------------------------------------------------
    all_records = []
    skipped = []
    for raw in raw_docs:
        rule = match_rule(raw.base_name, raw.side)
        if rule is None:
            logger.warning(
                "Файл не соответствует ни одному правилу [%s]: %s — пропускаем",
                raw.side, raw.path.name
            )
            skipped.append(raw.path.name)
            continue
        parser = get_parser(rule)
        records = parser.parse(raw, rule)
        all_records.extend(records)

    print(f"\nОбработано файлов: {len(raw_docs) - len(skipped)}")
    if skipped:
        print(f"Пропущено (нет правила): {len(skipped)}")
        for name in skipped:
            print(f"  — {name}")

    if not all_records:
        print("\n⚠  Не удалось извлечь ни одной записи из файлов.")
        return 1

    print(f"Итого записей: {len(all_records)}")

    # ------------------------------------------------------------------
    # 4. Агрегация
    # ------------------------------------------------------------------
    aggregated = aggregate(all_records)

    # ------------------------------------------------------------------
    # 5. Сверка
    # ------------------------------------------------------------------
    recon_rows = reconcile(aggregated, eps=config.eps)

    # ------------------------------------------------------------------
    # 6. Формирование отчёта
    # ------------------------------------------------------------------
    write_report(
        rows=recon_rows,
        output_path=config.output_file,
        period_label=config.period_label,
        eps=config.eps,
    )

    # ------------------------------------------------------------------
    # Итоговый вывод
    # ------------------------------------------------------------------
    print(f"\nОтчёт сохранён: {config.output_file}")
    print(f"Период: {config.period_label}")

    with_discrep = [r for r in recon_rows if r.diff_total is not None and abs(r.diff_total) > config.eps]
    if with_discrep:
        print(f"\nПоказатели с расхождениями ({len(with_discrep)}):")
        for r in with_discrep:
            dates_str = ", ".join(str(d) for d in r.diff_dates)
            print(f"  — {r.indicator}: {format_amount(r.diff_total)} ({dates_str})")
    else:
        print("\nРасхождений не обнаружено.")

    return 0
