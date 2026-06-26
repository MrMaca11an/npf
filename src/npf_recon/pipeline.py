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
from npf_recon.llm_export import write_llm_export
from npf_recon.normalize import format_amount, format_date
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
    skipped: list[str] = []
    # Счётчики по сторонам: {сторона: [файлов, записей]}
    side_stats: dict[str, list[int]] = {"ПУ": [0, 0], "БУ": [0, 0]}
    for raw in raw_docs:
        rule = match_rule(raw.base_name, raw.side)
        if rule is None:
            # Пропущенные файлы перечисляются в итоговой сводке — здесь только debug.
            logger.debug(
                "Файл не соответствует ни одному правилу [%s]: %s — пропускаем",
                raw.side, raw.path.name
            )
            skipped.append(raw.path.name)
            continue
        parser = get_parser(rule)
        records = parser.parse(raw, rule)
        all_records.extend(records)
        stat = side_stats.setdefault(raw.side, [0, 0])
        stat[0] += 1
        stat[1] += len(records)

    if not all_records:
        print("\n⚠  Не удалось извлечь ни одной записи из файлов.")
        return 1

    # ------------------------------------------------------------------
    # 4-5. Агрегация и сверка
    # ------------------------------------------------------------------
    aggregated = aggregate(all_records)
    recon_rows = reconcile(aggregated, eps=config.eps)

    # ------------------------------------------------------------------
    # 6-7. Отчёт + обезличенный экспорт для внешней LLM
    # ------------------------------------------------------------------
    write_report(
        rows=recon_rows,
        output_path=config.output_file,
        period_label=config.period_label,
        eps=config.eps,
    )
    summary_path = config.resolved_summary_file()
    prompt_path = config.resolved_prompt_file()
    write_llm_export(
        rows=recon_rows,
        period_label=config.period_label,
        summary_path=summary_path,
        prompt_path=prompt_path,
        eps=config.eps,
    )

    _print_summary(config, side_stats, skipped, recon_rows)
    return 0


def _signed_amount(x: float) -> str:
    """Форматирует расхождение со знаком: «+15 000,00» / «−5 000,00»."""
    sign = "+" if x > 0 else "−"
    return f"{sign}{format_amount(abs(x))}"


def _print_summary(
    config: Config,
    side_stats: dict[str, list[int]],
    skipped: list[str],
    recon_rows: list,
) -> None:
    """Печатает компактную структурированную сводку результата."""
    line = "─" * 60

    def _has_discrep(r) -> bool:
        total_diff = r.diff_total is not None and abs(r.diff_total) > config.eps
        return total_diff or bool(r.diff_dates)

    with_discrep = [r for r in recon_rows if _has_discrep(r)]

    print(f"\n{line}")
    print(f"  СВЕРКА БУ ↔ ПУ — {config.period_label}")
    print(line)

    # Источники
    print("  Источники:")
    for side in ("ПУ", "БУ"):
        files, recs = side_stats.get(side, [0, 0])
        print(f"    {side}   файлов: {files:>2}   записей: {recs:>4}")
    if skipped:
        print(f"    Пропущено (нет правила): {len(skipped)}")
        for name in skipped:
            print(f"        · {name}")

    # Результат сверки
    print("\n  Результат:")
    if with_discrep:
        print(f"    Расхождения по {len(with_discrep)} показателям:")
        name_w = max(len(r.indicator) for r in with_discrep)
        for r in with_discrep:
            dates_str = ", ".join(format_date(d) for d in sorted(r.diff_dates))
            if r.diff_total is not None and abs(r.diff_total) > config.eps:
                total_str = f"БУ−ПУ = {_signed_amount(r.diff_total):>14}"
            else:
                total_str = "итоги равны, отличия по дням"
            print(f"    • {r.indicator:<{name_w}}  {total_str}   {dates_str}")
    else:
        print("    Расхождений не обнаружено.")

    # Файлы результата
    print("\n  Файлы:")
    print(f"    Отчёт:    {config.output_file}")
    print(f"    Саммари:  {config.resolved_summary_file()}")
    print(f"    Промт:    {config.resolved_prompt_file()}")
    print(line)
