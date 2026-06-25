"""
Точка входа CLI для системы сверки БУ-ПУ.

Запуск:
  python -m npf_recon                  # обычный прогон по config/settings.yaml
  python -m npf_recon --generate-sample # сначала создать демо-данные, затем сверка
  python -m npf_recon --data-dir /path --period "Март 2026" --eps 0.01
  python src/npf_recon/main.py -v
"""
from __future__ import annotations

import argparse
import logging
import os
import sys
from pathlib import Path

from npf_recon.config import load_config
from npf_recon.pipeline import run


def setup_logging(verbose: bool = False) -> None:
    """Настраивает форматирование логов."""
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )
    # Подавляем шумные логи сторонних библиотек
    for noisy in ("openpyxl", "PIL"):
        logging.getLogger(noisy).setLevel(logging.WARNING)


def _parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="npf_recon",
        description="Автоматическая сверка БУ ↔ ПУ из Excel/JSON в итоговый Excel-отчёт.",
    )
    parser.add_argument("-v", "--verbose", action="store_true",
                        help="Подробный лог (DEBUG).")
    parser.add_argument("--menu", action="store_true",
                        help="Интерактивное консольное меню (удобно для демонстрации).")
    parser.add_argument("--diagnose", action="store_true",
                        help="Диагностика парсинга: показать разбор каждого файла.")
    parser.add_argument("--generate-sample", action="store_true",
                        help="Сгенерировать синтетические демо-данные перед сверкой.")
    parser.add_argument("--data-dir", metavar="PATH",
                        help="Базовая директория данных (переопределяет NPF_DATA_DIR).")
    parser.add_argument("--period", metavar="LABEL",
                        help="Метка периода для заголовка отчёта (напр. «Март 2026»).")
    parser.add_argument("--eps", type=float, metavar="RUB",
                        help="Допуск сравнения сумм в рублях (по умолчанию из настроек).")
    parser.add_argument("--output", metavar="PATH",
                        help="Путь к итоговому файлу отчёта.")
    return parser.parse_args(argv)


def _generate_sample(data_dir: Path | None) -> None:
    """Импортирует и запускает генератор демо-данных из scripts/."""
    scripts_dir = Path(__file__).resolve().parent.parent.parent / "scripts"
    if str(scripts_dir) not in sys.path:
        sys.path.insert(0, str(scripts_dir))
    import generate_sample_data  # noqa: WPS433 (поздний импорт намеренно)
    generate_sample_data.generate(data_dir=data_dir)


def main(argv: list[str] | None = None) -> int:
    """
    Главная функция CLI.

    :return: код завершения (0 — успех, 1 — ошибка).
    """
    args = _parse_args(argv)
    setup_logging(args.verbose)

    # --data-dir пробрасывается в конфиг через переменную окружения
    if args.data_dir:
        os.environ["NPF_DATA_DIR"] = str(Path(args.data_dir).resolve())

    # Интерактивное меню — отдельный режим
    if args.menu:
        from npf_recon.menu import run_menu
        return run_menu()

    print("=" * 60)
    print("  Система сверки БУ-ПУ (НПФ)")
    print("=" * 60)

    if args.generate_sample:
        data_dir = Path(args.data_dir).resolve() if args.data_dir else None
        _generate_sample(data_dir)

    config = load_config()

    # Переопределения из CLI
    if args.period:
        config.period_label = args.period
    if args.eps is not None:
        config.eps = args.eps
    if args.output:
        config.output_file = Path(args.output).resolve()

    print(f"Директория ПУ: {config.pu_dir}")
    print(f"Директория БУ: {config.bu_dir}")

    # Режим диагностики парсинга
    if args.diagnose:
        from npf_recon.diagnostics import run_diagnostics
        return 0 if run_diagnostics(config) else 1

    return run(config)


if __name__ == "__main__":
    sys.exit(main())
