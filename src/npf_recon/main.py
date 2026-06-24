"""
Точка входа CLI для системы сверки БУ-ПУ.

Запуск:
  python -m npf_recon
  python src/npf_recon/main.py
"""
from __future__ import annotations

import logging
import sys

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


def main() -> int:
    """
    Главная функция CLI.

    :return: код завершения (0 — успех, 1 — ошибка).
    """
    verbose = "--verbose" in sys.argv or "-v" in sys.argv
    setup_logging(verbose)

    print("=" * 60)
    print("  Система сверки БУ-ПУ (НПФ)")
    print("=" * 60)

    config = load_config()
    print(f"Директория ПУ: {config.pu_dir}")
    print(f"Директория БУ: {config.bu_dir}")

    return run(config)


if __name__ == "__main__":
    sys.exit(main())
