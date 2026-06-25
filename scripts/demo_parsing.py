"""
Демонстрация парсинга: показывает, что каждый входной файл (Excel ПУ, Excel БУ,
JSON БУ) корректно распознаётся и данные распределяются по показателям.

Запуск:
  python scripts/demo_parsing.py

Эквивалентно `python -m npf_recon --diagnose`, но как отдельный скрипт.
"""
from __future__ import annotations

import sys
from pathlib import Path

_REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_REPO_ROOT))
sys.path.insert(0, str(_REPO_ROOT / "src"))

from npf_recon.config import load_config
from npf_recon.diagnostics import run_diagnostics

if __name__ == "__main__":
    ok = run_diagnostics(load_config())
    sys.exit(0 if ok else 1)
