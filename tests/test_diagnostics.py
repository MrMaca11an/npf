"""
Smoke-тест диагностики парсинга: на сгенерированных данных все ожидаемые
источники должны распознаваться (run_diagnostics → True).
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest


def _generate_in(data_dir: Path) -> None:
    scripts_dir = Path(__file__).resolve().parent.parent / "scripts"
    sys.path.insert(0, str(scripts_dir))
    import generate_sample_data
    generate_sample_data.generate(data_dir=data_dir)


def test_diagnostics_all_sources_ok(tmp_path: Path, capsys):
    from npf_recon.config import Config
    from npf_recon.diagnostics import run_diagnostics

    _generate_in(tmp_path)
    config = Config(
        pu_dir=tmp_path / "ПУ",
        bu_dir=tmp_path / "БУ",
        output_file=tmp_path / "out.xlsx",
        period_label="Февраль 2026",
        eps=0.005,
    )
    ok = run_diagnostics(config)
    assert ok is True

    out = capsys.readouterr().out
    # Нерелевантный файл отмечен как пропущенный
    assert "нет правила" in out
    # Покрытие основных показателей присутствует
    assert "Пенсионные взносы ФЛ" in out
    assert "Страховой резерв" in out


def test_diagnostics_empty_dirs(tmp_path: Path):
    from npf_recon.config import Config
    from npf_recon.diagnostics import run_diagnostics

    (tmp_path / "ПУ").mkdir()
    (tmp_path / "БУ").mkdir()
    config = Config(
        pu_dir=tmp_path / "ПУ",
        bu_dir=tmp_path / "БУ",
        output_file=tmp_path / "out.xlsx",
        period_label="Февраль 2026",
        eps=0.005,
    )
    assert run_diagnostics(config) is False
