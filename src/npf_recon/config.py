"""
Загрузка конфигурации из config/settings.yaml.

Поддерживаются переменные окружения:
  NPF_DATA_DIR — переопределяет base_dir (путь к папке с данными).
  NPF_PU_DIR   — прямой путь к папке ПУ (переопределяет base_dir + pu_subdir).
  NPF_BU_DIR   — прямой путь к папке БУ (переопределяет base_dir + bu_subdir).
"""
from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

import yaml

# Корень репозитория — родитель папки src/
_REPO_ROOT = Path(__file__).resolve().parent.parent.parent
_SETTINGS_FILE = _REPO_ROOT / "config" / "settings.yaml"


@dataclass
class Config:
    """Конфигурация приложения, разрешённая до абсолютных путей."""

    pu_dir: Path          # Абсолютный путь к папке ПУ
    bu_dir: Path          # Абсолютный путь к папке БУ
    output_file: Path     # Абсолютный путь к итоговому файлу отчёта
    period_label: str     # Метка периода (напр. «Февраль 2026»)
    eps: float            # Допуск для сравнения сумм
    # Обезличенные артефакты для внешней LLM (если None — берутся рядом с отчётом)
    summary_file: Path | None = None   # JSON-саммари (без перс. данных)
    prompt_file: Path | None = None    # Готовый промт для LLM

    def resolved_summary_file(self) -> Path:
        """Путь к JSON-саммари (по умолчанию рядом с отчётом)."""
        return self.summary_file or self.output_file.with_name("Саммари_БУ_ПУ.json")

    def resolved_prompt_file(self) -> Path:
        """Путь к файлу промта (по умолчанию рядом с отчётом)."""
        return self.prompt_file or self.output_file.with_name("Промт_для_LLM.txt")


def load_config(settings_path: Path | None = None) -> Config:
    """
    Загружает конфигурацию из YAML-файла.

    :param settings_path: путь к файлу настроек; если None — используется
                          config/settings.yaml в корне репозитория.
    :return: объект Config с абсолютными путями.
    """
    path = settings_path or _SETTINGS_FILE
    with open(path, encoding="utf-8") as f:
        raw = yaml.safe_load(f)

    # Базовая директория данных — можно переопределить через окружение
    env_data_dir = os.environ.get("NPF_DATA_DIR")
    if env_data_dir:
        base = Path(env_data_dir).resolve()
    else:
        base_raw = raw.get("base_dir", "data")
        base = (Path(base_raw) if Path(base_raw).is_absolute()
                else _REPO_ROOT / base_raw)

    pu_subdir = raw.get("pu_subdir", "ПУ")
    bu_subdir = raw.get("bu_subdir", "БУ")

    # Прямые переменные окружения для продуктивных путей Windows
    pu_dir = Path(os.environ["NPF_PU_DIR"]).resolve() if "NPF_PU_DIR" in os.environ \
        else base / pu_subdir
    bu_dir = Path(os.environ["NPF_BU_DIR"]).resolve() if "NPF_BU_DIR" in os.environ \
        else base / bu_subdir

    output_raw = raw.get("output_file", "output/Сверка_БУ_ПУ.xlsx")
    output_file = (Path(output_raw) if Path(output_raw).is_absolute()
                   else _REPO_ROOT / output_raw)

    return Config(
        pu_dir=pu_dir,
        bu_dir=bu_dir,
        output_file=output_file,
        period_label=raw.get("period_label", ""),
        eps=float(raw.get("eps", 0.005)),
    )
