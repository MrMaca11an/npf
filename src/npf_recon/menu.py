"""
Интерактивное консольное меню — удобное управление без знания флагов.

Запуск:
  python -m npf_recon --menu
  ./run.sh    (Linux/macOS)
  run.bat     (Windows)
"""
from __future__ import annotations

from pathlib import Path

from npf_recon.config import Config, load_config
from npf_recon.diagnostics import run_diagnostics
from npf_recon.pipeline import run


_MENU = """
╔════════════════════════════════════════════════════════════╗
║            СИСТЕМА СВЕРКИ БУ ↔ ПУ  (НПФ)                    ║
╠════════════════════════════════════════════════════════════╣
║  1. Демо: сгенерировать тестовые данные и выполнить сверку ║
║  2. Сверка по моим файлам (из каталогов ПУ и БУ)           ║
║  3. Диагностика парсинга (проверка распознавания файлов)   ║
║  4. Где разместить входные файлы                           ║
║  5. Выход                                                  ║
╚════════════════════════════════════════════════════════════╝
"""


def _generate_sample(config: Config) -> None:
    """Генерирует тестовые данные в каталоги из конфигурации."""
    import sys
    scripts_dir = Path(__file__).resolve().parent.parent.parent / "scripts"
    if str(scripts_dir) not in sys.path:
        sys.path.insert(0, str(scripts_dir))
    import generate_sample_data
    # Базовая папка — родитель каталога ПУ (ПУ и БУ лежат рядом)
    generate_sample_data.generate(data_dir=config.pu_dir.parent)


def _show_paths(config: Config) -> None:
    print("\nРазместите входные файлы так:")
    print(f"  Папка ПУ: {config.pu_dir}")
    print("    • Выплата (пенсия, ВС, НС).xlsx")
    print("    • Пенсионные взносы ФЛ.xlsx / Пенсионные взносы ЮЛ.xlsx")
    print("    • Целевые взносы ФЛ.xlsx / Целевые взносы ЮЛ.xlsx")
    print("    • Страховой резерв.xlsx")
    print(f"  Папка БУ: {config.bu_dir}")
    print("    • 397.03 Страховой резерв <период>.xlsx")
    print("    • НПО_<период>.json")
    print("\nИмена распознаются по подстроке без учёта регистра — точный")
    print("суффикс с датой/периодом не важен.\n")


def run_menu() -> int:
    """Запускает интерактивное меню. Возвращает код завершения."""
    config = load_config()
    while True:
        print(_MENU)
        try:
            choice = input("Выберите пункт (1–5): ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nВыход.")
            return 0

        if choice == "1":
            _generate_sample(config)
            run(config)
        elif choice == "2":
            run(config)
        elif choice == "3":
            run_diagnostics(config)
        elif choice == "4":
            _show_paths(config)
        elif choice in ("5", "q", "exit", "выход"):
            print("Выход.")
            return 0
        else:
            print("⚠ Неизвестный пункт. Введите число от 1 до 5.")

        try:
            input("\nНажмите Enter, чтобы вернуться в меню…")
        except (EOFError, KeyboardInterrupt):
            return 0
