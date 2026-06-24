"""
Генерация синтетических тестовых данных для системы сверки БУ-ПУ.

Создаёт файлы в data/ПУ и data/БУ для периода Февраль 2026.

Введены намеренные расхождения:
  1. «Страховой резерв» на 17.02.2026: БУ на 15 000 больше ПУ.
  2. «Пенсионные взносы ФЛ» на 28.02.2026: БУ (JSON) на 5 000 больше ПУ (Excel).
"""
from __future__ import annotations

import json
import os
import sys
from datetime import date
from pathlib import Path

import openpyxl
from openpyxl.styles import Font

# Добавляем корень репозитория в PYTHONPATH
_REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_REPO_ROOT))
sys.path.insert(0, str(_REPO_ROOT / "src"))

from npf_recon.normalize import format_amount


def _get_data_dir() -> Path:
    """Возвращает базовую директорию данных (с поддержкой NPF_DATA_DIR)."""
    env = os.environ.get("NPF_DATA_DIR")
    if env:
        return Path(env)
    return _REPO_ROOT / "data"


# ------------------------------------------------------------------
# Даты периода
# ------------------------------------------------------------------
DATES = [
    date(2026, 2, 1),
    date(2026, 2, 3),
    date(2026, 2, 17),
    date(2026, 2, 28),
]

# Базовые суммы по датам (в рублях) — используются во всех показателях
# с разными коэффициентами, чтобы данные выглядели реалистично.
BASE_AMOUNTS = {
    date(2026, 2, 1):  100_000.00,
    date(2026, 2, 3):   75_500.50,
    date(2026, 2, 17): 200_000.00,
    date(2026, 2, 28): 312_455.75,
}


def _fmt_date_str(d: date) -> str:
    """Форматирует дату как строку DD.MM.YYYY (для ячеек Excel)."""
    return d.strftime("%d.%m.%Y")


# ------------------------------------------------------------------
# Вспомогательная функция: запись Excel-файла «generic» (A=дата, H или C или I = сумма)
# ------------------------------------------------------------------
def write_generic_excel(
    path: Path,
    indicator: str,
    amounts: dict[date, float],
    amount_col_idx: int = 7,  # 0-based: H=7, C=2, I=8
    use_string_dates: bool = False,
    use_space_amounts: bool = False,
) -> None:
    """
    Записывает Excel-файл «generic» формата (без заголовка, дата в A, сумма в нужной колонке).

    :param path:             путь к файлу.
    :param indicator:        название показателя (для шапки).
    :param amounts:          словарь {дата: сумма}.
    :param amount_col_idx:   0-based индекс колонки суммы (H=7, C=2, I=8).
    :param use_string_dates: если True — даты записываются как строки «DD.MM.YYYY»,
                             иначе — как Python date (Excel datetime).
    :param use_space_amounts: если True — суммы записываются в формате с пробелами и запятой.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Данные"

    # Строка 1: заголовок (шапка, будет пропущена парсером, т.к. не дата)
    ws.cell(row=1, column=1, value="Дата").font = Font(bold=True)
    ws.cell(row=1, column=amount_col_idx + 1, value=indicator).font = Font(bold=True)

    # Строка 2: пустая (имитирует реальный документ)
    ws.cell(row=2, column=1, value="Отчёт за февраль 2026")

    # Данные начиная со строки 3
    for row_idx, (d, amount) in enumerate(sorted(amounts.items()), start=3):
        if use_string_dates:
            ws.cell(row=row_idx, column=1, value=_fmt_date_str(d))
        else:
            # Записываем как Python date — openpyxl конвертирует в Excel datetime
            ws.cell(row=row_idx, column=1, value=d)

        if use_space_amounts:
            # Формат с пробелами и запятой: «1 234,56»
            ws.cell(row=row_idx, column=amount_col_idx + 1,
                    value=format_amount(amount))
        else:
            ws.cell(row=row_idx, column=amount_col_idx + 1, value=amount)

    # Строка-итог в конце (будет пропущена парсером, т.к. текст в колонке даты)
    last_row = len(amounts) + 3
    ws.cell(row=last_row, column=1, value="ИТОГО")
    ws.cell(row=last_row, column=amount_col_idx + 1,
            value=sum(amounts.values()))

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    try:
        rel = path.relative_to(_REPO_ROOT)
    except ValueError:
        rel = path
    print(f"  Создан: {rel}")


# ------------------------------------------------------------------
# Файл «Выплата»
# ------------------------------------------------------------------
def write_vyplata_excel(path: Path, vyplata_data: dict[str, dict[date, float]]) -> None:
    """
    Записывает Excel-файл выплат с заголовком.

    :param path:         путь к файлу.
    :param vyplata_data: {вид_выплаты: {дата: сумма}}.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Выплаты"

    # Строка заголовков (имена колонок, как ожидает парсер)
    headers = ["№", "Дата операции", "ВидВыплат", "Договор", "ФИО", "Сумма", "Примечание"]
    for col_idx, h in enumerate(headers, start=1):
        ws.cell(row=1, column=col_idx, value=h).font = Font(bold=True)

    row_idx = 2
    for type_name, amounts in vyplata_data.items():
        for d, amount in sorted(amounts.items()):
            ws.cell(row=row_idx, column=1, value=row_idx - 1)       # №
            # Чередуем: часть дат как строка, часть как datetime
            if row_idx % 2 == 0:
                ws.cell(row=row_idx, column=2, value=_fmt_date_str(d))
            else:
                ws.cell(row=row_idx, column=2, value=d)
            ws.cell(row=row_idx, column=3, value=type_name)
            ws.cell(row=row_idx, column=4, value=f"ДПО-{1000 + row_idx}")
            ws.cell(row=row_idx, column=5, value="Иванов И.И.")
            ws.cell(row=row_idx, column=6, value=amount)
            ws.cell(row=row_idx, column=7, value="")
            row_idx += 1

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    try:
        rel = path.relative_to(_REPO_ROOT)
    except ValueError:
        rel = path
    print(f"  Создан: {rel}")


# ------------------------------------------------------------------
# JSON НПО
# ------------------------------------------------------------------
def write_npo_json(
    path: Path,
    component_amounts: dict[str, dict[date, float]],
) -> None:
    """
    Записывает JSON-файл НПО.

    :param path:              путь к файлу.
    :param component_amounts: {имя_компонента: {дата: сумма}}.
    """
    # Собираем структуру: один item на дату
    items = []
    for d in DATES:
        components = []
        for comp_name, daily in component_amounts.items():
            amount = daily.get(d, 0.0)
            if amount:
                components.append({
                    "account": "76.01",
                    "name": comp_name,
                    # Чередуем форматы чисел для проверки нормализации
                    "amount": f"{amount:,.2f}" if d.day % 2 == 0 else str(amount)
                })

        if components:
            items.append({
                # Чередуем форматы дат
                "date": f"{d.isoformat()}T00:00:00" if d.day % 2 == 1 else d.isoformat(),
                "amount": str(sum(c["amount"] if isinstance(c["amount"], (int, float))
                                  else float(str(c["amount"]).replace(",", ""))
                                  for c in components)),
                "components": components,
            })

    data = {
        "report": {
            "startDate": "2026-02-01T00:00:00",
            "endDate": "2026-02-28T00:00:00",
            "items": items,
        }
    }

    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    try:
        rel = path.relative_to(_REPO_ROOT)
    except ValueError:
        rel = path
    print(f"  Создан: {rel}")


# ------------------------------------------------------------------
# Главная функция генерации
# ------------------------------------------------------------------
def generate(data_dir: Path | None = None) -> None:
    """
    Генерирует все синтетические файлы данных.

    :param data_dir: базовая директория данных (по умолчанию — из настроек).
    """
    if data_dir is None:
        data_dir = _get_data_dir()

    pu_dir = data_dir / "ПУ"
    bu_dir = data_dir / "БУ"
    pu_dir.mkdir(parents=True, exist_ok=True)
    bu_dir.mkdir(parents=True, exist_ok=True)

    print(f"\nГенерация тестовых данных в: {data_dir}")
    print(f"  ПУ: {pu_dir}")
    print(f"  БУ: {bu_dir}")
    print()

    # ==================================================================
    # ПУ: Пенсионные взносы ФЛ (колонка A=дата, H=сумма, индекс 7)
    # Намеренное расхождение: на 28.02.2026 ПУ имеет 307 455.75 (меньше БУ на 5 000)
    # ==================================================================
    pu_pv_fl_amounts = {
        date(2026, 2, 1):  100_000.00,
        date(2026, 2, 3):   75_500.50,
        date(2026, 2, 17): 200_000.00,
        date(2026, 2, 28): 307_455.75,   # ← ПУ: 307 455.75 (БУ будет 312 455.75, разница 5 000)
    }
    write_generic_excel(
        pu_dir / "Пенсионные взносы ФЛ_Февраль2026.xlsx",
        "Пенсионные взносы ФЛ",
        pu_pv_fl_amounts,
        amount_col_idx=7,
        use_string_dates=True,      # Тест нормализации строковых дат
        use_space_amounts=False,
    )

    # ==================================================================
    # ПУ: Пенсионные взносы ЮЛ
    # ==================================================================
    pu_pv_ul_amounts = {
        date(2026, 2, 1):  55_000.00,
        date(2026, 2, 3):  33_200.00,
        date(2026, 2, 17): 88_500.00,
        date(2026, 2, 28): 120_000.00,
    }
    write_generic_excel(
        pu_dir / "Пенсионные взносы ЮЛ_Февраль2026.xlsx",
        "Пенсионные взносы ЮЛ",
        pu_pv_ul_amounts,
        amount_col_idx=7,
        use_string_dates=False,
        use_space_amounts=False,
    )

    # ==================================================================
    # ПУ: Целевые взносы ФЛ
    # ==================================================================
    pu_cv_fl_amounts = {
        date(2026, 2, 1):  12_000.00,
        date(2026, 2, 3):   8_500.00,
        date(2026, 2, 17): 25_000.00,
        date(2026, 2, 28): 31_000.00,
    }
    write_generic_excel(
        pu_dir / "Целевые взносы ФЛ_Февраль2026.xlsx",
        "Целевые взносы ФЛ",
        pu_cv_fl_amounts,
        amount_col_idx=7,
        use_string_dates=True,
        use_space_amounts=True,  # Тест нормализации строковых сумм
    )

    # ==================================================================
    # ПУ: Целевые взносы ЮЛ (сумма в колонке C, индекс 2)
    # ==================================================================
    pu_cv_ul_amounts = {
        date(2026, 2, 1):  45_000.00,
        date(2026, 2, 3):  29_000.00,
        date(2026, 2, 17): 67_500.00,
        date(2026, 2, 28): 95_000.00,
    }
    write_generic_excel(
        pu_dir / "Целевые взносы ЮЛ_Февраль2026.xlsx",
        "Целевые взносы ЮЛ",
        pu_cv_ul_amounts,
        amount_col_idx=2,  # C=2
        use_string_dates=False,
        use_space_amounts=False,
    )

    # ==================================================================
    # ПУ: Страховой резерв (сумма в колонке I, индекс 8)
    # Намеренное расхождение: на 17.02.2026 ПУ имеет 85 000 (БУ будет 100 000, разница 15 000)
    # ==================================================================
    pu_sr_amounts = {
        date(2026, 2, 1):  50_000.00,
        date(2026, 2, 3):  30_000.00,
        date(2026, 2, 17): 85_000.00,   # ← ПУ: 85 000 (БУ будет 100 000, разница 15 000)
        date(2026, 2, 28): 60_000.00,
    }
    write_generic_excel(
        pu_dir / "Страховой резерв_Февраль2026.xlsx",
        "Страховой резерв",
        pu_sr_amounts,
        amount_col_idx=8,   # I=8
        use_string_dates=True,
        use_space_amounts=False,
    )

    # ==================================================================
    # ПУ: Выплата (файл с заголовком, ВидВыплат → показатель)
    # ==================================================================
    vyplata_data = {
        "Негосударственная пенсия": {
            date(2026, 2, 1):  8_500.00,
            date(2026, 2, 3):  7_200.00,
            date(2026, 2, 17): 9_100.00,
            date(2026, 2, 28): 8_800.00,
        },
        "Выкупная сумма (Расторжение)": {
            date(2026, 2, 3):  15_000.00,
            date(2026, 2, 17): 22_500.00,
        },
        "Выкупная сумма (Наследникам)": {
            date(2026, 2, 17): 5_000.00,
            date(2026, 2, 28): 3_300.00,
        },
        # Строка с видом выплаты, не входящим в маппинг (должна быть пропущена)
        "Прочие выплаты": {
            date(2026, 2, 28): 999.99,
        },
    }
    write_vyplata_excel(
        pu_dir / "Выплата_Февраль2026.xlsx",
        vyplata_data,
    )

    print()

    # ==================================================================
    # БУ: Страховой резерв (397.03, сумма в колонке I, индекс 8)
    # Намеренное расхождение: на 17.02.2026 БУ имеет 100 000 (ПУ 85 000, разница 15 000)
    # ==================================================================
    bu_sr_amounts = {
        date(2026, 2, 1):  50_000.00,
        date(2026, 2, 3):  30_000.00,
        date(2026, 2, 17): 100_000.00,  # ← БУ: 100 000 (ПУ 85 000, расхождение +15 000)
        date(2026, 2, 28): 60_000.00,
    }
    write_generic_excel(
        bu_dir / "397.03 Страховой резерв Февраль_2026.xlsx",
        "Страховой резерв",
        bu_sr_amounts,
        amount_col_idx=8,
        use_string_dates=False,
        use_space_amounts=True,  # Суммы в формате «100 000,00» для теста нормализации
    )

    # ==================================================================
    # БУ: JSON НПО (Пенсионные взносы ФЛ, ЮЛ, Целевые ФЛ, ЮЛ)
    # Намеренное расхождение: Пенсионные взносы ФЛ на 28.02.2026 → 312 455.75
    #   (ПУ имеет 307 455.75, разница 5 000)
    # ==================================================================
    npo_components = {
        "Пенсионные взносы ФЛ по договорам НПО": {
            date(2026, 2, 1):  100_000.00,
            date(2026, 2, 3):   75_500.50,
            date(2026, 2, 17): 200_000.00,
            date(2026, 2, 28): 312_455.75,   # ← БУ: 312 455.75 (ПУ 307 455.75, разница 5 000)
        },
        "Пенсионные взносы ЮЛ по договорам НПО": {
            date(2026, 2, 1):  55_000.00,
            date(2026, 2, 3):  33_200.00,
            date(2026, 2, 17): 88_500.00,
            date(2026, 2, 28): 120_000.00,
        },
        "Целевые взносы ФЛ": {
            date(2026, 2, 1):  12_000.00,
            date(2026, 2, 3):   8_500.00,
            date(2026, 2, 17): 25_000.00,
            date(2026, 2, 28): 31_000.00,
        },
        "Целевые взносы ЮЛ": {
            date(2026, 2, 1):  45_000.00,
            date(2026, 2, 3):  29_000.00,
            date(2026, 2, 17): 67_500.00,
            date(2026, 2, 28): 95_000.00,
        },
    }
    write_npo_json(
        bu_dir / "НПО_Февраль2026.json",
        npo_components,
    )

    print()
    print("=" * 60)
    print("  Тестовые данные сгенерированы.")
    print()
    print("  Введённые расхождения:")
    print("    1. Страховой резерв на 17.02.2026:")
    print("       БУ = 100 000,00 / ПУ = 85 000,00 → разница 15 000,00")
    print("    2. Пенсионные взносы ФЛ на 28.02.2026:")
    print("       БУ = 312 455,75 / ПУ = 307 455,75 → разница 5 000,00")
    print("=" * 60)


if __name__ == "__main__":
    generate()
