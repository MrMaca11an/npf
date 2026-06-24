"""
Генерация синтетических тестовых данных для системы сверки БУ-ПУ.

Создаёт файлы в data/ПУ и data/БУ для периода Февраль 2026, насыщенные
разнообразными форматами дат и чисел, несколькими операциями в один день
(для проверки суммирования по дням) и строками, которые должны игнорироваться.

Введены ТРИ намеренных расхождения:
  1. «Страховой резерв» на 17.02.2026: БУ на 15 000,00 больше ПУ (расхождение по итогу).
  2. «Пенсионные взносы ФЛ» на 28.02.2026: БУ на 5 000,00 больше ПУ (расхождение по итогу).
  3. «Целевые взносы ЮЛ»: месячные итоги БУ и ПУ РАВНЫ, но есть встречные
     отклонения по дням — 10.02 (+7 000) и 24.02 (−7 000). Демонстрирует
     обнаружение дневных расхождений даже при совпадении итогов.

Дополнительно «Целевые взносы ФЛ» содержат отрицательную корректировку
(возврат) на 14.02.2026 в обеих сторонах — проверка обработки отрицательных
значений (скобки / завершающий минус) без возникновения расхождения.
"""
from __future__ import annotations

import json
import os
import sys
from datetime import date
from itertools import cycle
from pathlib import Path

import openpyxl
from openpyxl.styles import Font

# Добавляем корень репозитория в PYTHONPATH
_REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_REPO_ROOT))
sys.path.insert(0, str(_REPO_ROOT / "src"))


# ------------------------------------------------------------------
# Даты периода (9 рабочих дат февраля 2026)
# ------------------------------------------------------------------
DATES = [
    date(2026, 2, 1),
    date(2026, 2, 3),
    date(2026, 2, 5),
    date(2026, 2, 10),
    date(2026, 2, 14),
    date(2026, 2, 17),
    date(2026, 2, 20),
    date(2026, 2, 24),
    date(2026, 2, 28),
]


def _get_data_dir() -> Path:
    """Возвращает базовую директорию данных (с поддержкой NPF_DATA_DIR)."""
    env = os.environ.get("NPF_DATA_DIR")
    if env:
        return Path(env)
    return _REPO_ROOT / "data"


# ==================================================================
# Рендеринг дат и сумм в разнообразных форматах
# ==================================================================

# Неразрывный пробел и узкий неразрывный пробел
_NBSP = " "
_NARROW = " "


def _render_date(d: date, style: str):
    """Возвращает дату в заданном строковом/datetime-представлении."""
    if style == "dd.mm.yyyy":
        return d.strftime("%d.%m.%Y")
    if style == "dd.mm.yy":
        return d.strftime("%d.%m.%y")
    if style == "iso":
        return d.isoformat()
    # "datetime" — отдаём объект date, openpyxl запишет как Excel datetime
    return d


def _render_amount(a: float, style: str):
    """
    Возвращает сумму в заданном представлении.

    plain   — числовое значение (float), отрицательные — обычный минус.
    space   — строка «1 234,56» (обычный пробел), отрицательные — «(1 234,56)».
    nbsp    — то же с неразрывным пробелом (U+00A0).
    narrow  — то же с узким неразрывным пробелом (U+202F).
    """
    if style == "plain":
        return a

    sep = {"space": " ", "nbsp": _NBSP, "narrow": _NARROW}[style]
    raw = f"{abs(a):,.2f}".replace(",", "\x00").replace(".", ",").replace("\x00", sep)
    return f"({raw})" if a < 0 else raw


def _split_rows(daily: dict[date, float], split_dates=()):
    """
    Превращает {дата: сумма} в список (дата, сумма), при необходимости
    разбивая сумму выбранных дат на несколько операций (для проверки
    суммирования нескольких строк в один день).
    """
    rows: list[tuple[date, float]] = []
    for d in sorted(daily):
        amt = daily[d]
        if d in split_dates and abs(amt) > 2:
            part1 = round(amt * 0.4, 2)
            part2 = round(amt - part1, 2)
            rows.append((d, part1))
            rows.append((d, part2))
        else:
            rows.append((d, amt))
    return rows


def _print_created(path: Path) -> None:
    try:
        rel = path.relative_to(_REPO_ROOT)
    except ValueError:
        rel = path
    print(f"  Создан: {rel}")


# ==================================================================
# Писатель Excel «generic» (без заголовка: дата в A, сумма в нужной колонке)
# ==================================================================
def write_generic_excel(
    path: Path,
    indicator: str,
    rows: list[tuple[date, float]],
    amount_col_idx: int,
    date_styles: list[str],
    amount_styles: list[str],
) -> None:
    """
    Записывает «generic» Excel: строка-шапка (пропускается парсером),
    строка-подзаголовок, строки данных, строка ИТОГО.

    Стили дат и сумм циклически чередуются по строкам — чтобы один файл
    содержал смесь форматов и проверял устойчивость нормализации.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Данные"

    # Строка 1: заголовок-шапка (текст в колонке даты → парсер пропустит)
    ws.cell(row=1, column=1, value="Дата").font = Font(bold=True)
    ws.cell(row=1, column=amount_col_idx + 1, value=indicator).font = Font(bold=True)
    # Строка 2: подзаголовок (имитация реального документа)
    ws.cell(row=2, column=1, value=f"Отчёт: {indicator} за февраль 2026")

    ds = cycle(date_styles)
    as_ = cycle(amount_styles)
    row_idx = 3
    for d, amount in rows:
        ws.cell(row=row_idx, column=1, value=_render_date(d, next(ds)))
        ws.cell(row=row_idx, column=amount_col_idx + 1,
                value=_render_amount(amount, next(as_)))
        row_idx += 1

    # Строка-итог (текст в колонке даты → пропускается)
    ws.cell(row=row_idx, column=1, value="ИТОГО")
    ws.cell(row=row_idx, column=amount_col_idx + 1,
            value=round(sum(a for _, a in rows), 2))

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    _print_created(path)


# ==================================================================
# Писатель файла «Выплата» (с заголовком)
# ==================================================================
def write_vyplata_excel(path: Path, vyplata_rows: list[tuple]) -> None:
    """
    vyplata_rows: список (дата, вид_выплаты, сумма).
    Даты чередуют строковый и datetime-форматы.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Выплаты"

    headers = ["№", "Дата операции", "ВидВыплат", "Договор", "ФИО", "Сумма", "Примечание"]
    for col_idx, h in enumerate(headers, start=1):
        ws.cell(row=1, column=col_idx, value=h).font = Font(bold=True)

    ds = cycle(["dd.mm.yyyy", "datetime", "iso"])
    for i, (d, type_name, amount) in enumerate(vyplata_rows):
        r = i + 2
        ws.cell(row=r, column=1, value=i + 1)
        ws.cell(row=r, column=2, value=_render_date(d, next(ds)))
        ws.cell(row=r, column=3, value=type_name)
        ws.cell(row=r, column=4, value=f"ДПО-{1000 + i}")
        ws.cell(row=r, column=5, value="Иванов И.И.")
        ws.cell(row=r, column=6, value=amount)
        ws.cell(row=r, column=7, value="")

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    _print_created(path)


# ==================================================================
# Писатель JSON НПО
# ==================================================================
def write_npo_json(path: Path, component_daily: dict[str, dict[date, float]]) -> None:
    """
    component_daily: {имя_компонента: {дата: сумма}} для стороны БУ.

    Форматы дат и сумм чередуются по позиции. Один неизвестный компонент
    добавляется для проверки пропуска. Отрицательные суммы кодируются
    скобками или завершающим минусом.
    """
    def _amount_str(v: float, idx: int) -> str:
        if v < 0:
            # Чередуем форматы отрицательных: скобки и завершающий минус
            return (f"({abs(v):,.2f})".replace(",", " ")
                    if idx % 2 == 0 else f"{abs(v):.2f}-")
        # Чередуем: американская группировка «100,000.00» и простое «100000.0»
        return f"{v:,.2f}" if idx % 2 == 0 else str(v)

    items = []
    for di, d in enumerate(DATES):
        components = []
        for ci, (comp_name, daily) in enumerate(component_daily.items()):
            amount = daily.get(d)
            if amount is None:
                continue
            components.append({
                "account": f"76.0{(ci % 5) + 1}",
                "name": comp_name,
                "amount": _amount_str(amount, di + ci),
            })

        # Один неизвестный компонент (должен игнорироваться парсером)
        if d == DATES[0]:
            components.append({
                "account": "99.01",
                "name": "Прочие операции НПО",
                "amount": "777.77",
            })

        if components:
            total = round(sum(daily.get(d, 0.0) for daily in component_daily.values()), 2)
            items.append({
                # Чередуем ISO с временем и без
                "date": f"{d.isoformat()}T00:00:00" if di % 2 == 0 else d.isoformat(),
                "amount": str(total),
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
    _print_created(path)


# ==================================================================
# Данные показателей (источник истины для генерации)
# ==================================================================

# Пенсионные взносы ФЛ — БУ и ПУ совпадают, кроме 28.02 (БУ +5 000)
PV_FL = {
    date(2026, 2, 1): 100_000.00, date(2026, 2, 3): 75_500.50,
    date(2026, 2, 5): 60_250.25, date(2026, 2, 10): 88_000.00,
    date(2026, 2, 14): 42_300.10, date(2026, 2, 17): 200_000.00,
    date(2026, 2, 20): 91_000.00, date(2026, 2, 24): 53_200.00,
    date(2026, 2, 28): 307_455.75,
}
PV_FL_BU = {**PV_FL, date(2026, 2, 28): 312_455.75}  # +5 000 → расхождение

# Пенсионные взносы ЮЛ — полностью совпадают
PV_UL = {
    date(2026, 2, 1): 55_000.00, date(2026, 2, 3): 33_200.00,
    date(2026, 2, 5): 27_000.00, date(2026, 2, 10): 41_000.00,
    date(2026, 2, 14): 19_500.00, date(2026, 2, 17): 88_500.00,
    date(2026, 2, 20): 36_000.00, date(2026, 2, 24): 22_750.50,
    date(2026, 2, 28): 120_000.00,
}

# Целевые взносы ФЛ — совпадают; 14.02 отрицательная корректировка (возврат)
CV_FL = {
    date(2026, 2, 1): 12_000.00, date(2026, 2, 3): 8_500.00,
    date(2026, 2, 5): 6_000.00, date(2026, 2, 10): 9_900.00,
    date(2026, 2, 14): -1_500.00, date(2026, 2, 17): 25_000.00,
    date(2026, 2, 20): 7_200.00, date(2026, 2, 24): 4_300.00,
    date(2026, 2, 28): 31_000.00,
}

# Целевые взносы ЮЛ — итоги равны, но дневной сдвиг 7 000 (10.02 ↔ 24.02)
CV_UL_BU = {
    date(2026, 2, 1): 45_000.00, date(2026, 2, 3): 29_000.00,
    date(2026, 2, 5): 18_000.00, date(2026, 2, 10): 30_000.00,
    date(2026, 2, 14): 12_500.00, date(2026, 2, 17): 67_500.00,
    date(2026, 2, 20): 21_000.00, date(2026, 2, 24): 15_000.00,
    date(2026, 2, 28): 95_000.00,
}
CV_UL_PU = {**CV_UL_BU,
            date(2026, 2, 10): 23_000.00,   # на 7 000 меньше БУ
            date(2026, 2, 24): 22_000.00}   # на 7 000 больше БУ (итог равен)

# Страховой резерв — совпадают, кроме 17.02 (БУ +15 000)
SR_PU = {
    date(2026, 2, 1): 50_000.00, date(2026, 2, 3): 30_000.00,
    date(2026, 2, 5): 22_000.00, date(2026, 2, 10): 41_000.00,
    date(2026, 2, 14): 18_000.00, date(2026, 2, 17): 85_000.00,
    date(2026, 2, 20): 27_500.00, date(2026, 2, 24): 16_000.00,
    date(2026, 2, 28): 60_000.00,
}
SR_BU = {**SR_PU, date(2026, 2, 17): 100_000.00}  # +15 000 → расхождение


# ==================================================================
# Главная функция генерации
# ==================================================================
def generate(data_dir: Path | None = None) -> None:
    """Генерирует все синтетические файлы данных."""
    if data_dir is None:
        data_dir = _get_data_dir()

    pu_dir = data_dir / "ПУ"
    bu_dir = data_dir / "БУ"
    pu_dir.mkdir(parents=True, exist_ok=True)
    bu_dir.mkdir(parents=True, exist_ok=True)

    print(f"\nГенерация тестовых данных в: {data_dir}")
    print(f"  ПУ: {pu_dir}")
    print(f"  БУ: {bu_dir}\n")

    # ---------------- ПУ ----------------
    # Пенсионные взносы ФЛ: строковые даты, несколько операций в день, простые суммы
    write_generic_excel(
        pu_dir / "Пенсионные взносы ФЛ_Февраль2026.xlsx",
        "Пенсионные взносы ФЛ",
        _split_rows(PV_FL, split_dates={date(2026, 2, 1), date(2026, 2, 17)}),
        amount_col_idx=7,                                   # H
        date_styles=["dd.mm.yyyy"],
        amount_styles=["plain"],
    )

    # Пенсионные взносы ЮЛ: datetime-даты, суммы с неразрывным пробелом
    write_generic_excel(
        pu_dir / "Пенсионные взносы ЮЛ_Февраль2026.xlsx",
        "Пенсионные взносы ЮЛ",
        _split_rows(PV_UL),
        amount_col_idx=7,                                   # H
        date_styles=["datetime"],
        amount_styles=["nbsp", "plain"],
    )

    # Целевые взносы ФЛ: короткий год DD.MM.YY, узкий пробел, отрицательная корректировка
    write_generic_excel(
        pu_dir / "Целевые взносы ФЛ_Февраль2026.xlsx",
        "Целевые взносы ФЛ",
        _split_rows(CV_FL),
        amount_col_idx=7,                                   # H
        date_styles=["dd.mm.yy", "dd.mm.yyyy"],
        amount_styles=["narrow", "space"],
    )

    # Целевые взносы ЮЛ: сумма в колонке C, datetime-даты, дневной сдвиг (ПУ)
    write_generic_excel(
        pu_dir / "Целевые взносы ЮЛ_Февраль2026.xlsx",
        "Целевые взносы ЮЛ",
        _split_rows(CV_UL_PU, split_dates={date(2026, 2, 28)}),
        amount_col_idx=2,                                   # C
        date_styles=["datetime", "iso"],
        amount_styles=["plain"],
    )

    # Страховой резерв: сумма в колонке I, смесь форматов дат и сумм
    write_generic_excel(
        pu_dir / "Страховой резерв_Февраль2026.xlsx",
        "Страховой резерв",
        _split_rows(SR_PU),
        amount_col_idx=8,                                   # I
        date_styles=["dd.mm.yyyy", "datetime"],
        amount_styles=["space", "plain"],
    )

    # Выплата: несколько операций в день, неизвестный вид (пропуск)
    NP = "Негосударственная пенсия"
    VR = "Выкупная сумма (Расторжение)"
    VN = "Выкупная сумма (Наследникам)"
    vyplata_rows = [
        (date(2026, 2, 1), NP, 8_500.00),
        (date(2026, 2, 1), NP, 2_300.00),       # вторая операция в тот же день
        (date(2026, 2, 3), NP, 7_200.00),
        (date(2026, 2, 5), NP, 6_400.00),
        (date(2026, 2, 10), NP, 8_900.00),
        (date(2026, 2, 14), NP, 7_750.00),
        (date(2026, 2, 17), NP, 9_100.00),
        (date(2026, 2, 20), NP, 8_050.00),
        (date(2026, 2, 24), NP, 7_600.00),
        (date(2026, 2, 28), NP, 8_800.00),
        (date(2026, 2, 3), VR, 15_000.00),
        (date(2026, 2, 14), VR, 18_400.00),
        (date(2026, 2, 17), VR, 22_500.00),
        (date(2026, 2, 28), VR, 12_300.00),
        (date(2026, 2, 17), VN, 5_000.00),
        (date(2026, 2, 24), VN, 6_700.00),
        (date(2026, 2, 28), VN, 3_300.00),
        (date(2026, 2, 28), "Прочие выплаты", 999.99),   # не в маппинге → пропуск
    ]
    write_vyplata_excel(pu_dir / "Выплата_Февраль2026.xlsx", vyplata_rows)

    print()

    # ---------------- БУ ----------------
    # Страховой резерв 397.03: колонка I, суммы с неразрывным пробелом
    write_generic_excel(
        bu_dir / "397.03 Страховой резерв Февраль_2026.xlsx",
        "Страховой резерв",
        _split_rows(SR_BU),
        amount_col_idx=8,                                   # I
        date_styles=["datetime", "dd.mm.yyyy"],
        amount_styles=["nbsp", "space"],
    )

    # JSON НПО: 4 показателя взносов (сторона БУ)
    write_npo_json(
        bu_dir / "НПО_Февраль2026.json",
        {
            "Пенсионные взносы ФЛ по договорам НПО": PV_FL_BU,
            "Пенсионные взносы ЮЛ по договорам НПО": PV_UL,
            "Целевые взносы ФЛ": CV_FL,
            "Целевые взносы ЮЛ": CV_UL_BU,
        },
    )

    print()
    print("=" * 64)
    print("  Тестовые данные сгенерированы.")
    print()
    print("  Намеренные расхождения:")
    print("    1. Страховой резерв, 17.02.2026:")
    print("       БУ = 100 000,00 / ПУ = 85 000,00 → разница +15 000,00")
    print("    2. Пенсионные взносы ФЛ, 28.02.2026:")
    print("       БУ = 312 455,75 / ПУ = 307 455,75 → разница +5 000,00")
    print("    3. Целевые взносы ЮЛ (итоги РАВНЫ, отличия по дням):")
    print("       10.02.2026: БУ−ПУ = +7 000,00; 24.02.2026: БУ−ПУ = −7 000,00")
    print("=" * 64)


if __name__ == "__main__":
    generate()
