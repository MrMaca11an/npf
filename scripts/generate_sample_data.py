"""
Генерация синтетических тестовых данных для системы сверки БУ-ПУ.

Создаёт файлы в data/ПУ и data/БУ за период Февраль 2026 (11 дат), насыщенные
разнообразными форматами дат и чисел, несколькими операциями в один день и
«мусорными» строками/файлами, которые система обязана корректно игнорировать.

Демонстрируемые форматы дат:    DD.MM.YYYY, DD.MM.YY, ISO (YYYY-MM-DD),
                                ISO с временем, Excel datetime.
Демонстрируемые форматы сумм:   1234.56; «1 234,56» (обычный/неразрывный/узкий
                                пробел); скобки «(1 500,00)»; завершающий минус
                                «1500.00-»; американская группировка «100,000.00».

Намеренные расхождения (ТРИ):
  1. «Страховой резерв» 17.02.2026: БУ на 15 000 больше ПУ (по итогу).
  2. «Пенсионные взносы ФЛ» 28.02.2026: БУ на 5 000 больше ПУ (по итогу).
  3. «Целевые взносы ЮЛ»: месячные итоги РАВНЫ, но встречные дневные отклонения
     10.02 (+7 000) и 24.02 (−7 000) — обнаружение расхождений на уровне дней.

Дополнительно «Целевые взносы ФЛ» содержат отрицательную корректировку (возврат)
14.02.2026 в обеих сторонах — проверка обработки отрицательных значений без
возникновения расхождения.
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

_REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_REPO_ROOT))
sys.path.insert(0, str(_REPO_ROOT / "src"))


# ------------------------------------------------------------------
# 11 рабочих дат февраля 2026
# ------------------------------------------------------------------
DATES = [
    date(2026, 2, d) for d in (1, 3, 5, 7, 10, 12, 14, 17, 20, 24, 28)
]


def _mk(values: list[float]) -> dict[date, float]:
    """Сопоставляет список значений датам DATES."""
    assert len(values) == len(DATES), "ожидается 11 значений"
    return dict(zip(DATES, values))


def _get_data_dir() -> Path:
    env = os.environ.get("NPF_DATA_DIR")
    return Path(env) if env else _REPO_ROOT / "data"


# ==================================================================
# Рендеринг дат и сумм
# ==================================================================
_NBSP = " "      # U+00A0
_NARROW = " "    # U+202F


def _render_date(d: date, style: str):
    if style == "dd.mm.yyyy":
        return d.strftime("%d.%m.%Y")
    if style == "dd.mm.yy":
        return d.strftime("%d.%m.%y")
    if style == "iso":
        return d.isoformat()
    if style == "iso_dt":
        return f"{d.isoformat()}T00:00:00"
    return d  # datetime


def _render_amount(a: float, style: str):
    if style == "plain":
        return a
    sep = {"space": " ", "nbsp": _NBSP, "narrow": _NARROW}[style]
    raw = f"{abs(a):,.2f}".replace(",", "\x00").replace(".", ",").replace("\x00", sep)
    return f"({raw})" if a < 0 else raw


def _split_rows(daily: dict[date, float], split_dates=()):
    """{дата: сумма} → [(дата, сумма)], дробя выбранные дни на 2 операции."""
    rows: list[tuple[date, float]] = []
    for d in sorted(daily):
        amt = daily[d]
        if d in split_dates and abs(amt) > 2:
            part1 = round(amt * 0.4, 2)
            rows.append((d, part1))
            rows.append((d, round(amt - part1, 2)))
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
# Писатель Excel «generic» (без заголовка)
# ==================================================================
def write_generic_excel(
    path: Path,
    indicator: str,
    rows: list[tuple[date, float]],
    amount_col_idx: int,
    date_styles: list[str],
    amount_styles: list[str],
    noise: bool = True,
) -> None:
    """
    Пишет «generic» Excel: шапка, подзаголовок, данные (со «мусором» в середине),
    строка ИТОГО. Все «мусорные» строки система обязана пропустить.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Данные"

    ws.cell(row=1, column=1, value="Дата").font = Font(bold=True)
    ws.cell(row=1, column=amount_col_idx + 1, value=indicator).font = Font(bold=True)
    ws.cell(row=2, column=1, value=f"Отчёт: {indicator} за февраль 2026")

    ds = cycle(date_styles)
    as_ = cycle(amount_styles)
    r = 3
    n = len(rows)
    for i, (d, amount) in enumerate(rows):
        ws.cell(row=r, column=1, value=_render_date(d, next(ds)))
        ws.cell(row=r, column=amount_col_idx + 1, value=_render_amount(amount, next(as_)))
        r += 1
        # Вставляем «мусор» примерно в середине (должен быть пропущен)
        if noise and i == n // 2:
            ws.cell(row=r, column=1, value=None)                      # пустая строка
            r += 1
            ws.cell(row=r, column=1, value="— промежуточный раздел —")  # текст вместо даты
            ws.cell(row=r, column=amount_col_idx + 1, value=12345)
            r += 1
            ws.cell(row=r, column=1, value=_render_date(d, "dd.mm.yyyy"))  # дата без суммы
            ws.cell(row=r, column=amount_col_idx + 1, value=None)
            r += 1

    ws.cell(row=r, column=1, value="ИТОГО")
    ws.cell(row=r, column=amount_col_idx + 1, value=round(sum(a for _, a in rows), 2))

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    _print_created(path)


# ==================================================================
# Писатель оборотно-сальдовой ведомости (ОСВ) — БУ «Страховой резерв»
# ==================================================================
def write_osv_excel(path: Path, daily_net: dict[date, float]) -> None:
    """
    Пишет реалистичную ОСВ по счёту 397.03: служебная шапка, строки оборотов
    по дням «Обороты за DD.MM.YY» с колонками Дебет/Кредит, итоговые строки.

    Показатель = Кредит − Дебет. Для большинства дней оборот идёт по Кредиту,
    но на одну дату добавляется Дебет (с компенсацией в Кредите) — проверка
    корректного вычитания. Итоговые строки без даты должны игнорироваться.

    Колонки: A=счёт/метка, B/C=сальдо нач (Дт/Кт), D/E=обороты (Дт/Кт),
             F/G=сальдо кон (Дт/Кт).
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ОСВ"

    ws.cell(row=1, column=1, value='АО "НПФ Пример"')
    ws.cell(row=2, column=1,
            value="Оборотно-сальдовая ведомость по счету (пост. 803-П) 397.03 "
                  "за Февраль 2026 г.")
    ws.cell(row=3, column=1, value="Выводимые данные: БУ (данные бухгалтерского учета)")

    # Шапка таблицы (строки 5-6) — нужна для авто-определения колонок оборотов
    ws.cell(row=5, column=1, value="Счет")
    ws.cell(row=5, column=2, value="Сальдо на начало периода")
    ws.cell(row=5, column=4, value="Обороты за период")
    ws.cell(row=5, column=6, value="Сальдо на конец периода")
    for col, label in ((2, "Дебет"), (3, "Кредит"), (4, "Дебет"),
                       (5, "Кредит"), (6, "Дебет"), (7, "Кредит")):
        ws.cell(row=6, column=col, value=label)

    total_net = round(sum(daily_net.values()), 2)
    extra_debit_date = date(2026, 2, 7)  # день с оборотом и по Дебету

    # Итоговая строка по счёту (без даты → парсер её пропустит)
    ws.cell(row=7, column=1, value="397.03")
    ws.cell(row=7, column=4, value=2_000.00)             # обороты Дебет (итог)
    ws.cell(row=7, column=5, value=total_net + 2_000.00)  # обороты Кредит (итог)

    # Группировка по стратегии (метка без даты → пропуск)
    r = 8
    ws.cell(row=r, column=1, value="(Стратегии) Базовая")
    r += 1

    ds = cycle(["dd.mm.yy"])
    for d in sorted(daily_net):
        net = daily_net[d]
        if d == extra_debit_date:
            debit, credit = 2_000.00, round(net + 2_000.00, 2)
        else:
            debit, credit = None, net
        ws.cell(row=r, column=1, value=f"Обороты за {_render_date(d, next(ds))}")
        if debit is not None:
            ws.cell(row=r, column=4, value=debit)
        ws.cell(row=r, column=5, value=credit)
        r += 1

    # Подытог по стратегии (без даты → пропуск)
    ws.cell(row=r, column=1, value="Итого по стратегии")
    ws.cell(row=r, column=5, value=total_net)

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    _print_created(path)


# ==================================================================
# Писатель файла «Выплата»
# ==================================================================
def write_vyplata_excel(path: Path, vyplata_rows: list[tuple]) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Выплаты"

    headers = ["№", "Дата операции", "ВидВыплат", "Договор", "ФИО", "Сумма", "Примечание"]
    for col_idx, h in enumerate(headers, start=1):
        ws.cell(row=1, column=col_idx, value=h).font = Font(bold=True)

    ds = cycle(["dd.mm.yyyy", "datetime", "iso"])
    r = 2
    for i, (d, type_name, amount) in enumerate(vyplata_rows):
        ws.cell(row=r, column=1, value=i + 1)
        ws.cell(row=r, column=2, value=_render_date(d, next(ds)) if d else None)
        ws.cell(row=r, column=3, value=type_name)
        ws.cell(row=r, column=4, value=f"ДПО-{1000 + i}")
        ws.cell(row=r, column=5, value="Иванов И.И.")
        ws.cell(row=r, column=6, value=amount)
        ws.cell(row=r, column=7, value="")
        r += 1

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    _print_created(path)


# ==================================================================
# Писатель JSON НПО
# ==================================================================
def write_npo_json(path: Path, component_daily: dict[str, dict[date, float]]) -> None:
    def _amount_str(v: float, idx: int) -> str:
        if v < 0:
            return (f"({abs(v):,.2f})".replace(",", " ")
                    if idx % 2 == 0 else f"{abs(v):.2f}-")
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
        # Неизвестный компонент (должен игнорироваться)
        if d == DATES[0]:
            components.append({"account": "99.01", "name": "Прочие операции НПО",
                               "amount": "777.77"})
        if components:
            total = round(sum(daily.get(d, 0.0) for daily in component_daily.values()), 2)
            items.append({
                "date": f"{d.isoformat()}T00:00:00" if di % 2 == 0 else d.isoformat(),
                "amount": str(total),
                "components": components,
            })

    data = {"report": {"startDate": "2026-02-01T00:00:00",
                       "endDate": "2026-02-28T00:00:00", "items": items}}
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    _print_created(path)


def _write_noise_files(pu_dir: Path) -> None:
    """Создаёт нерелевантные/временные файлы, которые система должна пропустить."""
    # Excel без подходящего правила
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Служебная информация", "Значение"])
    ws.append(["Подготовил", "Отдел сопровождения"])
    wb.save(pu_dir / "Служебная информация.xlsx")
    _print_created(pu_dir / "Служебная информация.xlsx")

    # Неподдерживаемое расширение
    (pu_dir / "readme.txt").write_text("Служебный файл, не для обработки.", encoding="utf-8")
    _print_created(pu_dir / "readme.txt")

    # Временный файл Excel (открытый в Office) — пропускается по префиксу ~$
    (pu_dir / "~$Выплата_Февраль2026.xlsx").write_bytes(b"PK\x03\x04tmp")
    _print_created(pu_dir / "~$Выплата_Февраль2026.xlsx")


# ==================================================================
# Данные показателей (источник истины)
#   индексы дат: 0=01 1=03 2=05 3=07 4=10 5=12 6=14 7=17 8=20 9=24 10=28
# ==================================================================
PV_FL = _mk([100_000.00, 75_500.50, 60_250.25, 33_000.00, 88_000.00, 51_000.00,
             42_300.10, 200_000.00, 91_000.00, 53_200.00, 307_455.75])
PV_FL_BU = {**PV_FL, date(2026, 2, 28): 312_455.75}          # +5 000 (28.02)

PV_UL = _mk([55_000.00, 33_200.00, 27_000.00, 19_000.00, 41_000.00, 28_000.00,
             19_500.00, 88_500.00, 36_000.00, 22_750.50, 120_000.00])

CV_FL = _mk([12_000.00, 8_500.00, 6_000.00, 4_200.00, 9_900.00, 5_500.00,
             -1_500.00, 25_000.00, 7_200.00, 4_300.00, 31_000.00])   # возврат 14.02

CV_UL_BU = _mk([45_000.00, 29_000.00, 18_000.00, 14_000.00, 30_000.00, 16_500.00,
                12_500.00, 67_500.00, 21_000.00, 15_000.00, 95_000.00])
CV_UL_PU = {**CV_UL_BU,
            date(2026, 2, 10): 23_000.00,    # −7 000 к БУ
            date(2026, 2, 24): 22_000.00}    # +7 000 к БУ (итог равен)

SR_PU = _mk([50_000.00, 30_000.00, 22_000.00, 17_000.00, 41_000.00, 25_000.00,
             18_000.00, 85_000.00, 27_500.00, 16_000.00, 60_000.00])
SR_BU = {**SR_PU, date(2026, 2, 17): 100_000.00}             # +15 000 (17.02)


# ==================================================================
# Главная функция генерации
# ==================================================================
def generate(data_dir: Path | None = None) -> None:
    if data_dir is None:
        data_dir = _get_data_dir()

    pu_dir = data_dir / "ПУ"
    bu_dir = data_dir / "БУ"
    pu_dir.mkdir(parents=True, exist_ok=True)
    bu_dir.mkdir(parents=True, exist_ok=True)

    print(f"\nГенерация тестовых данных в: {data_dir}")
    print(f"  ПУ: {pu_dir}\n  БУ: {bu_dir}\n")

    # ---------------- ПУ ----------------
    write_generic_excel(
        pu_dir / "Пенсионные взносы ФЛ_Февраль2026.xlsx", "Пенсионные взносы ФЛ",
        _split_rows(PV_FL, split_dates={date(2026, 2, 1), date(2026, 2, 17)}),
        amount_col_idx=7, date_styles=["dd.mm.yyyy"], amount_styles=["plain"])

    write_generic_excel(
        pu_dir / "Пенсионные взносы ЮЛ_Февраль2026.xlsx", "Пенсионные взносы ЮЛ",
        _split_rows(PV_UL), amount_col_idx=7,
        date_styles=["datetime"], amount_styles=["nbsp", "plain"])

    write_generic_excel(
        pu_dir / "Целевые взносы ФЛ_Февраль2026.xlsx", "Целевые взносы ФЛ",
        _split_rows(CV_FL), amount_col_idx=7,
        date_styles=["dd.mm.yy", "dd.mm.yyyy"], amount_styles=["narrow", "space"])

    write_generic_excel(
        pu_dir / "Целевые взносы ЮЛ_Февраль2026.xlsx", "Целевые взносы ЮЛ",
        _split_rows(CV_UL_PU, split_dates={date(2026, 2, 28)}), amount_col_idx=2,  # C
        date_styles=["datetime", "iso"], amount_styles=["plain"])

    write_generic_excel(
        pu_dir / "Страховой резерв_Февраль2026.xlsx", "Страховой резерв",
        _split_rows(SR_PU), amount_col_idx=8,                                       # I
        date_styles=["dd.mm.yyyy", "datetime"], amount_styles=["space", "plain"])

    # Значения ВидВыплат намеренно взяты как в реальных выгрузках: с опечаткой
    # («Негосударсвенная») и без пробела перед скобкой — проверка устойчивого
    # сопоставления по ключевым словам.
    NP, VR, VN = ("Негосударсвенная пенсия",
                  "Выкупная сумма(Расторжение)",
                  "Выкупная сумма(Наследникам)")
    # Соответствие ВидВыплат → показатель (для расчёта БУ-стороны выплат из JSON)
    _VYPL_IND = {NP: "Выплата пенсии", VR: "Выплата выкупных сумм",
                 VN: "Выплата наследуемых сумм"}
    vyplata_rows = [
        (date(2026, 2, 1), NP, 8_500.00), (date(2026, 2, 1), NP, 2_300.00),
        (date(2026, 2, 3), NP, 7_200.00), (date(2026, 2, 5), NP, 6_400.00),
        (date(2026, 2, 7), NP, 5_900.00), (date(2026, 2, 10), NP, 8_900.00),
        (date(2026, 2, 12), NP, 6_300.00), (date(2026, 2, 14), NP, 7_750.00),
        (date(2026, 2, 17), NP, 9_100.00), (date(2026, 2, 20), NP, 8_050.00),
        (date(2026, 2, 24), NP, 7_600.00), (date(2026, 2, 28), NP, 8_800.00),
        (date(2026, 2, 3), VR, 15_000.00), (date(2026, 2, 14), VR, 18_400.00),
        (date(2026, 2, 17), VR, 22_500.00), (date(2026, 2, 28), VR, 12_300.00),
        (date(2026, 2, 17), VN, 5_000.00), (date(2026, 2, 24), VN, 6_700.00),
        (date(2026, 2, 28), VN, 3_300.00),
        (date(2026, 2, 28), "Прочие выплаты", 999.99),   # не в маппинге → пропуск
        (None, NP, 1_000.00),                            # нет даты → пропуск
        (date(2026, 2, 20), NP, None),                   # нет суммы → пропуск
    ]
    write_vyplata_excel(pu_dir / "Выплата_Февраль2026.xlsx", vyplata_rows)

    _write_noise_files(pu_dir)
    print()

    # БУ-сторона выплат берётся из тех же данных НПО и совпадает с ПУ (демонстрация
    # того, что строки выплат теперь заполняются с обеих сторон и сверяются).
    pay_daily: dict[str, dict[date, float]] = {
        "Выплата пенсии": {}, "Выплата выкупных сумм": {}, "Выплата наследуемых сумм": {}}
    for d, type_name, amount in vyplata_rows:
        ind = _VYPL_IND.get(type_name)
        if d is None or amount is None or ind is None:
            continue
        pay_daily[ind][d] = round(pay_daily[ind].get(d, 0.0) + amount, 2)

    # ---------------- БУ ----------------
    # БУ «Страховой резерв» — оборотно-сальдовая ведомость (Кредит − Дебет).
    write_osv_excel(bu_dir / "397.03 Страховой резерв Февраль_2026.xlsx", SR_BU)

    write_npo_json(
        bu_dir / "НПО_Февраль2026.json",
        {
            "Пенсионные взносы ФЛ по договорам НПО": PV_FL_BU,
            "Пенсионные взносы ЮЛ по договорам НПО": PV_UL,
            "Целевые взносы ФЛ": CV_FL,
            "Целевые взносы ЮЛ": CV_UL_BU,
            "Выплаты негосударственных пенсий НПО": pay_daily["Выплата пенсии"],
            "Выплаты выкупных сумм НПО": pay_daily["Выплата выкупных сумм"],
            "Выплаты наследуемых сумм НПО": pay_daily["Выплата наследуемых сумм"],
        })

    print()
    print("=" * 66)
    print("  Тестовые данные сгенерированы (11 дат февраля 2026).")
    print()
    print("  Намеренные расхождения:")
    print("    1. Страховой резерв, 17.02.2026:  БУ−ПУ = +15 000,00 (по итогу)")
    print("    2. Пенсионные взносы ФЛ, 28.02.2026:  БУ−ПУ = +5 000,00 (по итогу)")
    print("    3. Целевые взносы ЮЛ:  итоги РАВНЫ, отличия по дням")
    print("       10.02.2026: +7 000,00 ; 24.02.2026: −7 000,00")
    print("=" * 66)


if __name__ == "__main__":
    generate()
