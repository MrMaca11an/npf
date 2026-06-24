"""
Формирование итогового Excel-отчёта сверки БУ-ПУ.

Лист «Свод»: сводная таблица с заголовком, 14 строк показателей,
  суммами и датами по БУ, ПУ и расхождению.
Лист «Детализация»: строки с расхождениями по каждой дате.

Структура заголовка (строки 1-3):
  Строка 1: «Сверка БУ-ПУ. <Период>» (A1:G1 merged)
  Строка 2: «Виды движений» (A2:A3 merged) | «БУ» (B2:C2) | «ПУ» (D2:E2) | «Расхождение» (F2:G2)
  Строка 3: пусто (A3 — под merge) | Сумма | Дата | Сумма | Дата | Сумма | Дата
  Строки 4-17: данные (14 показателей)
"""
from __future__ import annotations

import logging
from datetime import date
from pathlib import Path

import openpyxl
from openpyxl.styles import (
    Alignment,
    Border,
    Font,
    PatternFill,
    Side,
)
from openpyxl.utils import get_column_letter, column_index_from_string

from npf_recon.models import IndicatorSide, ReconRow
from npf_recon.normalize import format_amount, format_date

logger = logging.getLogger(__name__)

# Цвета
_HEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
_DISCREP_FILL = PatternFill("solid", fgColor="FFE0E0")

_THIN = Side(style="thin")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


def _apply_border(cell) -> None:
    cell.border = _BORDER


def _format_dates(dates: list[date]) -> str:
    """
    Форматирует список дат для ячейки отчёта.

    При количестве дат <= 5 перечисляет через «, »; иначе показывает диапазон «min–max».
    """
    if not dates:
        return ""
    sorted_d = sorted(dates)
    if len(sorted_d) <= 5:
        return ", ".join(format_date(d) for d in sorted_d)
    return f"{format_date(sorted_d[0])}–{format_date(sorted_d[-1])}"


def _format_side_dates(side: IndicatorSide | None) -> str:
    if side is None or not side.dates:
        return ""
    return _format_dates(side.dates)


def write_report(
    rows: list[ReconRow],
    output_path: Path,
    period_label: str,
    eps: float = 0.005,
) -> None:
    """
    Записывает итоговый отчёт в Excel-файл.

    :param rows:         список из 14 ReconRow от reconcile().
    :param output_path:  путь к выходному файлу.
    :param period_label: метка периода (напр. «Февраль 2026»).
    :param eps:          допуск: расхождения |d| <= eps не выделяются.
    """
    output_path.parent.mkdir(parents=True, exist_ok=True)

    wb = openpyxl.Workbook()
    _write_svod(wb, rows, period_label, eps)
    _write_details(wb, rows)

    wb.save(output_path)
    logger.info("Отчёт сохранён: %s", output_path)


# ------------------------------------------------------------------
# Лист «Свод»
# ------------------------------------------------------------------

def _write_svod(
    wb: openpyxl.Workbook,
    rows: list[ReconRow],
    period_label: str,
    eps: float,
) -> None:
    ws = wb.active
    ws.title = "Свод"
    bold = Font(bold=True)

    # --- Строка 1: заголовок периода (A1:G1) ---
    ws.merge_cells("A1:G1")
    title_cell = ws["A1"]
    title_cell.value = f"Сверка БУ-ПУ. {period_label}"
    title_cell.font = Font(bold=True, size=13)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    # --- Строка 2: блоки разделов ---
    # A2:A3 — «Виды движений» (объединяем строки 2 и 3)
    ws.merge_cells("A2:A3")
    head_a = ws["A2"]
    head_a.value = "Виды движений"
    head_a.font = bold
    head_a.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    head_a.fill = _HEADER_FILL
    _apply_border(head_a)

    # Блоки: (начальная колонка в букве, название)
    block_defs = [
        ("B", "Бухгалтерский учёт"),
        ("D", "Персонифицированный учёт"),
        ("F", "Расхождение"),
    ]
    for col_letter, label in block_defs:
        col_idx = column_index_from_string(col_letter)
        col_end = get_column_letter(col_idx + 1)
        # Заголовок блока (строка 2)
        merge_ref = f"{col_letter}2:{col_end}2"
        ws.merge_cells(merge_ref)
        cell = ws[f"{col_letter}2"]
        cell.value = label
        cell.font = bold
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = _HEADER_FILL
        _apply_border(cell)
        # Правая ячейка того же merge — только бордер (остаётся MergedCell, не назначаем value)
        _apply_border(ws[f"{col_end}2"])

        # Подзаголовки «Сумма» / «Дата» (строка 3)
        ws.cell(row=3, column=col_idx).value = "Сумма"
        ws.cell(row=3, column=col_idx).font = bold
        ws.cell(row=3, column=col_idx).alignment = Alignment(horizontal="center")
        ws.cell(row=3, column=col_idx).fill = _HEADER_FILL
        _apply_border(ws.cell(row=3, column=col_idx))

        ws.cell(row=3, column=col_idx + 1).value = "Дата"
        ws.cell(row=3, column=col_idx + 1).font = bold
        ws.cell(row=3, column=col_idx + 1).alignment = Alignment(horizontal="center")
        ws.cell(row=3, column=col_idx + 1).fill = _HEADER_FILL
        _apply_border(ws.cell(row=3, column=col_idx + 1))

    # Ячейка A3 — часть merged A2:A3, не трогаем value, только стиль
    _apply_border(ws["A3"])

    # --- Строки данных (строки 4-17) ---
    for data_row_idx, recon in enumerate(rows, start=4):
        bu = recon.bu
        pu = recon.pu
        has_discrep = (
            recon.diff_total is not None and abs(recon.diff_total) > eps
        )

        row_data = [
            # (значение, колонка 1-based)
            (recon.indicator, 1),
            (format_amount(bu.total) if bu else "", 2),
            (_format_side_dates(bu), 3),
            (format_amount(pu.total) if pu else "", 4),
            (_format_side_dates(pu), 5),
            (format_amount(recon.diff_total) if has_discrep else "", 6),
            (_format_dates(recon.diff_dates) if recon.diff_dates else "", 7),
        ]

        for value, col_idx in row_data:
            cell = ws.cell(row=data_row_idx, column=col_idx, value=value)
            cell.alignment = Alignment(
                horizontal="left" if col_idx == 1 else "right",
                wrap_text=True,
            )
            _apply_border(cell)
            if has_discrep and col_idx in (6, 7):
                cell.fill = _DISCREP_FILL

    # --- Ширины колонок ---
    ws.column_dimensions["A"].width = 42
    for col_letter in ("B", "D", "F"):
        ws.column_dimensions[col_letter].width = 14
    for col_letter in ("C", "E", "G"):
        ws.column_dimensions[col_letter].width = 28

    ws.freeze_panes = "A4"


# ------------------------------------------------------------------
# Лист «Детализация»
# ------------------------------------------------------------------

def _write_details(
    wb: openpyxl.Workbook,
    rows: list[ReconRow],
) -> None:
    ws = wb.create_sheet("Детализация")
    bold = Font(bold=True)

    headers = ["Показатель", "Дата", "Сумма БУ", "Сумма ПУ", "Разница"]
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = bold
        cell.fill = _HEADER_FILL
        cell.alignment = Alignment(horizontal="center")
        _apply_border(cell)

    detail_row = 2
    for recon in rows:
        if not recon.diff_dates:
            continue
        for d in sorted(recon.diff_dates):
            bu_val = recon.bu.daily.get(d, 0.0) if recon.bu else 0.0
            pu_val = recon.pu.daily.get(d, 0.0) if recon.pu else 0.0
            diff_val = recon.diff_by_date.get(d)

            row_data = [
                recon.indicator,
                format_date(d),
                format_amount(bu_val),
                format_amount(pu_val),
                format_amount(diff_val),
            ]
            for col_idx, value in enumerate(row_data, start=1):
                cell = ws.cell(row=detail_row, column=col_idx, value=value)
                cell.alignment = Alignment(
                    horizontal="left" if col_idx <= 2 else "right"
                )
                _apply_border(cell)
                if col_idx == 5:
                    cell.fill = _DISCREP_FILL
            detail_row += 1

    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 14
    for col_letter in ("C", "D", "E"):
        ws.column_dimensions[col_letter].width = 16

    ws.freeze_panes = "A2"
