"""
Тесты парсеров: excel_generic, excel_vyplata, json_npo.

Используют временные файлы в pytest tmp_path для изолированного тестирования.
"""
from __future__ import annotations

import datetime
import json
from pathlib import Path

import openpyxl
import pytest

from npf_recon.models import Record
from npf_recon.sources.base import RawDocument


# ------------------------------------------------------------------
# Вспомогательные фабрики
# ------------------------------------------------------------------

def _make_raw(path: Path, side: str = "ПУ") -> RawDocument:
    ext = path.suffix.lower().lstrip(".")
    fmt = "excel" if ext in ("xlsx", "xls") else "json"
    return RawDocument(
        path=path,
        base_name=path.stem.lower(),
        ext=ext,
        fmt=fmt,
        side=side,  # type: ignore[arg-type]
    )


def _make_generic_rule(indicator: str, date_col: int, amount_col: int, side: str = "ПУ"):
    """Создаёт FileRule для excel_generic парсера."""
    import sys, os
    sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "config"))
    from config.mappings import FileRule
    return FileRule(
        name_contains="test",
        side=side,  # type: ignore[arg-type]
        fmt="excel",
        parser_name="excel_generic",
        params={"date_col": date_col, "amount_col": amount_col, "indicator": indicator},
    )


def _make_vyplata_rule():
    from config.mappings import FileRule, VYPLATA_TYPE_MAP
    return FileRule(
        name_contains="выплата",
        side="ПУ",
        fmt="excel",
        parser_name="vyplata",
        params={
            "date_col_header": "Дата операции",
            "type_col_header": "ВидВыплат",
            "amount_col_header": "Сумма",
            "type_map": VYPLATA_TYPE_MAP,
        },
    )


def _make_npo_rule():
    from config.mappings import FileRule, NPO_COMPONENT_MAP
    return FileRule(
        name_contains="нпо",
        side="БУ",
        fmt="json",
        parser_name="json_npo",
        params={"component_map": NPO_COMPONENT_MAP},
    )


# ------------------------------------------------------------------
# Тесты excel_generic
# ------------------------------------------------------------------

class TestExcelGenericParser:
    """Тесты универсального Excel-парсера."""

    def _write_excel(self, path: Path, rows: list[tuple]) -> None:
        """Создаёт тестовый Excel с позиционными данными."""
        wb = openpyxl.Workbook()
        ws = wb.active
        # Строка 1: заголовок (должен пропускаться)
        ws.append(["Дата", None, None, None, None, None, None, "Сумма"])
        for row in rows:
            ws.append(list(row))
        wb.save(path)

    def test_basic_parse(self, tmp_path: Path):
        """Базовый разбор: дата в A, сумма в H (индекс 7)."""
        from npf_recon.parsers.excel_generic import ExcelGenericParser

        xlsx = tmp_path / "test.xlsx"
        # 8 столбцов: A=дата, B-G=пустые, H=сумма
        self._write_excel(xlsx, [
            [datetime.date(2026, 2, 1), None, None, None, None, None, None, 100_000.0],
            [datetime.date(2026, 2, 3), None, None, None, None, None, None, 75_500.50],
            ["ИТОГО", None, None, None, None, None, None, 175_500.50],  # должен пропуститься
        ])

        rule = _make_generic_rule("Тест ФЛ", date_col=0, amount_col=7)
        raw = _make_raw(xlsx)
        parser = ExcelGenericParser()
        records = parser.parse(raw, rule)

        assert len(records) == 2
        assert all(r.indicator == "Тест ФЛ" for r in records)
        assert all(r.side == "ПУ" for r in records)
        assert records[0].date == datetime.date(2026, 2, 1)
        assert records[0].amount == pytest.approx(100_000.0)
        assert records[1].date == datetime.date(2026, 2, 3)
        assert records[1].amount == pytest.approx(75_500.50)

    def test_string_dates(self, tmp_path: Path):
        """Строковые даты формата DD.MM.YYYY должны парситься корректно."""
        from npf_recon.parsers.excel_generic import ExcelGenericParser

        xlsx = tmp_path / "test_str_dates.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Дата", None, None, None, None, None, None, "Сумма"])
        ws.append(["01.02.2026", None, None, None, None, None, None, 50_000.0])
        ws.append(["28.02.2026", None, None, None, None, None, None, "1 234,56"])
        wb.save(xlsx)

        rule = _make_generic_rule("Тест", date_col=0, amount_col=7)
        raw = _make_raw(xlsx)
        parser = ExcelGenericParser()
        records = parser.parse(raw, rule)

        assert len(records) == 2
        assert records[0].date == datetime.date(2026, 2, 1)
        assert records[1].date == datetime.date(2026, 2, 28)
        assert records[1].amount == pytest.approx(1234.56)

    def test_amount_col_c(self, tmp_path: Path):
        """Сумма в колонке C (индекс 2)."""
        from npf_recon.parsers.excel_generic import ExcelGenericParser

        xlsx = tmp_path / "test_col_c.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Дата", None, "Сумма"])  # C=2
        ws.append([datetime.date(2026, 2, 17), None, 67_500.0])
        wb.save(xlsx)

        rule = _make_generic_rule("Целевые ЮЛ", date_col=0, amount_col=2)
        raw = _make_raw(xlsx)
        parser = ExcelGenericParser()
        records = parser.parse(raw, rule)

        assert len(records) == 1
        assert records[0].amount == pytest.approx(67_500.0)

    def test_empty_file(self, tmp_path: Path):
        """Пустой файл возвращает пустой список."""
        from npf_recon.parsers.excel_generic import ExcelGenericParser

        xlsx = tmp_path / "empty.xlsx"
        wb = openpyxl.Workbook()
        wb.save(xlsx)

        rule = _make_generic_rule("Тест", date_col=0, amount_col=7)
        raw = _make_raw(xlsx)
        parser = ExcelGenericParser()
        records = parser.parse(raw, rule)
        assert records == []

    def test_source_file_name(self, tmp_path: Path):
        """source_file содержит имя файла."""
        from npf_recon.parsers.excel_generic import ExcelGenericParser

        xlsx = tmp_path / "МойФайл_2026.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append([datetime.date(2026, 2, 1), None, None, None, None, None, None, 1000.0])
        wb.save(xlsx)

        rule = _make_generic_rule("Тест", date_col=0, amount_col=7)
        raw = _make_raw(xlsx)
        parser = ExcelGenericParser()
        records = parser.parse(raw, rule)
        assert len(records) == 1
        assert records[0].source_file == "МойФайл_2026.xlsx"


# ------------------------------------------------------------------
# Тесты excel_vyplata
# ------------------------------------------------------------------

class TestExcelVyplataParser:
    """Тесты парсера файла выплат."""

    def _write_vyplata(self, path: Path, rows: list[tuple]) -> None:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["№", "Дата операции", "ВидВыплат", "Договор", "ФИО", "Сумма"])
        for row in rows:
            ws.append(list(row))
        wb.save(path)

    def test_basic_parse(self, tmp_path: Path):
        """Базовый разбор выплат с маппингом видов."""
        from npf_recon.parsers.excel_vyplata import ExcelVyplataParser

        xlsx = tmp_path / "Выплата_тест.xlsx"
        self._write_vyplata(xlsx, [
            [1, datetime.date(2026, 2, 1), "Негосударственная пенсия", "ДПО-1", "Иванов", 8_500.0],
            [2, datetime.date(2026, 2, 3), "Выкупная сумма (Расторжение)", "ДПО-2", "Петров", 15_000.0],
            [3, datetime.date(2026, 2, 17), "Выкупная сумма (Наследникам)", "ДПО-3", "Сидоров", 5_000.0],
            [4, datetime.date(2026, 2, 28), "Прочие выплаты", "ДПО-4", "Козлов", 999.0],  # пропуск
        ])

        rule = _make_vyplata_rule()
        raw = _make_raw(xlsx)
        parser = ExcelVyplataParser()
        records = parser.parse(raw, rule)

        assert len(records) == 3
        indicators = {r.indicator for r in records}
        assert "Выплата пенсии" in indicators
        assert "Выплата выкупных сумм" in indicators
        assert "Выплата наследуемых сумм" in indicators
        # «Прочие выплаты» должна быть пропущена
        assert all(r.indicator != "Прочие выплаты" for r in records)

    def test_string_dates(self, tmp_path: Path):
        """Строковые даты в колонке «Дата операции»."""
        from npf_recon.parsers.excel_vyplata import ExcelVyplataParser

        xlsx = tmp_path / "Выплата_str.xlsx"
        self._write_vyplata(xlsx, [
            [1, "01.02.2026", "Негосударственная пенсия", "ДПО-1", "Иванов", 8_500.0],
        ])

        rule = _make_vyplata_rule()
        raw = _make_raw(xlsx)
        parser = ExcelVyplataParser()
        records = parser.parse(raw, rule)

        assert len(records) == 1
        assert records[0].date == datetime.date(2026, 2, 1)

    def test_missing_column(self, tmp_path: Path):
        """Если нет обязательной колонки — возвращается пустой список."""
        from npf_recon.parsers.excel_vyplata import ExcelVyplataParser

        xlsx = tmp_path / "Выплата_broken.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        # Нет колонки «ВидВыплат»
        ws.append(["Дата операции", "Сумма"])
        ws.append([datetime.date(2026, 2, 1), 1000.0])
        wb.save(xlsx)

        rule = _make_vyplata_rule()
        raw = _make_raw(xlsx)
        parser = ExcelVyplataParser()
        records = parser.parse(raw, rule)
        assert records == []

    def test_mapping_all_types(self, tmp_path: Path):
        """Все три типа выплат маппируются корректно."""
        from npf_recon.parsers.excel_vyplata import ExcelVyplataParser

        xlsx = tmp_path / "Выплата_all.xlsx"
        self._write_vyplata(xlsx, [
            [1, datetime.date(2026, 2, 1), "Негосударственная пенсия", "ДПО-1", "А", 100.0],
            [2, datetime.date(2026, 2, 1), "Выкупная сумма (Расторжение)", "ДПО-2", "Б", 200.0],
            [3, datetime.date(2026, 2, 1), "Выкупная сумма (Наследникам)", "ДПО-3", "В", 300.0],
        ])

        rule = _make_vyplata_rule()
        raw = _make_raw(xlsx)
        parser = ExcelVyplataParser()
        records = parser.parse(raw, rule)

        assert len(records) == 3
        by_ind = {r.indicator: r.amount for r in records}
        assert by_ind["Выплата пенсии"] == pytest.approx(100.0)
        assert by_ind["Выплата выкупных сумм"] == pytest.approx(200.0)
        assert by_ind["Выплата наследуемых сумм"] == pytest.approx(300.0)


# ------------------------------------------------------------------
# Тесты json_npo
# ------------------------------------------------------------------

class TestJsonNpoParser:
    """Тесты парсера JSON НПО."""

    def _write_json(self, path: Path, data: dict) -> None:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def _make_npo_data(self, items: list[dict]) -> dict:
        return {
            "report": {
                "startDate": "2026-02-01T00:00:00",
                "endDate": "2026-02-28T00:00:00",
                "items": items,
            }
        }

    def test_basic_parse(self, tmp_path: Path):
        """Базовый разбор JSON НПО с маппингом компонентов."""
        from npf_recon.parsers.json_npo import JsonNpoParser

        json_path = tmp_path / "НПО_тест.json"
        self._write_json(json_path, self._make_npo_data([
            {
                "date": "2026-02-01T00:00:00",
                "amount": "175500.50",
                "components": [
                    {"account": "76.01", "name": "Пенсионные взносы ФЛ по договорам НПО", "amount": "100000.00"},
                    {"account": "76.02", "name": "Пенсионные взносы ЮЛ по договорам НПО", "amount": "75500.50"},
                ],
            },
            {
                "date": "2026-02-03",
                "amount": "12000.00",
                "components": [
                    {"account": "76.01", "name": "Целевые взносы ФЛ", "amount": "12000.00"},
                ],
            },
        ]))

        rule = _make_npo_rule()
        raw = _make_raw(json_path, side="БУ")
        parser = JsonNpoParser()
        records = parser.parse(raw, rule)

        assert len(records) == 3
        pv_fl = [r for r in records if r.indicator == "Пенсионные взносы ФЛ"]
        assert len(pv_fl) == 1
        assert pv_fl[0].amount == pytest.approx(100_000.0)
        assert pv_fl[0].date == datetime.date(2026, 2, 1)
        assert pv_fl[0].side == "БУ"
        assert pv_fl[0].extra.get("account") == "76.01"

    def test_unmapped_components_skipped(self, tmp_path: Path):
        """Компоненты вне маппинга пропускаются."""
        from npf_recon.parsers.json_npo import JsonNpoParser

        json_path = tmp_path / "НПО_unmapped.json"
        self._write_json(json_path, self._make_npo_data([
            {
                "date": "2026-02-01T00:00:00",
                "amount": "999.00",
                "components": [
                    {"account": "99.01", "name": "Прочие операции НПО", "amount": "999.00"},
                ],
            },
        ]))

        rule = _make_npo_rule()
        raw = _make_raw(json_path, side="БУ")
        parser = JsonNpoParser()
        records = parser.parse(raw, rule)
        assert records == []

    def test_iso_date_formats(self, tmp_path: Path):
        """Оба ISO-формата дат (с T и без) разбираются корректно."""
        from npf_recon.parsers.json_npo import JsonNpoParser

        json_path = tmp_path / "НПО_dates.json"
        self._write_json(json_path, self._make_npo_data([
            {
                "date": "2026-02-17T00:00:00",
                "amount": "200000.00",
                "components": [
                    {"account": "76.01", "name": "Пенсионные взносы ФЛ по договорам НПО", "amount": "200000.00"},
                ],
            },
            {
                "date": "2026-02-28",
                "amount": "312455.75",
                "components": [
                    {"account": "76.01", "name": "Пенсионные взносы ФЛ по договорам НПО", "amount": "312455.75"},
                ],
            },
        ]))

        rule = _make_npo_rule()
        raw = _make_raw(json_path, side="БУ")
        parser = JsonNpoParser()
        records = parser.parse(raw, rule)

        assert len(records) == 2
        dates = {r.date for r in records}
        assert datetime.date(2026, 2, 17) in dates
        assert datetime.date(2026, 2, 28) in dates

    def test_broken_json(self, tmp_path: Path):
        """Битый JSON возвращает пустой список."""
        from npf_recon.parsers.json_npo import JsonNpoParser

        json_path = tmp_path / "broken.json"
        json_path.write_text("{broken json", encoding="utf-8")

        rule = _make_npo_rule()
        raw = _make_raw(json_path, side="БУ")
        parser = JsonNpoParser()
        records = parser.parse(raw, rule)
        assert records == []

    def test_amount_string_with_comma(self, tmp_path: Path):
        """Суммы в формате «100,000.00» (float с запятой) разбираются корректно."""
        from npf_recon.parsers.json_npo import JsonNpoParser

        json_path = tmp_path / "НПО_comma.json"
        self._write_json(json_path, self._make_npo_data([
            {
                "date": "2026-02-01T00:00:00",
                "amount": "100,000.00",
                "components": [
                    {"account": "76.01", "name": "Целевые взносы ЮЛ", "amount": "100,000.00"},
                ],
            },
        ]))

        rule = _make_npo_rule()
        raw = _make_raw(json_path, side="БУ")
        parser = JsonNpoParser()
        records = parser.parse(raw, rule)

        assert len(records) == 1
        assert records[0].amount == pytest.approx(100_000.0)
