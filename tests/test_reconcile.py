"""
Тесты агрегации и сверки.
"""
from __future__ import annotations

import datetime

import pytest

from npf_recon.aggregate import aggregate, group_records
from npf_recon.models import IndicatorSide, Record, ReconRow
from npf_recon.reconcile import reconcile


# ------------------------------------------------------------------
# Вспомогательная фабрика записей
# ------------------------------------------------------------------

def _rec(indicator: str, date_: datetime.date, amount: float, side: str) -> Record:
    return Record(
        indicator=indicator,
        date=date_,
        amount=amount,
        side=side,
        source_file="test",
    )


D1 = datetime.date(2026, 2, 1)
D2 = datetime.date(2026, 2, 3)
D3 = datetime.date(2026, 2, 17)
D4 = datetime.date(2026, 2, 28)


# ------------------------------------------------------------------
# Тесты aggregate / IndicatorSide
# ------------------------------------------------------------------

class TestIndicatorSide:
    def test_basic(self):
        records = [
            _rec("Страховой резерв", D1, 50_000.0, "БУ"),
            _rec("Страховой резерв", D2, 30_000.0, "БУ"),
        ]
        side = IndicatorSide.from_records(records)
        assert side.total == pytest.approx(80_000.0)
        assert side.daily[D1] == pytest.approx(50_000.0)
        assert side.daily[D2] == pytest.approx(30_000.0)
        assert side.dates == [D1, D2]

    def test_sum_same_date(self):
        """Несколько записей на одну дату суммируются."""
        records = [
            _rec("Тест", D1, 10_000.0, "ПУ"),
            _rec("Тест", D1, 5_000.0, "ПУ"),
        ]
        side = IndicatorSide.from_records(records)
        assert side.daily[D1] == pytest.approx(15_000.0)
        assert side.total == pytest.approx(15_000.0)

    def test_dates_sorted(self):
        """Даты возвращаются в хронологическом порядке."""
        records = [
            _rec("Тест", D4, 1.0, "БУ"),
            _rec("Тест", D1, 2.0, "БУ"),
            _rec("Тест", D3, 3.0, "БУ"),
        ]
        side = IndicatorSide.from_records(records)
        assert side.dates == [D1, D3, D4]


class TestAggregate:
    def test_groups_by_indicator_and_side(self):
        records = [
            _rec("А", D1, 100.0, "БУ"),
            _rec("А", D1, 200.0, "ПУ"),
            _rec("Б", D1, 300.0, "БУ"),
        ]
        agg = aggregate(records)
        assert ("А", "БУ") in agg
        assert ("А", "ПУ") in agg
        assert ("Б", "БУ") in agg
        assert agg[("А", "БУ")].total == pytest.approx(100.0)
        assert agg[("А", "ПУ")].total == pytest.approx(200.0)


# ------------------------------------------------------------------
# Тесты reconcile
# ------------------------------------------------------------------

class TestReconcile:

    def _run(self, records: list[Record], eps: float = 0.005) -> dict[str, ReconRow]:
        agg = aggregate(records)
        rows = reconcile(agg, eps=eps)
        return {r.indicator: r for r in rows}

    def test_14_rows_always(self):
        """Reconcile всегда возвращает ровно 14 строк."""
        rows = reconcile({})
        assert len(rows) == 14

    def test_total_equals_sum_of_daily(self):
        """ИНВАРИАНТ: итог показателя == сумме распарсенных дневных значений."""
        records = [
            _rec("Пенсионные взносы ФЛ", D1, 100_000.0, "БУ"),
            _rec("Пенсионные взносы ФЛ", D2, 42_000.0, "БУ"),
            _rec("Пенсионные взносы ФЛ", D3, 8_000.0, "БУ"),
        ]
        row = self._run(records)["Пенсионные взносы ФЛ"]
        assert row.bu.total == pytest.approx(150_000.0)
        assert round(sum(row.bu.daily.values()), 2) == pytest.approx(row.bu.total)

    def test_diff_total_equals_sum_of_daily_diffs(self):
        """ИНВАРИАНТ: сумма расхождения == сумме дневных разниц (Свод == Детализация)."""
        records = [
            # БУ и ПУ: одинаковый месячный итог, но разное распределение по дням
            _rec("Целевые взносы ЮЛ", D1, 30_000.0, "БУ"),
            _rec("Целевые взносы ЮЛ", D3, 15_000.0, "БУ"),
            _rec("Целевые взносы ЮЛ", D1, 23_000.0, "ПУ"),
            _rec("Целевые взносы ЮЛ", D3, 22_000.0, "ПУ"),
        ]
        row = self._run(records)["Целевые взносы ЮЛ"]
        # Итоги совпадают → нетто-разница 0, но по дням есть отклонения
        assert row.diff_total == pytest.approx(0.0)
        assert round(sum(row.diff_by_date.values()), 2) == pytest.approx(row.diff_total)
        assert row.diff_total == pytest.approx(row.bu.total - row.pu.total)
        assert set(row.diff_dates) == {D1, D3}

    def test_empty_row(self):
        """Показатели без данных имеют bu=None, pu=None, diff_total=None."""
        rows = reconcile({})
        by_ind = {r.indicator: r for r in rows}
        row = by_ind["ИД"]
        assert row.bu is None
        assert row.pu is None
        assert row.diff_total is None
        assert row.diff_dates == []

    def test_discrepancy_detected(self):
        """Расхождение по Страховому резерву: БУ=100 000, ПУ=85 000 → diff=15 000."""
        records = [
            _rec("Страховой резерв", D3, 100_000.0, "БУ"),
            _rec("Страховой резерв", D3, 85_000.0, "ПУ"),
        ]
        by_ind = self._run(records)
        row = by_ind["Страховой резерв"]
        assert row.diff_total == pytest.approx(15_000.0)
        assert D3 in row.diff_dates
        assert row.diff_by_date[D3] == pytest.approx(15_000.0)

    def test_no_discrepancy_within_eps(self):
        """Разница в пределах eps не считается расхождением."""
        records = [
            _rec("Страховой резерв", D1, 50_000.001, "БУ"),
            _rec("Страховой резерв", D1, 50_000.000, "ПУ"),
        ]
        by_ind = self._run(records, eps=0.005)
        row = by_ind["Страховой резерв"]
        # Разница 0.001 < eps 0.005 → diff_dates пусто
        assert row.diff_dates == []

    def test_one_sided_no_discrepancy(self):
        """Показатель только с одной стороны → diff_total=None."""
        records = [
            _rec("Выплата пенсии", D1, 8_500.0, "ПУ"),
            _rec("Выплата пенсии", D2, 7_200.0, "ПУ"),
        ]
        by_ind = self._run(records)
        row = by_ind["Выплата пенсии"]
        assert row.pu is not None
        assert row.bu is None
        assert row.diff_total is None
        assert row.diff_dates == []

    def test_pv_fl_discrepancy_on_date(self):
        """Расхождение по Пенсионным взносам ФЛ на 28.02.2026."""
        records = [
            _rec("Пенсионные взносы ФЛ", D1, 100_000.0, "БУ"),
            _rec("Пенсионные взносы ФЛ", D1, 100_000.0, "ПУ"),
            _rec("Пенсионные взносы ФЛ", D4, 312_455.75, "БУ"),
            _rec("Пенсионные взносы ФЛ", D4, 307_455.75, "ПУ"),
        ]
        by_ind = self._run(records)
        row = by_ind["Пенсионные взносы ФЛ"]
        assert row.diff_total == pytest.approx(5_000.0)
        assert D4 in row.diff_dates
        assert D1 not in row.diff_dates

    def test_exact_match_no_dates(self):
        """Идеальное совпадение: diff_total=0, diff_dates=[]."""
        records = [
            _rec("Страховой резерв", D1, 50_000.0, "БУ"),
            _rec("Страховой резерв", D1, 50_000.0, "ПУ"),
            _rec("Страховой резерв", D2, 30_000.0, "БУ"),
            _rec("Страховой резерв", D2, 30_000.0, "ПУ"),
        ]
        by_ind = self._run(records)
        row = by_ind["Страховой резерв"]
        assert row.diff_total == pytest.approx(0.0)
        assert row.diff_dates == []

    def test_multiple_date_discrepancies(self):
        """Расхождения на нескольких датах."""
        records = [
            _rec("Страховой резерв", D1, 50_000.0, "БУ"),
            _rec("Страховой резерв", D1, 45_000.0, "ПУ"),  # -5 000
            _rec("Страховой резерв", D3, 100_000.0, "БУ"),
            _rec("Страховой резерв", D3, 85_000.0, "ПУ"),  # -15 000
        ]
        by_ind = self._run(records)
        row = by_ind["Страховой резерв"]
        assert len(row.diff_dates) == 2
        assert D1 in row.diff_dates
        assert D3 in row.diff_dates
        assert row.diff_by_date[D1] == pytest.approx(5_000.0)
        assert row.diff_by_date[D3] == pytest.approx(15_000.0)

    def test_net_zero_total_but_daily_diff(self):
        """Итоги совпадают, но есть встречные дневные отклонения → дни в diff_dates."""
        records = [
            # БУ: 10.02=30000, 24.02=15000 (итог 45000)
            _rec("Целевые взносы ЮЛ", D3, 30_000.0, "БУ"),
            _rec("Целевые взносы ЮЛ", D4, 15_000.0, "БУ"),
            # ПУ: 10.02=23000, 24.02=22000 (итог 45000 — равен БУ)
            _rec("Целевые взносы ЮЛ", D3, 23_000.0, "ПУ"),
            _rec("Целевые взносы ЮЛ", D4, 22_000.0, "ПУ"),
        ]
        by_ind = self._run(records)
        row = by_ind["Целевые взносы ЮЛ"]
        assert row.diff_total == pytest.approx(0.0)         # итоги равны
        assert set(row.diff_dates) == {D3, D4}              # но дни отличаются
        assert row.diff_by_date[D3] == pytest.approx(7_000.0)
        assert row.diff_by_date[D4] == pytest.approx(-7_000.0)

    def test_negative_correction_no_discrepancy(self):
        """Отрицательная корректировка, равная в БУ и ПУ, не даёт расхождения."""
        records = [
            _rec("Целевые взносы ФЛ", D2, 8_500.0, "БУ"),
            _rec("Целевые взносы ФЛ", D3, -1_500.0, "БУ"),
            _rec("Целевые взносы ФЛ", D2, 8_500.0, "ПУ"),
            _rec("Целевые взносы ФЛ", D3, -1_500.0, "ПУ"),
        ]
        by_ind = self._run(records)
        row = by_ind["Целевые взносы ФЛ"]
        assert row.bu.total == pytest.approx(7_000.0)
        assert row.diff_total == pytest.approx(0.0)
        assert row.diff_dates == []

    def test_fixed_order(self):
        """Порядок строк соответствует REPORT_INDICATORS."""
        from config.mappings import REPORT_INDICATORS
        rows = reconcile({})
        assert [r.indicator for r in rows] == REPORT_INDICATORS

    def test_rppо_rows_empty(self):
        """Строки РППО (нет данных в файлах) остаются пустыми."""
        rows = reconcile({})
        by_ind = {r.indicator: r for r in rows}
        for ind in [
            "РППО инвестиционный, остаток на начало периода",
            "РППО страховой, остаток на начало периода",
            "РППО инвестиционный, остаток на конец периода",
            "РППО страховой, остаток на конец периода",
        ]:
            assert by_ind[ind].bu is None
            assert by_ind[ind].pu is None
