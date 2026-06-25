"""
Тесты обезличенного экспорта для LLM: build_summary, build_prompt, write_llm_export.
"""
from __future__ import annotations

import datetime
import json
from pathlib import Path

import pytest

from npf_recon.aggregate import aggregate
from npf_recon.llm_export import build_prompt, build_summary, write_llm_export
from npf_recon.models import Record
from npf_recon.reconcile import reconcile


D1 = datetime.date(2026, 2, 10)
D2 = datetime.date(2026, 2, 17)
D3 = datetime.date(2026, 2, 24)


def _rows(records: list[Record]):
    return reconcile(aggregate(records), eps=0.005)


def _by_ind(summary: dict) -> dict[str, dict]:
    return {p["показатель"]: p for p in summary["показатели"]}


class TestBuildSummary:
    def test_has_14_indicators(self):
        s = build_summary(_rows([]), "Февраль 2026")
        assert len(s["показатели"]) == 14
        assert s["итоги"]["всего_показателей"] == 14
        assert s["обезличено"] is True

    def test_total_discrepancy(self):
        recs = [
            Record("Страховой резерв", D2, 100_000.0, "БУ", "f"),
            Record("Страховой резерв", D2, 85_000.0, "ПУ", "f"),
        ]
        p = _by_ind(build_summary(_rows(recs), "Февраль 2026"))["Страховой резерв"]
        assert p["сверка_возможна"] is True
        assert p["есть_расхождение"] is True
        assert p["тип_расхождения"] in ("по_итогу", "по_итогу_и_по_дням")
        assert p["расхождение_итог"] == pytest.approx(15_000.0)
        assert p["дни_расхождений"][0]["дата"] == "17.02.2026"
        assert p["дни_расхождений"][0]["разница"] == pytest.approx(15_000.0)

    def test_daily_only_discrepancy(self):
        """Итоги равны, но дни различаются → тип «только_по_дням»."""
        recs = [
            Record("Целевые взносы ЮЛ", D1, 30_000.0, "БУ", "f"),
            Record("Целевые взносы ЮЛ", D3, 15_000.0, "БУ", "f"),
            Record("Целевые взносы ЮЛ", D1, 23_000.0, "ПУ", "f"),
            Record("Целевые взносы ЮЛ", D3, 22_000.0, "ПУ", "f"),
        ]
        p = _by_ind(build_summary(_rows(recs), "Февраль 2026"))["Целевые взносы ЮЛ"]
        assert p["тип_расхождения"] == "только_по_дням"
        assert p["расхождение_итог"] == pytest.approx(0.0)
        assert {x["дата"] for x in p["дни_расхождений"]} == {"10.02.2026", "24.02.2026"}

    def test_one_sided_not_reconcilable(self):
        recs = [Record("Выплата пенсии", D1, 8_500.0, "ПУ", "f")]
        p = _by_ind(build_summary(_rows(recs), "Февраль 2026"))["Выплата пенсии"]
        assert p["сверка_возможна"] is False
        assert p["тип_расхождения"] == "сверка_невозможна"
        assert p["итог_ПУ"] == pytest.approx(8_500.0)
        assert p["итог_БУ"] is None

    def test_no_data_indicator(self):
        p = _by_ind(build_summary(_rows([]), "Февраль 2026"))["ИД"]
        assert p["есть_данные_БУ"] is False and p["есть_данные_ПУ"] is False
        assert p["есть_расхождение"] is False

    def test_counts(self):
        recs = [
            Record("Страховой резерв", D2, 100_000.0, "БУ", "f"),
            Record("Страховой резерв", D2, 85_000.0, "ПУ", "f"),
        ]
        s = build_summary(_rows(recs), "Февраль 2026")
        assert s["итоги"]["показателей_с_расхождениями"] == 1
        assert s["итоги"]["сверка_возможна_по"] == 1
        assert s["итоги"]["есть_расхождения"] is True

    def test_serializable(self):
        s = build_summary(_rows([]), "Февраль 2026")
        # Должно сериализоваться без ошибок
        json.dumps(s, ensure_ascii=False)


class TestAnonymization:
    """Гарантия отсутствия персональных данных в экспортируемом саммари."""

    def test_no_personal_fields(self):
        recs = [
            Record("Выплата пенсии", D1, 8_500.0, "ПУ", "Выплата_Февраль2026.xlsx",
                   extra={"account": "76.01", "fio": "Иванов И.И."}),
        ]
        text = json.dumps(build_summary(_rows(recs), "Февраль 2026"), ensure_ascii=False)
        for forbidden in ("Иванов", "ДПО", "76.01", "account", "source_file", ".xlsx"):
            assert forbidden not in text, f"в саммари просочилось: {forbidden}"


class TestBuildPrompt:
    def test_prompt_contains_instruction_and_json(self):
        s = build_summary(_rows([]), "Февраль 2026")
        prompt = build_prompt(s)
        assert "ОБЕЗЛИЧЕННЫЕ" in prompt
        assert "ВХОДНЫЕ ДАННЫЕ (JSON):" in prompt
        # JSON-часть валидна
        payload = prompt.split("ВХОДНЫЕ ДАННЫЕ (JSON):", 1)[1].strip()
        assert json.loads(payload)["период"] == "Февраль 2026"


class TestWriteExport:
    def test_writes_both_files(self, tmp_path: Path):
        recs = [
            Record("Страховой резерв", D2, 100_000.0, "БУ", "f"),
            Record("Страховой резерв", D2, 85_000.0, "ПУ", "f"),
        ]
        summary_path = tmp_path / "out" / "Саммари.json"
        prompt_path = tmp_path / "out" / "Промт.txt"
        returned = write_llm_export(
            _rows(recs), "Февраль 2026", summary_path, prompt_path,
            generated_at=datetime.date(2026, 3, 1),
        )
        assert summary_path.exists() and prompt_path.exists()
        assert returned["дата_формирования"] == "01.03.2026"
        on_disk = json.loads(summary_path.read_text(encoding="utf-8"))
        assert on_disk["период"] == "Февраль 2026"
