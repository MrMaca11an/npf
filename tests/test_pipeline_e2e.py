"""
Сквозной (e2e) тест pipeline: генерация данных → run → проверка отчёта.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

import openpyxl
import pytest


def _generate_in(data_dir: Path) -> None:
    """Генерирует тестовые данные в указанную директорию."""
    # Импортируем generate_sample_data, передавая data_dir
    scripts_dir = Path(__file__).resolve().parent.parent / "scripts"
    sys.path.insert(0, str(scripts_dir))
    import generate_sample_data
    generate_sample_data.generate(data_dir=data_dir)


@pytest.fixture
def temp_data_dir(tmp_path: Path):
    """
    Фикстура: создаёт временную директорию с тестовыми данными и
    настраивает NPF_DATA_DIR, чтобы pipeline использовал её.
    """
    _generate_in(tmp_path)

    # Переопределяем окружение
    old_env = os.environ.get("NPF_DATA_DIR")
    os.environ["NPF_DATA_DIR"] = str(tmp_path)
    yield tmp_path
    # Восстанавливаем
    if old_env is None:
        os.environ.pop("NPF_DATA_DIR", None)
    else:
        os.environ["NPF_DATA_DIR"] = old_env


@pytest.fixture
def output_file(tmp_path: Path) -> Path:
    """Путь к выходному файлу в temp_path."""
    return tmp_path / "output" / "Сверка_БУ_ПУ.xlsx"


def _run_pipeline(data_dir: Path, output_path: Path) -> int:
    """Запускает pipeline с временными путями."""
    from npf_recon.config import Config
    from npf_recon.pipeline import run

    config = Config(
        pu_dir=data_dir / "ПУ",
        bu_dir=data_dir / "БУ",
        output_file=output_path,
        period_label="Февраль 2026",
        eps=0.005,
    )
    return run(config)


# ------------------------------------------------------------------
# e2e тест
# ------------------------------------------------------------------

class TestPipelineE2E:
    """Сквозной тест от генерации данных до проверки отчёта."""

    def test_pipeline_runs_successfully(self, temp_data_dir: Path, output_file: Path):
        """Pipeline завершается с кодом 0."""
        exit_code = _run_pipeline(temp_data_dir, output_file)
        assert exit_code == 0

    def test_output_file_created(self, temp_data_dir: Path, output_file: Path):
        """Файл отчёта создаётся."""
        _run_pipeline(temp_data_dir, output_file)
        assert output_file.exists()

    def test_svod_sheet_exists(self, temp_data_dir: Path, output_file: Path):
        """Лист «Свод» присутствует."""
        _run_pipeline(temp_data_dir, output_file)
        wb = openpyxl.load_workbook(output_file)
        assert "Свод" in wb.sheetnames

    def test_svod_has_14_data_rows(self, temp_data_dir: Path, output_file: Path):
        """На листе «Свод» ровно 14 строк данных (плюс 3 строки шапки)."""
        _run_pipeline(temp_data_dir, output_file)
        wb = openpyxl.load_workbook(output_file)
        ws = wb["Свод"]
        # Строки 4..17 — данные (3 строки шапки + 14 показателей)
        data_rows = []
        for row in ws.iter_rows(min_row=4, max_row=17, min_col=1, max_col=1):
            cell = row[0]
            val = cell.value
            if val:
                data_rows.append(str(val))
        assert len(data_rows) == 14

    def test_svod_indicators_in_order(self, temp_data_dir: Path, output_file: Path):
        """Показатели в «Свод» расположены в фиксированном порядке."""
        from config.mappings import REPORT_INDICATORS
        _run_pipeline(temp_data_dir, output_file)
        wb = openpyxl.load_workbook(output_file)
        ws = wb["Свод"]
        actual = []
        for row in ws.iter_rows(min_row=4, max_row=17, min_col=1, max_col=1):
            cell = row[0]
            if cell.value:
                actual.append(str(cell.value))
        assert actual == REPORT_INDICATORS

    def test_strahovoy_rezerv_has_discrepancy(self, temp_data_dir: Path, output_file: Path):
        """«Страховой резерв» показывает расхождение на дату 17.02.2026."""
        from config.mappings import REPORT_INDICATORS
        _run_pipeline(temp_data_dir, output_file)
        wb = openpyxl.load_workbook(output_file)
        ws = wb["Свод"]

        # Находим строку «Страховой резерв»
        sr_row_idx = None
        for row in ws.iter_rows(min_row=4, max_row=17):
            if row[0].value == "Страховой резерв":
                sr_row_idx = row[0].row
                break

        assert sr_row_idx is not None, "Строка 'Страховой резерв' не найдена в Свод"

        # Колонка F (6) — «Расхождение Сумма»
        discrep_sum = ws.cell(row=sr_row_idx, column=6).value
        assert discrep_sum, "Расхождение по Страховому резерву должно быть непустым"

        # Колонка G (7) — «Расхождение Дата»
        discrep_date = ws.cell(row=sr_row_idx, column=7).value
        assert discrep_date is not None and "17.02.2026" in str(discrep_date), \
            f"Ожидалась дата 17.02.2026 в колонке расхождений, получено: {discrep_date}"

    def test_id_row_is_empty(self, temp_data_dir: Path, output_file: Path):
        """Строка «ИД» не имеет данных (нет файлов)."""
        _run_pipeline(temp_data_dir, output_file)
        wb = openpyxl.load_workbook(output_file)
        ws = wb["Свод"]

        for row in ws.iter_rows(min_row=4, max_row=17):
            if row[0].value == "ИД":
                bu_sum = row[1].value
                pu_sum = row[3].value
                assert not bu_sum, f"БУ Сумма для ИД должна быть пустой, получено: {bu_sum}"
                assert not pu_sum, f"ПУ Сумма для ИД должна быть пустой, получено: {pu_sum}"
                return
        pytest.fail("Строка 'ИД' не найдена в Свод")

    def test_details_sheet_exists(self, temp_data_dir: Path, output_file: Path):
        """Лист «Детализация» присутствует."""
        _run_pipeline(temp_data_dir, output_file)
        wb = openpyxl.load_workbook(output_file)
        assert "Детализация" in wb.sheetnames

    def test_details_has_discrepancy_row(self, temp_data_dir: Path, output_file: Path):
        """«Детализация» содержит строку с расхождением по Страховому резерву."""
        _run_pipeline(temp_data_dir, output_file)
        wb = openpyxl.load_workbook(output_file)
        ws = wb["Детализация"]

        found = False
        for row in ws.iter_rows(min_row=2, values_only=True):
            indicator, date_str, bu_sum, pu_sum, diff = row
            if indicator == "Страховой резерв" and date_str == "17.02.2026":
                found = True
                assert diff is not None and diff != "", \
                    "Разница в строке Детализации должна быть непустой"
                break

        assert found, "В Детализации не найдена строка Страховой резерв / 17.02.2026"

    def test_pv_fl_discrepancy_in_details(self, temp_data_dir: Path, output_file: Path):
        """«Детализация» содержит строку с расхождением по Пенсионным взносам ФЛ."""
        _run_pipeline(temp_data_dir, output_file)
        wb = openpyxl.load_workbook(output_file)
        ws = wb["Детализация"]

        found = False
        for row in ws.iter_rows(min_row=2, values_only=True):
            indicator, date_str, bu_sum, pu_sum, diff = row
            if indicator == "Пенсионные взносы ФЛ" and date_str == "28.02.2026":
                found = True
                break

        assert found, "В Детализации не найдена строка Пенсионные взносы ФЛ / 28.02.2026"

    def test_missing_dirs_creates_them_and_exits_0(self, tmp_path: Path):
        """Если директории данных отсутствуют, pipeline создаёт их и возвращает 0."""
        from npf_recon.config import Config
        from npf_recon.pipeline import run

        nonexistent = tmp_path / "nonexistent_data"
        config = Config(
            pu_dir=nonexistent / "ПУ",
            bu_dir=nonexistent / "БУ",
            output_file=tmp_path / "out.xlsx",
            period_label="Тест",
            eps=0.005,
        )
        exit_code = run(config)
        assert exit_code == 0
        assert (nonexistent / "ПУ").is_dir()
        assert (nonexistent / "БУ").is_dir()
