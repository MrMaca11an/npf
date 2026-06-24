"""
Тесты нормализации чисел и дат.
"""
from __future__ import annotations

import datetime

import pandas as pd
import pytest

from npf_recon.normalize import format_amount, format_date, parse_date, parse_number


# ------------------------------------------------------------------
# parse_number
# ------------------------------------------------------------------

class TestParseNumber:
    """Тесты функции parse_number."""

    def test_plain_integer(self):
        assert parse_number(1234) == 1234.0

    def test_plain_float(self):
        assert parse_number(1234.56) == pytest.approx(1234.56)

    def test_string_integer(self):
        assert parse_number("1234") == 1234.0

    def test_string_float_dot(self):
        assert parse_number("1234.56") == pytest.approx(1234.56)

    def test_string_float_comma(self):
        assert parse_number("1234,56") == pytest.approx(1234.56)

    def test_thousands_regular_space(self):
        assert parse_number("1 234") == 1234.0

    def test_thousands_nbsp(self):
        # Неразрывный пробел (U+00A0)
        assert parse_number("1 234") == 1234.0

    def test_thousands_narrow_nbsp(self):
        # Узкий неразрывный пробел (U+202F)
        assert parse_number("1 234") == 1234.0

    def test_thousands_and_comma_decimal(self):
        assert parse_number("1 234,56") == pytest.approx(1234.56)

    def test_thousands_nbsp_and_comma_decimal(self):
        assert parse_number("1 234,56") == pytest.approx(1234.56)

    def test_dot_thousands_comma_decimal(self):
        # Оба разделителя: точка = тысячи, запятая = дробная
        assert parse_number("1.234,56") == pytest.approx(1234.56)

    def test_parentheses_negative(self):
        assert parse_number("(1 234,56)") == pytest.approx(-1234.56)

    def test_parentheses_negative_no_spaces(self):
        assert parse_number("(1234.56)") == pytest.approx(-1234.56)

    def test_negative_sign(self):
        # Обычный минус через float-конверсию
        assert parse_number("-500.0") == pytest.approx(-500.0)

    def test_none_returns_none(self):
        assert parse_number(None) is None

    def test_empty_string(self):
        assert parse_number("") is None

    def test_blank_string(self):
        assert parse_number("   ") is None

    def test_nan_float(self):
        import math
        assert parse_number(float("nan")) is None

    def test_nan_string(self):
        assert parse_number("nan") is None

    def test_zero(self):
        assert parse_number(0) == 0.0

    def test_large_number_with_spaces(self):
        assert parse_number("1 234 567,89") == pytest.approx(1_234_567.89)

    def test_bool_returns_none(self):
        assert parse_number(True) is None
        assert parse_number(False) is None

    def test_large_plain_string(self):
        assert parse_number("312455.75") == pytest.approx(312_455.75)

    def test_formatted_amount_string(self):
        # Формат format_amount: «312 455,75» должен быть разобран обратно
        assert parse_number("312 455,75") == pytest.approx(312_455.75)

    # --- Расширенные форматы (умная нормализация) ---

    def test_unicode_minus(self):
        # U+2212 MINUS SIGN
        assert parse_number("−500") == pytest.approx(-500.0)

    def test_unicode_minus_with_decimals(self):
        assert parse_number("−1 234,56") == pytest.approx(-1234.56)

    def test_trailing_minus_accounting(self):
        # Завершающий минус (выгрузка 1С): «1 234,56-» → отрицательное
        assert parse_number("1 234,56-") == pytest.approx(-1234.56)

    def test_trailing_minus_plain(self):
        assert parse_number("1500.00-") == pytest.approx(-1500.0)

    def test_leading_plus(self):
        assert parse_number("+1000") == pytest.approx(1000.0)

    def test_currency_ruble_symbol(self):
        assert parse_number("1 234,56 ₽") == pytest.approx(1234.56)

    def test_currency_rub_word(self):
        assert parse_number("1 234,56 руб.") == pytest.approx(1234.56)

    def test_currency_rub_latin(self):
        assert parse_number("1000 RUB") == pytest.approx(1000.0)

    def test_en_dash_is_empty(self):
        # Прочерк «–» в ячейке означает отсутствие значения
        assert parse_number("–") is None

    def test_comma_thousands_three_digits(self):
        # «12,345» — три цифры после запятой → разделитель тысяч
        assert parse_number("12,345") == pytest.approx(12_345.0)

    def test_comma_decimal_two_digits(self):
        # «12,34» — две цифры → десятичный разделитель
        assert parse_number("12,34") == pytest.approx(12.34)

    def test_american_grouping(self):
        # «100,000.00» — запятые тысячи, точка дробная
        assert parse_number("100,000.00") == pytest.approx(100_000.0)

    def test_paren_negative_american(self):
        assert parse_number("(100,000.00)") == pytest.approx(-100_000.0)


# ------------------------------------------------------------------
# parse_date
# ------------------------------------------------------------------

class TestParseDate:
    """Тесты функции parse_date."""

    def test_dd_mm_yyyy(self):
        assert parse_date("01.02.2026") == datetime.date(2026, 2, 1)

    def test_dd_mm_yy(self):
        assert parse_date("01.02.26") == datetime.date(2026, 2, 1)

    def test_yyyy_mm_dd(self):
        assert parse_date("2026-02-01") == datetime.date(2026, 2, 1)

    def test_iso_datetime_string(self):
        assert parse_date("2026-02-01T00:00:00") == datetime.date(2026, 2, 1)

    def test_iso_datetime_with_z(self):
        assert parse_date("2026-02-01T00:00:00Z") == datetime.date(2026, 2, 1)

    def test_python_date(self):
        d = datetime.date(2026, 2, 17)
        assert parse_date(d) == d

    def test_python_datetime(self):
        dt = datetime.datetime(2026, 2, 17, 12, 30, 0)
        assert parse_date(dt) == datetime.date(2026, 2, 17)

    def test_pandas_timestamp(self):
        ts = pd.Timestamp("2026-02-17")
        assert parse_date(ts) == datetime.date(2026, 2, 17)

    def test_none_returns_none(self):
        assert parse_date(None) is None

    def test_empty_string(self):
        assert parse_date("") is None

    def test_nan_string(self):
        assert parse_date("nan") is None

    def test_pandas_nat(self):
        assert parse_date(pd.NaT) is None

    def test_unparseable_string(self):
        assert parse_date("не дата") is None

    def test_header_text(self):
        # Типичная шапка Excel — должна возвращать None
        assert parse_date("Дата операции") is None

    def test_dd_mm_yyyy_last_day(self):
        assert parse_date("28.02.2026") == datetime.date(2026, 2, 28)


# ------------------------------------------------------------------
# format_date
# ------------------------------------------------------------------

class TestFormatDate:
    def test_basic(self):
        assert format_date(datetime.date(2026, 2, 1)) == "01.02.2026"

    def test_none(self):
        assert format_date(None) == ""

    def test_last_day(self):
        assert format_date(datetime.date(2026, 2, 28)) == "28.02.2026"


# ------------------------------------------------------------------
# format_amount
# ------------------------------------------------------------------

class TestFormatAmount:
    def test_positive(self):
        assert format_amount(1234.5) == "1 234,50"

    def test_negative(self):
        assert format_amount(-1234.5) == "(1 234,50)"

    def test_zero(self):
        assert format_amount(0.0) == "0,00"

    def test_large(self):
        assert format_amount(1_234_567.89) == "1 234 567,89"

    def test_none(self):
        assert format_amount(None) == ""

    def test_roundtrip(self):
        """format_amount → parse_number должен вернуть исходное значение."""
        val = 312_455.75
        formatted = format_amount(val)
        assert parse_number(formatted) == pytest.approx(val)

    def test_negative_roundtrip(self):
        val = -15_000.0
        formatted = format_amount(val)
        assert parse_number(formatted) == pytest.approx(val)

    def test_large_negative(self):
        assert format_amount(-1_234_567.89) == "(1 234 567,89)"

    def test_small_value(self):
        assert format_amount(5.0) == "5,00"

    def test_trailing_minus_roundtrip(self):
        # Бухгалтерский завершающий минус разбирается обратно
        assert parse_number("1 234 567,89-") == pytest.approx(-1_234_567.89)
