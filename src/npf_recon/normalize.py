"""
Нормализация дат и числовых значений из различных источников данных.

Функции этого модуля являются «клеем» между сырыми данными (Excel-ячейки,
строки JSON) и унифицированными типами Python, используемыми в бизнес-логике.
"""
from __future__ import annotations

import datetime
import re
from typing import Union

import pandas as pd

# Пробельные символы-разделители тысяч, встречающиеся в финансовых документах:
#     U+0020 — обычный пробел
#     U+00A0 — неразрывный пробел (NBSP)
#     U+202F — узкий неразрывный пробел (NARROW NBSP)
#     U+2009 — тонкий пробел (THIN SPACE)
_SPACE_CHARS = r"[    ]"

# Форматы дат, поддерживаемые parse_date
_DATE_FORMATS = [
    "%d.%m.%Y",   # 01.02.2026
    "%d.%m.%y",   # 01.02.26
    "%Y-%m-%d",   # 2026-02-01
]


def parse_number(value: object) -> float | None:
    """
    Разбирает числовое значение из различных форматов.

    Поддерживаемые форматы:
      - Целые и дробные числа (int/float) — возвращаются как есть.
      - «1234», «1 234», «1 234,56», «1234.56» — стандартные форматы.
      - Пробелы-разделители тысяч: обычный, неразрывный, узкий неразрывный.
      - Запятая как десятичный разделитель: «1 234,56» → 1234.56.
      - Когда в строке есть и точка, и запятая, запятая считается десятичной,
        а точки/пробелы — разделителями тысяч: «1.234,56» → 1234.56.
      - Отрицательные через скобки: «(1 234,56)» → -1234.56.
      - None, пустая строка, строка из пробелов → None.

    :param value: входное значение любого типа.
    :return: float или None.
    """
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        try:
            if pd.isna(value):
                return None
        except Exception:
            pass
        return float(value)

    s = str(value).strip()
    if not s:
        return None

    # Нормализуем спецсимволы минуса:
    #   U+2212 (MINUS SIGN), U+2012..U+2015 (различные тире) → обычный дефис-минус.
    s = re.sub(r"[−‒–—―]", "-", s)

    # Удаляем валютные обозначения (₽, «руб», «руб.», «р.», «RUB») — встречаются в 1С.
    s = re.sub(r"(?i)\s*(₽|руб\.?|р\.|rub)\s*$", "", s).strip()

    # Проверяем на NaN/nan
    if s.lower() in ("nan", "none", "-", ""):
        return None

    # Обнаруживаем знак минус:
    #   1) бухгалтерские скобки: «(1 234,56)» → отрицательное;
    #   2) завершающий минус: «1 234,56-» (выгрузка 1С) → отрицательное;
    #   3) ведущий «+» — просто отбрасываем.
    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative = True
        s = s[1:-1].strip()
    elif s.endswith("-"):
        negative = True
        s = s[:-1].strip()
    elif s.startswith("+"):
        s = s[1:].strip()

    # Убираем пробельные символы-разделители тысяч
    s_clean = re.sub(_SPACE_CHARS, "", s)

    # Если есть и точка, и запятая — определяем роли по позиции последнего вхождения:
    #   «1.234,56»  — последняя запятая ПОСЛЕ последней точки → запятая=дробная, точка=тысячи
    #   «1,234.56»  — последняя точка ПОСЛЕ последней запятой → точка=дробная, запятая=тысячи
    if "," in s_clean and "." in s_clean:
        last_dot = s_clean.rfind(".")
        last_comma = s_clean.rfind(",")
        if last_comma > last_dot:
            # Европейский формат: точки/пробелы = тысячи, запятая = дробная
            s_clean = s_clean.replace(".", "").replace(",", ".")
        else:
            # Американский формат: запятые = тысячи, точка = дробная
            s_clean = s_clean.replace(",", "")
    elif "," in s_clean:
        # Только запятая → может быть десятичной (если одна и <= 2 цифры после неё)
        parts = s_clean.split(",")
        if len(parts) == 2 and len(parts[1]) <= 2:
            s_clean = s_clean.replace(",", ".")
        else:
            # Иначе запятая — разделитель тысяч
            s_clean = s_clean.replace(",", "")
    # Иначе только точка или ничего — оставляем как есть

    try:
        result = float(s_clean)
    except ValueError:
        return None

    return -result if negative else result


def _is_nat(value: object) -> bool:
    """Проверяет, является ли значение pandas NaT без риска выброса TypeError."""
    try:
        return value is pd.NaT
    except Exception:
        return False


def parse_date(value: object) -> datetime.date | None:
    """
    Разбирает дату из различных форматов.

    Поддерживаемые форматы:
      - datetime.date, datetime.datetime — возвращаются как datetime.date.
      - pandas.Timestamp — конвертируется в datetime.date.
      - pandas.NaT — возвращает None.
      - ISO-строка с временем: «2026-02-01T00:00:00» и «2026-02-01T00:00:00Z».
      - Строки «DD.MM.YYYY», «DD.MM.YY», «YYYY-MM-DD».
      - None, пустая строка, непарсируемые значения → None.

    :param value: входное значение любого типа.
    :return: datetime.date или None.
    """
    if value is None:
        return None

    # pandas NaT — обрабатываем отдельно
    if _is_nat(value):
        return None

    # pandas Timestamp
    if isinstance(value, pd.Timestamp):
        try:
            if pd.isna(value):
                return None
        except Exception:
            pass
        return value.date()

    # datetime.datetime (должен быть перед datetime.date, т.к. является подклассом)
    if isinstance(value, datetime.datetime):
        return value.date()

    # datetime.date
    if isinstance(value, datetime.date):
        return value

    s = str(value).strip()
    if not s or s.lower() in ("nan", "none", "nat"):
        return None

    # Отрезаем временну̀ю часть ISO: «2026-02-01T00:00:00» → «2026-02-01»
    if "T" in s:
        s = s.split("T")[0]

    # Убираем суффикс Z, если остался
    s = s.rstrip("Z")

    for fmt in _DATE_FORMATS:
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            continue

    return None


def format_date(d: datetime.date | None) -> str:
    """
    Форматирует дату в строку «DD.MM.YYYY».

    :param d: объект datetime.date или None.
    :return: строка или пустая строка, если d is None.
    """
    if d is None:
        return ""
    return d.strftime("%d.%m.%Y")


def format_amount(x: float | None) -> str:
    """
    Форматирует сумму: обычный пробел (U+0020) как разделитель тысяч,
    запятая как десятичный разделитель, 2 знака после запятой.
    Отрицательные — в скобках.

    Примеры:
      1234.5        -> «1 234,50»
      -1234.5       -> «(1 234,50)»
      1234567.89    -> «1 234 567,89»
      0             -> «0,00»
      None          -> «»

    :param x: числовое значение или None.
    :return: отформатированная строка.
    """
    if x is None:
        return ""

    negative = x < 0
    abs_val = abs(x)

    # Python f"{n:,.2f}" дает тысячи через запятую, дробная — через точку:
    #   1234567.89 -> "1,234,567.89"
    # Переводим к нужному формату через промежуточный символ:
    raw = f"{abs_val:,.2f}"          # "1,234,567.89"
    raw = raw.replace(",", "\x00")   # "1\x00234\x00567.89"
    raw = raw.replace(".", ",")      # "1\x00234\x00567,89"
    formatted = raw.replace("\x00", " ")  # "1 234 567,89"  (обычный пробел U+0020)

    return f"({formatted})" if negative else formatted
