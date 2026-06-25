"""
Подготовка ОБЕЗЛИЧЕННЫХ данных для внешней LLM.

Назначение модуля — собрать из результатов сверки компактное саммари БЕЗ какой
бы то ни было персональной информации и сформировать готовый промт, который во
внутреннем контуре (на отдельном сервере) передаётся LLM для генерации краткого
текстового описания отчёта и расхождений.

Что попадает в саммари (только агрегаты):
  — название показателя (это категория движения, не персональные данные);
  — наличие данных по сторонам БУ/ПУ;
  — итоги за период по БУ и ПУ;
  — факт, тип, сумма и даты расхождений;
  — для каждой даты расхождения: суммы БУ/ПУ и разница.

Что НЕ попадает (исключается на уровне модели данных Record и здесь):
  — ФИО, номера договоров, счета, имена файлов, любые строковые реквизиты.

`Record` по своей структуре не содержит персональных полей (только показатель,
дату, сумму и сторону), поэтому источник данных уже обезличен; этот модуль
дополнительно отдаёт наружу исключительно агрегаты.
"""
from __future__ import annotations

import datetime
import json
import logging
from pathlib import Path

from npf_recon.models import ReconRow

logger = logging.getLogger(__name__)

# Версия схемы саммари — на случай изменений формата на стороне потребителя.
SCHEMA_VERSION = "1.0"


def _round2(x: float | None) -> float | None:
    return None if x is None else round(float(x), 2)


def _classify(row: ReconRow, eps: float) -> tuple[bool, bool, str]:
    """
    Определяет статус сверки по строке.

    :return: (сверка_возможна, есть_расхождение, тип_расхождения)
    тип ∈ {«нет», «по_итогу», «только_по_дням», «по_итогу_и_по_дням», «сверка_невозможна»}
    """
    reconcilable = row.bu is not None and row.pu is not None
    if not reconcilable:
        return False, False, "сверка_невозможна"

    has_total = row.diff_total is not None and abs(row.diff_total) > eps
    has_days = bool(row.diff_dates)

    if has_total and has_days:
        kind = "по_итогу_и_по_дням"
    elif has_total:
        kind = "по_итогу"
    elif has_days:
        kind = "только_по_дням"
    else:
        kind = "нет"

    return True, (has_total or has_days), kind


def build_summary(
    rows: list[ReconRow],
    period_label: str,
    *,
    eps: float = 0.005,
    generated_at: datetime.date | None = None,
) -> dict:
    """
    Строит обезличенное саммари результатов сверки.

    :param rows:         результат reconcile() (14 строк).
    :param period_label: метка периода (напр. «Февраль 2026»).
    :param eps:          допуск сравнения сумм.
    :param generated_at: дата формирования (по умолчанию — сегодня).
    :return:             словарь, готовый к сериализации в JSON.
    """
    gen = generated_at or datetime.date.today()

    indicators_payload: list[dict] = []
    discrepancy_count = 0
    reconcilable_count = 0
    total_discrepancy_abs = 0.0

    for row in rows:
        reconcilable, has_discrep, kind = _classify(row, eps)
        if reconcilable:
            reconcilable_count += 1
        if has_discrep:
            discrepancy_count += 1

        days_payload: list[dict] = []
        for d in sorted(row.diff_dates):
            bu_val = row.bu.daily.get(d, 0.0) if row.bu else 0.0
            pu_val = row.pu.daily.get(d, 0.0) if row.pu else 0.0
            diff = row.diff_by_date.get(d)
            total_discrepancy_abs += abs(diff or 0.0)
            days_payload.append({
                "дата": d.strftime("%d.%m.%Y"),
                "сумма_БУ": _round2(bu_val),
                "сумма_ПУ": _round2(pu_val),
                "разница": _round2(diff),
            })

        indicators_payload.append({
            "показатель": row.indicator,
            "есть_данные_БУ": row.bu is not None,
            "есть_данные_ПУ": row.pu is not None,
            "итог_БУ": _round2(row.bu.total) if row.bu else None,
            "итог_ПУ": _round2(row.pu.total) if row.pu else None,
            "сверка_возможна": reconcilable,
            "расхождение_итог": _round2(row.diff_total) if reconcilable else None,
            "есть_расхождение": has_discrep,
            "тип_расхождения": kind,
            "дни_расхождений": days_payload,
        })

    return {
        "схема": SCHEMA_VERSION,
        "период": period_label,
        "дата_формирования": gen.strftime("%d.%m.%Y"),
        "обезличено": True,
        "итоги": {
            "всего_показателей": len(rows),
            "сверка_возможна_по": reconcilable_count,
            "показателей_с_расхождениями": discrepancy_count,
            "сумма_расхождений_по_модулю": _round2(total_discrepancy_abs),
            "есть_расхождения": discrepancy_count > 0,
        },
        "показатели": indicators_payload,
    }


# Системная инструкция для внешней LLM (русский, деловой стиль).
_PROMPT_INSTRUCTION = """\
Ты — финансовый аналитик негосударственного пенсионного фонда. На вход тебе
передаются ОБЕЗЛИЧЕННЫЕ итоги автоматической сверки бухгалтерского (БУ) и
персонифицированного (ПУ) учёта за период. В данных нет персональной
информации — только показатели (виды движений), даты и суммы.

Сформируй краткое деловое саммари на русском языке (5–10 предложений):
1) Укажи период и общий итог: сколько показателей сверено и по скольким
   обнаружены расхождения.
2) Перечисли показатели с расхождениями: для каждого назови тип
   («по итогу» / «только по дням» / «по итогу и по дням»), сумму расхождения
   и даты.
3) Отдельно отметь показатели, где месячные итоги БУ и ПУ совпали, но есть
   отличия по отдельным дням (тип «только по дням»).
4) Если расхождений нет — прямо сообщи, что сверка прошла без расхождений.
5) Заверши одной короткой рекомендацией (какие даты/показатели проверить).

Пиши по существу, без воды. Не добавляй данные, которых нет во входе,
и не делай числовых выводов сверх приведённых сумм.

ВХОДНЫЕ ДАННЫЕ (JSON):
"""


def build_prompt(summary: dict) -> str:
    """
    Собирает готовый промт: системная инструкция + JSON-контекст.

    :param summary: словарь из build_summary().
    :return:        строка-промт для передачи LLM.
    """
    payload = json.dumps(summary, ensure_ascii=False, indent=2)
    return _PROMPT_INSTRUCTION + payload + "\n"


def write_llm_export(
    rows: list[ReconRow],
    period_label: str,
    summary_path: Path,
    prompt_path: Path,
    *,
    eps: float = 0.005,
    generated_at: datetime.date | None = None,
) -> dict:
    """
    Формирует и записывает обезличенное саммари (JSON) и промт (текст).

    :return: построенное саммари (для повторного использования/логирования).
    """
    summary = build_summary(rows, period_label, eps=eps, generated_at=generated_at)

    summary_path.parent.mkdir(parents=True, exist_ok=True)
    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)

    prompt_path.parent.mkdir(parents=True, exist_ok=True)
    with open(prompt_path, "w", encoding="utf-8") as f:
        f.write(build_prompt(summary))

    logger.info("Обезличенное саммари сохранено: %s", summary_path)
    logger.info("Промт для LLM сохранён: %s", prompt_path)
    return summary
