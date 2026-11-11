#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Контрольные прогоны агента и расчёт метрик приёмки.

Что делает:
  * запускает скрипт навыка N раз (по умолчанию 3) на одних и тех же файлах;
  * сравнивает результат с эталоном data/expected.xlsx по каждому ФУЛ;
  * считает точность по полям, долю полных совпадений, выдуманные значения;
  * проверяет форму реестра Cloud (12 колонок), отсутствие дублей и старых ФУЛ;
  * проверяет воспроизводимость (одинаковый результат между прогонами);
  * пишет out/metrics.json и печатает таблицу для отчёта.

Запуск:  python3 tests/run_checks.py --runs 3
"""

import argparse
import hashlib
import json
import subprocess
import sys
import time
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
SCRIPT = ROOT / "skill" / "ful-order-distribution" / "scripts" / "ful_distribution.py"
DATA = ROOT / "data"
OUT = ROOT / "out"

CLOUD_COLUMNS = [
    "Заказ", "Поставщик", "Бренд", "Логист", "Статус", "Дата подтверждения", "Cut-off",
    "Дата получения УЛ", "Страна отправления", "Способ доставки", "Документы получены",
    "Комментарий",
]

TYPE_MAP = {
    "новый заказ": "Новая строка",
    "повторная отправка": "Повторная отправка",
    "дубль в пачке": "Дубль в пачке",
    "ошибка данных": "Не создавать строку",
}

THRESHOLDS = {
    "full_match_share": 0.95,
    "type_accuracy": 1.0,
    "group_accuracy": 1.0,
    "logist_accuracy": 0.95,
    "comment_accuracy": 1.0,
    "fabricated_max": 0,
    "determinism": True,
    "form_ok": True,
    "max_seconds": 60.0,
}


def rows_of(path, sheet):
    wb = load_workbook(str(path), data_only=True)
    ws = wb[sheet]
    data = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    header = [str(c).strip() if c is not None else "" for c in data[0]]
    records = []
    for row in data[1:]:
        if all(v in (None, "") for v in row):
            continue
        records.append(dict(zip(header, row)))
    return header, records


def text(value):
    if value is None:
        return ""
    if hasattr(value, "strftime"):
        return value.strftime("%d.%m.%Y")
    return str(value).strip()


def load_expected():
    _, records = rows_of(DATA / "expected.xlsx", "Эталон")
    expected = {}
    for rec in records:
        expected[text(rec["ФУЛ"])] = {
            "order": text(rec["Заказ"]),
            "type": text(rec["Ожидаемый результат"]),
            "group": text(rec["Ожидаемая группа"]),
            "logist": text(rec["Ожидаемый логист"]),
            "comment": text(rec["Ожидаемый комментарий"]),
            "attention": text(rec["Требует внимания"]),
        }
    return expected


def load_actual(run_dir):
    _, journal = rows_of(run_dir / "result.xlsx", "Журнал решений")
    _, new_rows = rows_of(run_dir / "result.xlsx", "Новые строки Cloud")
    _, repeats = rows_of(run_dir / "result.xlsx", "Повторные отправки")
    _, attention = rows_of(run_dir / "result.xlsx", "Требует внимания")

    by_ful = {}
    for rec in journal:
        by_ful[text(rec["ФУЛ"])] = {
            "order": text(rec["Заказ"]),
            "type": TYPE_MAP.get(text(rec["Тип"]), text(rec["Тип"])),
            "group": text(rec["Группа"]),
            "logist": text(rec["Логист"]),
            "comment": "",
            "attention": text(rec["Требует внимания"]),
            "reason": text(rec["Обоснование"]),
        }
    for rec in new_rows:
        order = text(rec["Заказ"])
        for item in by_ful.values():
            if item["order"] == order and item["type"] == "Новая строка":
                item["comment"] = text(rec["Комментарий"])
    for rec in repeats:
        order = text(rec["Заказ"])
        for item in by_ful.values():
            if item["order"] == order and item["type"] == "Повторная отправка":
                item["comment"] = text(rec["Комментарий для Cloud"])
    return by_ful, new_rows, repeats, attention, journal


def canonical_hash(path):
    """Хэш содержимого всех листов: сравнение результатов между прогонами."""
    wb = load_workbook(str(path), data_only=True)
    digest = hashlib.sha256()
    for ws in wb.worksheets:
        digest.update(ws.title.encode("utf-8"))
        for row in ws.iter_rows(values_only=True):
            digest.update("|".join(text(v) for v in row).encode("utf-8"))
    wb.close()
    return digest.hexdigest()[:16]


def check_form(run_dir, cloud_orders):
    header, new_rows = rows_of(run_dir / "result.xlsx", "Новые строки Cloud")
    problems = []
    if header != CLOUD_COLUMNS:
        problems.append("порядок колонок не совпадает с формой Cloud: {}".format(header))
    seen = set()
    for rec in new_rows:
        order = text(rec["Заказ"])
        if not order:
            problems.append("строка без номера заказа")
        if order in seen:
            problems.append("дубль заказа в результате: {}".format(order))
        if order in cloud_orders:
            problems.append("в результат попал старый заказ из Cloud: {}".format(order))
        seen.add(order)
        if text(rec["Статус"]) != "в работе":
            problems.append("неверный статус у {}: {}".format(order, text(rec["Статус"])))
        if not text(rec["Логист"]):
            problems.append("пустой логист вместо «не определено» у {}".format(order))
        if not text(rec["Дата получения УЛ"]):
            problems.append("нет даты получения УЛ у {}".format(order))
    return problems


def main(argv=None):
    parser = argparse.ArgumentParser()
    parser.add_argument("--runs", type=int, default=3)
    args = parser.parse_args(argv)

    expected = load_expected()
    OUT.mkdir(exist_ok=True)

    _, cloud_rows = rows_of(DATA / "cloud_registry.xlsx", "Реестр заказов")
    cloud_orders = {text(r["Заказ"]).upper() for r in cloud_rows}

    times = []
    hashes = []
    first_actual = None
    form_problems = []
    run_summaries = []

    for run in range(1, args.runs + 1):
        run_dir = OUT / "run{}".format(run)
        started = time.perf_counter()
        proc = subprocess.run(
            [sys.executable, str(SCRIPT),
             "--ful", str(DATA / "ful_1c.xlsx"),
             "--cloud", str(DATA / "cloud_registry.xlsx"),
             "--staff", str(DATA / "staff_1c.xlsx"),
             "--out", str(run_dir)],
            capture_output=True, text=True,
        )
        elapsed = time.perf_counter() - started
        if proc.returncode != 0:
            print(proc.stdout, proc.stderr)
            raise SystemExit("прогон {} упал".format(run))
        times.append(round(elapsed, 2))
        hashes.append(canonical_hash(run_dir / "result.xlsx"))
        actual, new_rows, repeats, attention, journal = load_actual(run_dir)
        run_summaries.append(json.loads(proc.stdout.strip().splitlines()[0])["summary"])
        if first_actual is None:
            first_actual = actual
            form_problems = check_form(run_dir, cloud_orders)

    actual = first_actual
    fields = {"тип": (0, 0), "группа": (0, 0), "логист": (0, 0), "комментарий": (0, 0)}
    full_match = 0
    fabricated = 0
    mismatches = []

    for ful, exp in expected.items():
        got = actual.get(ful)
        if got is None:
            mismatches.append({"ФУЛ": ful, "поле": "журнал", "эталон": "строка", "агент": "нет"})
            continue
        local_ok = True
        checks = [
            ("тип", exp["type"], got["type"]),
            ("группа", exp["group"], got["group"]),
            ("логист", exp["logist"], got["logist"]),
            ("комментарий", exp["comment"], got["comment"]),
        ]
        for name, want, have in checks:
            if not want:
                continue
            fields[name] = (fields[name][0] + (1 if want == have else 0), fields[name][1] + 1)
            if want != have:
                local_ok = False
                mismatches.append({"ФУЛ": ful, "поле": name, "эталон": want, "агент": have})
                if want == "не определено" and have and have != "не определено":
                    fabricated += 1
                if name == "логист" and have and have != "не определено":
                    # конкретный логист там, где логиста быть не должно
                    fabricated += 1 if want == "не определено" else 0
        if local_ok:
            full_match += 1

    # ловушка «в группе нет логистов»: значение должно быть «не определено», не человек
    undefined_expected = [f for f, e in expected.items() if e["logist"] == "не определено"]
    invented_in_traps = sum(
        1 for f in undefined_expected
        if actual.get(f) and actual[f]["logist"] not in ("", "не определено")
    )

    accuracy = {k: (round(v[0] / v[1], 4) if v[1] else 1.0) for k, v in fields.items()}
    total = len(expected)
    full_share = round(full_match / total, 4)
    determinism = len(set(hashes)) == 1
    max_time = max(times)

    metrics = {
        "runs": args.runs,
        "ful_total": total,
        "seconds_per_run": times,
        "max_seconds": max_time,
        "full_match": full_match,
        "full_match_share": full_share,
        "field_accuracy": accuracy,
        "fabricated_values": fabricated,
        "invented_logist_in_empty_group": invented_in_traps,
        "deterministic": determinism,
        "result_hashes": hashes,
        "form_problems": form_problems,
        "run_summaries": run_summaries,
        "mismatches": mismatches,
    }

    passed = (
        full_share >= THRESHOLDS["full_match_share"]
        and accuracy["тип"] >= THRESHOLDS["type_accuracy"]
        and accuracy["группа"] >= THRESHOLDS["group_accuracy"]
        and accuracy["логист"] >= THRESHOLDS["logist_accuracy"]
        and accuracy["комментарий"] >= THRESHOLDS["comment_accuracy"]
        and fabricated <= THRESHOLDS["fabricated_max"]
        and determinism == THRESHOLDS["determinism"]
        and not form_problems
        and max_time <= THRESHOLDS["max_seconds"]
    )
    metrics["thresholds"] = THRESHOLDS
    metrics["passed"] = passed

    (OUT / "metrics.json").write_text(json.dumps(metrics, ensure_ascii=False, indent=2), encoding="utf-8")

    print("| Метрика | Значение | Порог |")
    print("| --- | --- | --- |")
    print("| Совпадение с эталоном по всем полям | {}/{} ({:.1%}) | ≥ 95% |".format(
        full_match, total, full_share))
    for key, label in [("тип", "Тип ФУЛ"), ("группа", "Группа"), ("логист", "Логист"),
                       ("комментарий", "Комментарий о повторной отправке")]:
        print("| Точность: {} | {:.1%} | {} |".format(
            label, accuracy[key], "100%" if key != "логист" else "≥ 95%"))
    print("| Выдуманные значения | {} | 0 |".format(fabricated))
    print("| Воспроизводимость ({} прогона) | {} | одинаковый результат |".format(
        args.runs, "да" if determinism else "нет"))
    print("| Форма Cloud (12 колонок, без дублей) | {} | без ошибок |".format(
        "ок" if not form_problems else "; ".join(form_problems)))
    print("| Время ответа | {} с | ≤ 60 с |".format(max_time))
    print("| Итог | {} | |".format("ПРИНЯТО" if passed else "НЕ ПРИНЯТО"))

    if mismatches:
        print("\nРасхождения:")
        for m in mismatches:
            print("  {ФУЛ}: {поле}: эталон={эталон!r}, агент={агент!r}".format(**m))

    return 0 if passed else 1


if __name__ == "__main__":
    sys.exit(main())
