#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Проверка устойчивости к «чужим» файлам заказчика.

Берёт те же данные, но:
  * переставляет колонки в обратном порядке;
  * добавляет лишние колонки (в начале и в конце);
  * переименовывает листы;
  * переносит таблицу на второй лист (первый — пустая заставка).

Результат агента должен совпасть с результатом на исходных файлах.

Запуск: python3 tests/test_header_robustness.py
"""

import hashlib
import json
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
SCRIPT = ROOT / "skill" / "ful-order-distribution" / "scripts" / "ful_distribution.py"
DATA = ROOT / "data"


def rewrite(source, target, extra_before, extra_after, sheet_name, second_sheet=False):
    wb_in = load_workbook(str(source), data_only=True)
    ws_in = wb_in.worksheets[0]
    rows = [list(r) for r in ws_in.iter_rows(values_only=True)]

    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Титул" if second_sheet else sheet_name
    if second_sheet:
        ws["A1"] = "Выгрузка подготовлена автоматически"
        ws = wb.create_sheet(sheet_name)
    else:
        ws = wb.active

    header = rows[0]
    data = rows[1:]
    order = list(range(len(header)))[::-1]
    new_header = [header[i] for i in order] + extra_after
    ws.append(extra_before + new_header)
    for row in data:
        values = [row[i] if i < len(row) else None for i in order]
        ws.append(["служебное"] * len(extra_before) + values + [None] * len(extra_after))
    wb.save(str(target))
    wb_in.close()


def canonical_hash(path):
    wb = load_workbook(str(path), data_only=True)
    digest = hashlib.sha256()
    for ws in wb.worksheets:
        digest.update(ws.title.encode("utf-8"))
        for row in ws.iter_rows(values_only=True):
            digest.update("|".join("" if v is None else str(v) for v in row).encode("utf-8"))
    wb.close()
    return digest.hexdigest()


def run(files, out):
    proc = subprocess.run(
        [sys.executable, str(SCRIPT), "--files"] + [str(f) for f in files] + ["--out", str(out)],
        capture_output=True, text=True,
    )
    if proc.returncode != 0:
        print(proc.stdout, proc.stderr)
        raise SystemExit("агент упал на файлах варианта")
    wb = load_workbook(str(Path(out) / "result.xlsx"), data_only=True)
    ws = wb["Новые строки Cloud"]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    return rows


def main():
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        rewrite(DATA / "ful_1c.xlsx", tmp / "a.xlsx", ["Служебная колонка"], ["Пустая колонка"],
                "FUL_2026_11", second_sheet=True)
        rewrite(DATA / "cloud_registry.xlsx", tmp / "b.xlsx", [], ["Менеджер"],
                "CLOUD_REGISTRY")
        rewrite(DATA / "staff_1c.xlsx", tmp / "c.xlsx", ["Подразделение 2"], [],
                "EMPLOYEES")

        base = run([DATA / "ful_1c.xlsx", DATA / "cloud_registry.xlsx", DATA / "staff_1c.xlsx"],
                   tmp / "out_base")
        variant = run([tmp / "c.xlsx", tmp / "a.xlsx", tmp / "b.xlsx"], tmp / "out_variant")

    same = base == variant
    print(json.dumps({"same_result": same, "rows": len(base) - 1}, ensure_ascii=False))
    if not same:
        for left, right in zip(base, variant):
            if left != right:
                print("эталонный вариант:", left)
                print("вариант заказчика:", right)
    return 0 if same else 1


if __name__ == "__main__":
    sys.exit(main())
