#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Распределение упаковочных листов (ФУЛ) между логистами и подготовка строк
реестра Cloud.

Детерминированный скрипт: все решения о статусе ФУЛ, группе и логисте
принимаются здесь, языковая модель только пересказывает результат.

Запуск:
    python3 ful_distribution.py --ful /path/ful.xlsx \
        --cloud /path/cloud.xlsx --staff /path/staff.xlsx --out /path/out

Зависимости: openpyxl (xlsx), xlrd (xls, опционально).
"""

import argparse
import json
import os
import re
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------- константы

UNDEF = "не определено"

GROUP_AIR = "Авиаперевозки"
GROUP_VED = "ВЭД"
GROUP_MULTI = "Мультимодальные перевозки"

CLOUD_COLUMNS = [
    "Заказ",
    "Поставщик",
    "Бренд",
    "Логист",
    "Статус",
    "Дата подтверждения",
    "Cut-off",
    "Дата получения УЛ",
    "Страна отправления",
    "Способ доставки",
    "Документы получены",
    "Комментарий",
]

# Синонимы заголовков. Ключ — каноническое поле, значения — варианты написания
# (сравниваются после нормализации: нижний регистр, ё->е, без пунктуации).
ALIASES = {
    "ful": {
        "ful_date": ["дата", "дата фул", "дата и время фул", "дата отправки фул",
                     "дата упаковочного листа", "дата поступления"],
        "cutoff": ["cut-off date", "cut off date", "cutoff", "cut off",
                   "cut-off", "катофф", "катоф дата"],
        "ful_number": ["номер", "номер фул", "номер упаковочного листа",
                       "номер листа", "№"],
        "ful_status": ["статус", "статус фул"],
        "comment": ["комментарий", "комментарии", "примечание"],
        "method": ["способ доставки", "способ отправки", "тип доставки",
                   "вид доставки", "тип перевозки"],
        "brand": ["бизнес единица", "бренд", "бизнес-единица", "торговая марка"],
        "logist_1c": ["логист", "логист 1с"],
        "supplier": ["контрагент", "поставщик", "контрагент поставщик"],
        "country": ["страна", "страна отправления", "страна отправитель"],
    },
    "cloud": {
        "order": ["заказ", "номер заказа", "заказ номер"],
        "supplier": ["поставщик", "контрагент"],
        "brand": ["бренд", "бизнес единица", "бизнес-единица"],
        "logist": ["логист", "логист cloud"],
        "cloud_status": ["статус", "состояние заказа"],
        "confirm_date": ["дата подтверждения", "подтвержден", "дата подтверждения заказа"],
        "cutoff": ["cut-off", "cut off", "cutoff", "cut-off date", "катофф"],
        "ul_date": ["дата получения ул", "дата получения упаковочного листа",
                    "получен ул", "дата получения фул"],
        "country": ["страна отправления", "страна"],
        "method": ["способ доставки", "способ отправки", "тип доставки"],
        "docs_date": ["документы получены", "дата получения документов",
                      "документы", "дата документов"],
        "comment": ["комментарий", "комментарии", "примечание"],
    },
    "staff": {
        "name": ["наименование", "фио", "сотрудник", "фамилия имя отчество"],
        "state": ["текущее состояние", "состояние", "статус", "статус сотрудника"],
        "position": ["должность", "позиция"],
        "code": ["код", "табельный номер", "таб номер", "таб. №", "код сотрудника"],
    },
}

REQUIRED = {
    "ful": ["ful_date", "ful_number", "comment", "method", "supplier", "country"],
    "cloud": ["order", "supplier", "logist", "cloud_status", "ul_date"],
    "staff": ["name", "state", "position"],
}

# Страны/территории, которые относятся к Китаю для правила ВЭД.
CHINA_KEYS = {
    "китай", "china", "cn", "chn", "кнр", "prc", "китайская народная республика",
    "гонконг", "hong kong", "hk", "hongkong", "hong kong sar", "китай гонконг",
    "макао", "macao", "macau",
}

METHOD_ALIASES = {
    "air": "air", "авиа": "air", "aviation": "air", "авиаперевозка": "air",
    "sea": "sea", "море": "sea", "морская": "sea", "ocean": "sea",
    "truck": "truck", "авто": "truck", "фура": "truck",
    "direct truck": "direct truck", "directtruck": "direct truck",
    "прямой авто": "direct truck", "прямая машина": "direct truck",
    "railway": "railway", "жд": "railway", "ж/д": "railway", "железная дорога": "railway",
    "rail": "railway",
}

POSITION_GROUPS = [
    ("авиа", GROUP_AIR),
    ("вэд", GROUP_VED),
    ("мультимодаль", GROUP_MULTI),
]

# ---------------------------------------------------------------- утилиты


def norm_text(value) -> str:
    """Нижний регистр, ё->е, без лишних пробелов."""
    if value is None:
        return ""
    s = str(value).replace("\u00a0", " ").replace("ё", "е").replace("Ё", "Е")
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s


def norm_header(value) -> str:
    """Заголовок для сопоставления: без пунктуации и пробелов по краям."""
    s = norm_text(value)
    s = re.sub(r"[«»\"'`]", " ", s)
    s = re.sub(r"[^0-9a-zа-я]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def parse_date(value):
    """Дата из datetime / Excel-сериала / строки. None, если распознать нельзя."""
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        # Excel serial date (1900 system)
        try:
            base = datetime(1899, 12, 30)
            return (base + __import__("datetime").timedelta(days=float(value))).date()
        except Exception:
            return None
    s = str(value).strip()
    if not s:
        return None
    s = s.split("T")[0].split(" ")[0]
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%d.%m.%y", "%d-%m-%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def fmt_date(value) -> str:
    d = parse_date(value)
    return d.strftime("%d.%m.%Y") if d else ""


def norm_method(value) -> str:
    key = norm_header(value)
    if not key:
        return ""
    if key in METHOD_ALIASES:
        return METHOD_ALIASES[key]
    for alias, canon in METHOD_ALIASES.items():
        if key.startswith(alias) or alias in key:
            return canon
    return key


def is_china(country) -> bool:
    key = norm_header(country)
    if not key:
        return False
    return key in CHINA_KEYS or "китай" in key or "china" in key or "hongkong" in key


def supplier_key(name) -> str:
    """Ключ поставщика: регистр, пробелы, кавычки, орг-форма и дефисы не важны."""
    s = norm_text(name)
    s = re.sub(r"[«»\"'`()\[\]]", " ", s)
    s = re.sub(r"[.,;:/\\\-–—_]+", " ", s)
    s = re.sub(r"\b(ооо|оао|зао|пао|ао|ип|llc|ltd|inc|co|gmbh|тов)\b", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


ORDER_RE = re.compile(r"[Ss]\s?0*\d{4,}")


def extract_order(comment) -> str:
    """Номер заказа из комментария «Номер: S0001234 от 01.11.2026»."""
    if comment is None:
        return ""
    s = str(comment).replace("\u00a0", " ")
    m = ORDER_RE.search(s)
    if m:
        return re.sub(r"\s+", "", m.group(0)).upper()
    # запасной вариант: «Номер: 12345» — только токен, содержащий цифры
    m = re.search(r"номер\s*[:\-]?\s*([A-Za-zА-Яа-я]{0,4}[-_ ]?\d{3,}[A-Za-z0-9\-_/]*)", s, re.IGNORECASE)
    if m:
        return re.sub(r"\s+", "", m.group(1)).upper()
    return ""


def order_key(order) -> str:
    return re.sub(r"\s+", "", str(order or "")).upper()


def norm_state(value) -> tuple:
    """(доступен, причина недоступности)."""
    s = norm_text(value)
    if not s:
        return False, "статус сотрудника не указан"
    if "увол" in s:
        return False, "уволен"
    if "отпуск" in s or "декрет" in s or "беременност" in s or "родам" in s:
        return False, "отпуск"
    if "больничн" in s or "нетрудоспособ" in s:
        return False, "больничный"
    if s.startswith("работ") or s in {"активен", "в строю"}:
        return True, ""
    return False, "недоступен ({})".format(s)


def position_group(position) -> str:
    p = norm_text(position)
    if "ассистент" in p or "помощник" in p:
        return ""
    if "руководител" in p or "начальник" in p or "глава" in p or "директор" in p:
        return ""
    for needle, group in POSITION_GROUPS:
        if needle in p:
            return group
    return ""


def initials_ok(short_initials: list, full_name: str) -> bool:
    parts = [p for p in re.split(r"\s+", full_name.replace(".", " ").strip()) if p]
    if len(parts) < 2:
        return False
    surname, first, *rest = parts
    if not short_initials:
        return True
    for idx, ch in enumerate(short_initials):
        target = first if idx == 0 else (rest[0] if rest else "")
        if not target:
            return False
        if target[0].lower() != ch:
            return False
    return True


# ---------------------------------------------------------------- чтение Excel


class Table:
    """Лист Excel: заголовки, карта канонических полей, строки."""

    def __init__(self, path, sheet_name, header_index, headers, mapping, rows):
        self.path = Path(path)
        self.sheet_name = sheet_name
        self.header_index = header_index
        self.headers = headers
        self.mapping = mapping
        self.rows = rows

    def value(self, row, field):
        idx = self.mapping.get(field)
        if idx is None or idx >= len(row):
            return None
        return row[idx]

    def has(self, field):
        return field in self.mapping


def read_sheets(path):
    """Список (имя листа, матрица значений) для xlsx/xls."""
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix == ".xls":
        try:
            import xlrd
        except ImportError:
            raise SystemExit("Для .xls нужен пакет xlrd: pip install xlrd")
        book = xlrd.open_workbook(str(path))
        out = []
        for sheet in book.sheets():
            rows = []
            for r in range(sheet.nrows):
                rows.append([sheet.cell_value(r, c) for c in range(sheet.ncols)])
            out.append((sheet.name, rows))
        return out

    wb = load_workbook(str(path), data_only=True, read_only=True)
    out = []
    for ws in wb.worksheets:
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        out.append((ws.title, rows))
    wb.close()
    return out


def map_headers(headers, role):
    """Канонические поля -> индекс колонки. Сопоставление по заголовку, не по номеру."""
    normalized = [norm_header(h) for h in headers]
    mapping = {}
    for field, variants in ALIASES[role].items():
        for variant in variants:
            v = norm_header(variant)
            if v in normalized:
                idx = normalized.index(v)
                if idx not in mapping.values():
                    mapping[field] = idx
                    break
        if field in mapping:
            continue
        for idx, h in enumerate(normalized):
            if idx in mapping.values():
                continue
            if any(norm_header(v) == h for v in variants):
                mapping[field] = idx
                break
    return mapping


def load_table(path, role) -> Table:
    """Ищет в книге лист и строку заголовков, наиболее подходящие под роль."""
    best = None
    for sheet_name, rows in read_sheets(path):
        for header_index, headers in enumerate(rows[:10]):
            mapping = map_headers(headers, role)
            score = len(mapping) + 3 * sum(1 for f in REQUIRED[role] if f in mapping)
            if best is None or score > best[0]:
                data_rows = rows[header_index + 1:]
                data_rows = [r for r in data_rows if any(v not in (None, "") for v in r)]
                best = (score, sheet_name, header_index, headers, mapping, data_rows)
    if best is None:
        raise SystemExit("Не найдено ни одного листа в файле {}".format(path))
    score, sheet_name, header_index, headers, mapping, data_rows = best
    missing = [f for f in REQUIRED[role] if f not in mapping]
    if missing:
        raise SystemExit(
            "В файле {} не найдены обязательные колонки: {}. Заголовки листа: {}".format(
                path, ", ".join(missing), ", ".join(str(h) for h in headers if h)
            )
        )
    return Table(path, sheet_name, header_index, headers, mapping, data_rows)


ROLE_TITLES = {"ful": "выгрузка ФУЛ", "cloud": "реестр Cloud", "staff": "сотрудники"}


def detect_role(path, used_roles):
    """Определяет, какой из трёх файлов перед нами, по заголовкам."""
    scores = {}
    for role in ("ful", "cloud", "staff"):
        if role in used_roles:
            continue
        try:
            table = load_table(path, role)
            scores[role] = len(table.mapping) + 3 * len(REQUIRED[role])
        except SystemExit:
            continue
    if not scores:
        return None
    return max(scores, key=lambda r: scores[r])


# ---------------------------------------------------------------- модель данных


@dataclass
class Employee:
    name: str
    state: str
    position: str
    code: str
    group: str
    available: bool
    reason: str

    @property
    def short(self) -> str:
        parts = [p for p in re.split(r"\s+", self.name.strip()) if p]
        if len(parts) >= 3:
            return "{} {}.{}.".format(parts[0], parts[1][0], parts[2][0])
        if len(parts) == 2:
            return "{} {}.{}".format(parts[0], parts[1][0], "")
        return self.name


class StaffBook:
    """Сотрудники 1С + сопоставление коротких имён из Cloud с ФИО."""

    def __init__(self, employees):
        self.employees = employees
        self.by_name = {e.name: e for e in employees}
        self.by_surname = defaultdict(list)
        for e in employees:
            self.by_surname[norm_text(e.name.split()[0])].append(e)
        self.cloud_display = {}  # каноническое ФИО -> как писали в Cloud

    def resolve(self, name):
        """Короткое имя из Cloud -> каноническое ФИО (или None)."""
        raw = str(name or "").strip()
        if not raw:
            return None
        if raw in self.by_name:
            self.cloud_display.setdefault(raw, raw)
            return raw
        cleaned = raw.replace(".", " ").replace(",", " ")
        parts = [p for p in re.split(r"\s+", cleaned) if p]
        if not parts:
            return None
        surname = norm_text(parts[0])
        initials = [p[0].lower() for p in parts[1:] if p]
        candidates = [e for e in self.by_surname.get(surname, []) if initials_ok(initials, e.name)]
        if len(candidates) == 1:
            full = candidates[0].name
            self.cloud_display.setdefault(full, raw)
            return full
        return None

    def display_for(self, full_name):
        return self.cloud_display.get(full_name, full_name)

    def available_in_group(self, group):
        return [e for e in self.employees if e.group == group and e.available]

    def group_members(self, group):
        return [e for e in self.employees if e.group == group]


@dataclass
class FulItem:
    number: str
    ful_date: object
    order: str
    supplier: str
    brand: str
    method: str
    country: str
    comment: str


# ---------------------------------------------------------------- основной расчёт


def build_staff(table: Table) -> StaffBook:
    employees = []
    for row in table.rows:
        name = str(table.value(row, "name") or "").strip()
        if not name:
            continue
        state = str(table.value(row, "state") or "").strip()
        position = str(table.value(row, "position") or "").strip()
        code = str(table.value(row, "code") or "").strip()
        available, reason = norm_state(state)
        group = position_group(position)
        if not group:
            available = False
            reason = reason or "должность не распределяет заказы ({})".format(norm_text(position))
        employees.append(Employee(name, state, position, code, group, available, reason))
    if not employees:
        raise SystemExit("Файл сотрудников пуст")
    return StaffBook(employees)


def cloud_context(table: Table, staff: StaffBook):
    """Заказы Cloud, загрузка логистов и история по поставщикам."""
    orders = {}
    load = Counter()
    load_raw = Counter()
    history = defaultdict(Counter)
    history_last = {}
    last_ul_date = None
    supplier_names = defaultdict(Counter)

    for row in table.rows:
        order = order_key(table.value(row, "order"))
        if not order:
            continue
        raw_logist = str(table.value(row, "logist") or "").strip()
        canonical = staff.resolve(raw_logist) if raw_logist else None
        status = norm_text(table.value(row, "cloud_status"))
        supplier = supplier_key(table.value(row, "supplier"))
        raw_supplier = re.sub(r"\s+", " ", str(table.value(row, "supplier") or "")).strip()
        ul_date = parse_date(table.value(row, "ul_date"))
        if supplier and raw_supplier:
            supplier_names[supplier][raw_supplier] += 1

        orders[order] = {
            "supplier": supplier,
            "supplier_raw": table.value(row, "supplier"),
            "logist_raw": raw_logist,
            "logist": canonical or raw_logist,
            "status": status,
            "comment": table.value(row, "comment"),
            "ul_date": ul_date,
            "row": row,
        }
        if status == "в работе":
            if canonical:
                load[canonical] += 1
            elif raw_logist:
                load_raw[raw_logist] += 1
        if supplier and raw_logist:
            key = canonical or raw_logist
            history[supplier][key] += 1
            prev = history_last.get((supplier, key))
            if ul_date and (prev is None or ul_date > prev):
                history_last[(supplier, key)] = ul_date
        if ul_date and (last_ul_date is None or ul_date > last_ul_date):
            last_ul_date = ul_date

    return {
        "orders": orders,
        "load": load,
        "load_raw": load_raw,
        "history": history,
        "history_last": history_last,
        "last_ul_date": last_ul_date,
        "supplier_names": supplier_names,
    }


def supplier_display(name, ctx, batch_names=None):
    """Как поставщик записан в Cloud; для нового — самый полный вариант из 1С."""
    key = supplier_key(name)
    variants = ctx["supplier_names"].get(key)
    if variants:
        return sorted(variants.items(), key=lambda kv: (-kv[1], kv[0]))[0][0]
    if batch_names and batch_names.get(key):
        return batch_names[key]
    return re.sub(r"\s+", " ", str(name or "")).strip()


COUNTRY_DISPLAY = {
    "china": "Китай", "cn": "Китай", "chn": "Китай", "кнр": "Китай", "prc": "Китай",
    "turkey": "Турция", "tr": "Турция",
    "india": "Индия", "in": "Индия",
    "vietnam": "Вьетнам", "vn": "Вьетнам",
    "hong kong": "Гонконг", "hongkong": "Гонконг", "hk": "Гонконг",
}


def country_display(value):
    raw = re.sub(r"\s+", " ", str(value or "")).strip()
    if not raw:
        return ""
    return COUNTRY_DISPLAY.get(norm_header(raw), raw)


def quoted(name):
    name = str(name or "").strip()
    return name if ("«" in name or '"' in name) else "«{}»".format(name)


def pick_logist(group, supplier, supplier_label, ctx, staff, load_state):
    """(логист, обоснование, причина_пустоты)."""
    candidates = staff.available_in_group(group)
    if not candidates:
        members = staff.group_members(group)
        detail = "; ".join("{} — {}".format(e.name, e.reason) for e in members) or "в 1С нет сотрудников этой группы"
        return UNDEF, "в группе «{}» нет доступных логистов ({})".format(group, detail), "нет доступных логистов в группе"

    hist = ctx["history"].get(supplier, Counter())
    ranked = sorted(hist.items(), key=lambda kv: (-kv[1],
                                                  -(ctx["history_last"].get((supplier, kv[0])) or date.min).toordinal(),
                                                  kv[0]))
    preferred, pref_count = (ranked[0] if ranked else ("", 0))
    preferred_emp = next((e for e in candidates if e.name == preferred), None)

    def load_of(name):
        return load_state.get(name, 0)

    if preferred_emp is not None:
        justification = (
            "поставщик {}: логист чаще всех работал с ним — {} заказ(ов) в истории Cloud".format(
                quoted(supplier_label), pref_count
            )
        )
        return preferred_emp.name, justification, ""

    # предпочтительный логист есть, но недоступен/в другой группе — берём наименее загруженного
    note = ""
    if preferred:
        emp = staff.by_name.get(preferred)
        if emp is not None:
            note = "; предпочтительный логист {} недоступен ({})".format(emp.name, emp.reason or "другая группа")
        else:
            note = "; предпочтительный логист «{}» не найден в списке сотрудников".format(preferred)

    ordered = sorted(candidates, key=lambda e: (load_of(e.name), e.name))
    chosen = ordered[0]
    loads_text = ", ".join("{}: {}".format(e.name.split()[0], load_of(e.name)) for e in ordered)
    justification = "наименее загружен в группе «{}» — {} заказ(ов) «в работе» ({}){}".format(
        group, load_of(chosen.name), loads_text, note
    )
    return chosen.name, justification, ""


def process(ful_table: Table, cloud_table: Table, staff_table: Table):
    staff = build_staff(staff_table)
    ctx = cloud_context(cloud_table, staff)
    load_state = Counter(ctx["load"])

    items = []
    for row in ful_table.rows:
        items.append(
            FulItem(
                number=str(ful_table.value(row, "ful_number") or "").strip(),
                ful_date=parse_date(ful_table.value(row, "ful_date")),
                order=extract_order(ful_table.value(row, "comment")),
                supplier=str(ful_table.value(row, "supplier") or "").strip(),
                brand=str(ful_table.value(row, "brand") or "").strip(),
                method=norm_method(ful_table.value(row, "method")),
                country=str(ful_table.value(row, "country") or "").strip(),
                comment=str(ful_table.value(row, "comment") or "").strip(),
            )
        )

    # порядок обработки: дата ФУЛ, затем номер листа — без случайностей
    items.sort(key=lambda i: (i.ful_date or date.max, i.number))

    # как показывать поставщика, которого ещё нет в Cloud: самый полный вариант из пачки
    def _supplier_score(value):
        text = re.sub(r"\s+", " ", str(value or "")).strip()
        score = 0
        if re.search(r"\b(ооо|оао|зао|пао|ао|ип|llc|ltd)\b", norm_text(text)):
            score += 2
        if "«" in text or '"' in text:
            score += 1
        return (score, len(text))

    batch_names = {}
    for it in items:
        key = supplier_key(it.supplier)
        if not key:
            continue
        if key not in batch_names or _supplier_score(it.supplier) > _supplier_score(batch_names[key]):
            batch_names[key] = re.sub(r"\s+", " ", it.supplier).strip()

    new_rows = []
    repeats = []
    gaps = []
    attention = []
    journal = []
    seen_in_batch = {}
    supplier_lock = {}

    def journal_add(item, kind, group, logist, reason, flag="", note=""):
        journal.append({
            "ФУЛ": item.number,
            "Дата ФУЛ": item.ful_date,
            "Заказ": item.order or UNDEF,
            "Поставщик": supplier_shown,
            "Тип": kind,
            "Группа": group,
            "Логист": logist,
            "Обоснование": reason,
            "Требует внимания": "да" if flag else "нет",
            "Примечание": note or flag,
        })

    def attention_add(item, reason, action):
        attention.append({
            "ФУЛ": item.number,
            "Заказ": item.order or UNDEF,
            "Поставщик": item.supplier,
            "Дата ФУЛ": item.ful_date,
            "Причина": reason,
            "Что сделать": action,
        })

    for item in items:
        supplier_shown = supplier_display(item.supplier, ctx, batch_names)

        # 1. нет номера заказа
        if not item.order:
            kind = "ошибка данных"
            reason = "в комментарии ФУЛ нет номера заказа"
            attention_add(item, reason, "запросить у поставщика корректный ФУЛ, строку в Cloud не создавать")
            journal_add(item, kind, UNDEF, UNDEF, reason, flag=reason)
            continue

        # 2. дубль внутри пачки
        if item.order in seen_in_batch:
            first = seen_in_batch[item.order]
            reason = "заказ {} уже есть в этой пачке (ФУЛ {}), повторно строку не создаём".format(
                item.order, first.number or "без номера"
            )
            attention_add(item, "дубль внутри пачки: " + reason,
                          "сверить ФУЛ с поставщиком; строку в Cloud создаёт только первый ФУЛ")
            journal_add(item, "дубль в пачке", UNDEF, UNDEF, reason, flag=reason)
            continue
        seen_in_batch[item.order] = item

        # 3. повторная отправка (заказ уже в Cloud)
        if item.order in ctx["orders"]:
            entry = ctx["orders"][item.order]
            comment = "Повторная отправка ФУЛа, {}".format(
                item.ful_date.strftime("%d.%m.%Y") if item.ful_date else "дата не определена"
            )
            repeats.append({
                "ФУЛ": item.number,
                "Заказ": item.order,
                "Дата ФУЛ": item.ful_date,
                "Поставщик": item.supplier,
                "Логист в Cloud": entry["logist_raw"],
                "Статус в Cloud": entry["status"],
                "Комментарий для Cloud": comment,
                "Что сделать": "дописать комментарий в существующую строку Cloud, новую строку не создавать",
            })
            journal_add(item, "повторная отправка", UNDEF, entry["logist_raw"] or UNDEF,
                        "заказ {} уже есть в реестре Cloud (статус: {})".format(
                            item.order, entry["status"] or UNDEF))
            continue

        # 4. неполные данные
        missing = []
        if not item.method:
            missing.append("способ доставки")
        if not item.country and item.method != "air":
            missing.append("страна отправления")
        if not item.supplier:
            missing.append("поставщик")
        if missing:
            reason = "не определено: нет данных — {}".format(", ".join(missing))
            attention_add(item, reason, "уточнить данные в 1С и внести заказ вручную")
            journal_add(item, "ошибка данных", UNDEF, UNDEF, reason, flag=reason)
            continue

        # 5. группа
        if item.method == "air":
            group = GROUP_AIR
            group_reason = "способ доставки air → авиаперевозки"
        elif is_china(item.country):
            group = GROUP_VED
            group_reason = "не air + страна отправления {} → ВЭД".format(country_display(item.country))
        else:
            group = GROUP_MULTI
            group_reason = "не air + страна отправления {} (не Китай) → мультимодальные перевозки".format(
                country_display(item.country))

        # 6. один поставщик в пачке — один логист
        skey = supplier_key(item.supplier)
        locked = supplier_lock.get(skey)
        flags = []
        logist = ""
        justification = ""

        if locked is not None and locked["group"] == group:
            logist = locked["logist"]
            if logist == UNDEF:
                justification = ("один поставщик — один логист: у поставщика {} в этой пачке "
                                 "уже нет доступного логиста".format(quoted(supplier_shown)))
                flags.append(("в группе «{}» нет доступных логистов".format(group),
                              "распределить вручную; в другую группу заказ не передаётся"))
            else:
                justification = ("один поставщик — один логист: {} в этой пачке уже "
                                 "распределён на {}".format(quoted(supplier_shown), staff.display_for(logist)))
        else:
            if locked is not None:
                flags.append((
                    "заказы поставщика {} попали в разные группы ({} и {}) — правило "
                    "«один поставщик — один логист» выполнить нельзя".format(
                        quoted(supplier_shown), locked["group"], group),
                    "проверить вручную: возможно, разные поставщики записаны одинаково"))
            logist, justification, _ = pick_logist(group, skey, supplier_shown, ctx, staff, load_state)
            if logist == UNDEF:
                flags.append(("в группе «{}» нет доступных логистов".format(group),
                              "распределить вручную; в другую группу заказ не передаётся"))

        if logist and logist != UNDEF:
            load_state[logist] += 1
        supplier_lock.setdefault(skey, {"group": group, "logist": logist})

        row = {
            "Заказ": item.order,
            "Поставщик": supplier_shown,
            "Бренд": item.brand or UNDEF,
            "Логист": staff.display_for(logist) if logist and logist != UNDEF else UNDEF,
            "Статус": "в работе",
            "Дата подтверждения": "",
            "Cut-off": ful_cutoff(ful_table, item),
            "Дата получения УЛ": item.ful_date,
            "Страна отправления": country_display(item.country) or UNDEF,
            "Способ доставки": item.method,
            "Документы получены": "",
            "Комментарий": "",
        }

        # 7. возможный пропуск: ФУЛ раньше последней записи Cloud, но заказа там нет
        if item.ful_date and ctx["last_ul_date"] and item.ful_date < ctx["last_ul_date"]:
            gaps.append({
                "ФУЛ": item.number,
                "Заказ": item.order,
                "Дата ФУЛ": item.ful_date,
                "Последняя дата получения УЛ в Cloud": ctx["last_ul_date"],
                "Поставщик": supplier_shown,
                "Причина": "заказа нет в Cloud, хотя ФУЛ пришёл раньше последней записи реестра",
                "Что сделать": "проверить, не потерялся ли заказ; строка подготовлена, но требует подтверждения",
            })
            flags.append((
                "возможный пропуск: ФУЛ от {} раньше последней записи Cloud ({})".format(
                    item.ful_date.strftime("%d.%m.%Y"), ctx["last_ul_date"].strftime("%d.%m.%Y")),
                "проверить заказ в 1С/Cloud перед переносом строки"))
            row["Комментарий"] = "Возможный пропуск: ФУЛ от {}".format(item.ful_date.strftime("%d.%m.%Y"))

        if not item.country and item.method == "air":
            flags.append(("не определено: нет страны отправления (для air группа определяется без неё)",
                          "уточнить страну и дописать в строку Cloud"))

        for reason, action in flags:
            attention_add(item, reason, action)

        new_rows.append(row)
        journal_add(item, "новый заказ", group,
                    staff.display_for(logist) if logist and logist != UNDEF else UNDEF,
                    "{}; {}".format(group_reason, justification),
                    flag="; ".join(reason for reason, _ in flags))

    new_rows.sort(key=lambda r: r["Заказ"])

    summary = {
        "ful_total": len(items),
        "new_rows": len(new_rows),
        "repeats": len(repeats),
        "gaps": len(gaps),
        "attention": len(attention),
        "duplicates": sum(1 for j in journal if j["Тип"] == "дубль в пачке"),
        "data_errors": sum(1 for j in journal if j["Тип"] == "ошибка данных"),
        "unassigned": sum(1 for r in new_rows if r["Логист"] == UNDEF),
    }
    return {
        "summary": summary,
        "new_rows": new_rows,
        "repeats": repeats,
        "gaps": gaps,
        "attention": attention,
        "journal": journal,
        "staff": staff,
        "ctx": ctx,
    }


def ful_cutoff(ful_table, item):
    """Cut-off берётся из строки ФУЛ как есть."""
    for row in ful_table.rows:
        number = str(ful_table.value(row, "ful_number") or "").strip()
        if number and number == item.number:
            return parse_date(ful_table.value(row, "cutoff")) or ful_table.value(row, "cutoff")
    return ""


# ---------------------------------------------------------------- вывод в Excel

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(bold=True, color="FFFFFF")
WARN_FILL = PatternFill("solid", fgColor="FFF2CC")
UNDEF_FILL = PatternFill("solid", fgColor="F8CBAD")


def style_header(ws, widths):
    for idx, width in enumerate(widths, start=1):
        cell = ws.cell(row=1, column=idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(idx)].width = width
    ws.freeze_panes = "A2"


def write_rows(ws, columns, rows, date_columns=(), warn_columns=()):
    ws.append(columns)
    for row in rows:
        ws.append([row.get(c) for c in columns])
    for row_idx in range(2, ws.max_row + 1):
        for col_idx, col_name in enumerate(columns, start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if col_name in date_columns:
                cell.number_format = "DD.MM.YYYY"
            if isinstance(cell.value, str) and cell.value == UNDEF:
                cell.fill = UNDEF_FILL
            elif col_name in warn_columns and cell.value not in (None, ""):
                cell.fill = WARN_FILL


def write_result(result, out_dir):
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / "result.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "Новые строки Cloud"
    write_rows(ws, CLOUD_COLUMNS, result["new_rows"], date_columns={"Дата получения УЛ", "Cut-off"},
               warn_columns={"Логист"})
    style_header(ws, [14, 26, 16, 26, 12, 16, 12, 16, 20, 16, 16, 34])
    ws.auto_filter.ref = "A1:L{}".format(max(ws.max_row, 1))

    ws = wb.create_sheet("Повторные отправки")
    cols = ["ФУЛ", "Заказ", "Дата ФУЛ", "Поставщик", "Логист в Cloud", "Статус в Cloud",
            "Комментарий для Cloud", "Что сделать"]
    write_rows(ws, cols, result["repeats"], date_columns={"Дата ФУЛ"})
    style_header(ws, [14, 14, 12, 26, 24, 12, 32, 46])

    ws = wb.create_sheet("Возможные пропуски")
    cols = ["ФУЛ", "Заказ", "Дата ФУЛ", "Последняя дата получения УЛ в Cloud",
            "Поставщик", "Причина", "Что сделать"]
    write_rows(ws, cols, result["gaps"], date_columns={"Дата ФУЛ", "Последняя дата получения УЛ в Cloud"})
    style_header(ws, [14, 14, 12, 22, 26, 50, 46])

    ws = wb.create_sheet("Требует внимания")
    cols = ["ФУЛ", "Заказ", "Поставщик", "Дата ФУЛ", "Причина", "Что сделать"]
    write_rows(ws, cols, result["attention"], date_columns={"Дата ФУЛ"})
    style_header(ws, [14, 14, 26, 12, 60, 46])

    ws = wb.create_sheet("Журнал решений")
    cols = ["ФУЛ", "Дата ФУЛ", "Заказ", "Поставщик", "Тип", "Группа", "Логист",
            "Обоснование", "Требует внимания", "Примечание"]
    write_rows(ws, cols, result["journal"], date_columns={"Дата ФУЛ"})
    style_header(ws, [14, 12, 14, 26, 18, 24, 26, 70, 16, 46])

    ws = wb.create_sheet("Сводка")
    s = result["summary"]
    ws.append(["Показатель", "Значение"])
    ws.append(["ФУЛ в пачке", s["ful_total"]])
    ws.append(["Новых строк Cloud", s["new_rows"]])
    ws.append(["Повторных отправок", s["repeats"]])
    ws.append(["Дублей внутри пачки", s["duplicates"]])
    ws.append(["Возможных пропусков", s["gaps"]])
    ws.append(["Требуют внимания", s["attention"]])
    ws.append(["Строк без логиста", s["unassigned"]])
    ws.append(["Ошибок данных", s["data_errors"]])
    style_header(ws, [34, 12])

    # фиксированные метаданные: результат побайтово одинаков при повторных запусках
    wb.properties.created = datetime(2026, 1, 1, 0, 0, 0)
    wb.properties.modified = datetime(2026, 1, 1, 0, 0, 0)
    wb.properties.creator = "ful-distribution"
    wb.properties.lastModifiedBy = "ful-distribution"
    wb.save(str(path))
    stabilize_xlsx(path)
    return path


def stabilize_xlsx(path):
    """Одинаковые байты при повторных запусках: фиксируем даты внутри zip."""
    import shutil
    import tempfile
    import zipfile

    path = Path(path)
    with zipfile.ZipFile(path, "r") as source:
        entries = [(info.filename, source.read(info.filename)) for info in source.infolist()]
    # openpyxl перезаписывает дату изменения при сохранении — фиксируем её в core.xml
    entries = [
        (name,
         re.sub(rb"(<dcterms:modified[^>]*>)[^<]+",
                rb"\g<1>2026-01-01T00:00:00Z", data) if name == "docProps/core.xml" else data)
        for name, data in entries
    ]
    fd, tmp_name = tempfile.mkstemp(suffix=".xlsx", dir=str(path.parent))
    os.close(fd)
    with zipfile.ZipFile(tmp_name, "w", zipfile.ZIP_DEFLATED) as target:
        for name, data in entries:
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o600 << 16
            target.writestr(info, data)
    shutil.move(tmp_name, path)


def chat_answer(summary, path):
    return (
        "Готово: новых строк — {new_rows}, повторных отправок — {repeats}, "
        "требуют внимания — {attention}.\nФайл: {path}".format(path=path, **summary)
    )


# ---------------------------------------------------------------- CLI


def main(argv=None):
    parser = argparse.ArgumentParser(description="Распределение ФУЛ между логистами")
    parser.add_argument("--ful", action="append", default=[], help="файл(ы) выгрузки ФУЛ (можно не указывать при --files)")
    parser.add_argument("--cloud", help="реестр Cloud")
    parser.add_argument("--staff", help="сотрудники из 1С")
    parser.add_argument("--files", nargs="*", default=[], help="любые входные файлы: роли определяются по заголовкам")
    parser.add_argument("--out", default="out", help="папка для result.xlsx")
    args = parser.parse_args(argv)

    paths = list(args.files)
    for key in ("ful", "cloud", "staff"):
        value = getattr(args, key)
        if value:
            paths.extend(value if isinstance(value, list) else [value])
    if len(paths) < 3:
        parser.error("нужно три файла: ФУЛ, реестр Cloud, сотрудники")

    roles = {}
    if args.cloud and args.staff and args.ful:
        roles = {"ful": args.ful[0], "cloud": args.cloud, "staff": args.staff}
    else:
        remaining = list(dict.fromkeys(paths))
        used = set()
        while remaining:
            best_role, best_path, best_score = None, None, -1
            for path in remaining:
                for role in ("ful", "cloud", "staff"):
                    if role in used:
                        continue
                    try:
                        table = load_table(path, role)
                    except SystemExit:
                        continue
                    score = len(table.mapping) + 3 * len(REQUIRED[role])
                    if score > best_score:
                        best_role, best_path, best_score = role, path, score
            if best_role is None:
                break
            roles[best_role] = best_path
            used.add(best_role)
            remaining.remove(best_path)
    missing = [r for r in ("ful", "cloud", "staff") if r not in roles]
    if missing:
        raise SystemExit("Не удалось определить файлы: {}".format(", ".join(ROLE_TITLES[r] for r in missing)))

    ful_table = load_table(roles["ful"], "ful")
    cloud_table = load_table(roles["cloud"], "cloud")
    staff_table = load_table(roles["staff"], "staff")

    result = process(ful_table, cloud_table, staff_table)
    path = write_result(result, args.out)

    answer = chat_answer(result["summary"], path)
    print(json.dumps({"summary": result["summary"], "file": str(path)}, ensure_ascii=False))
    print(answer)
    return 0


if __name__ == "__main__":
    sys.exit(main())
