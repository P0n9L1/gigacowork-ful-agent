#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Минимальный harness поверх GigaChat API (вариант из ТЗ: «скрипт навыка
запускается напрямую, а модель получает его вывод и пишет ответ пользователю»).

Что делает:
  1. запускает scripts/ful_distribution.py на приложенных файлах;
  2. отправляет GigaChat системный промпт агента + сводку скрипта;
  3. печатает ответ модели в чат (то, что увидит сотрудник).

Ключ авторизации GigaChat берётся из переменной окружения GIGACHAT_AUTH_KEY
или из файла ~/.gigachat_key. Корневой сертификат Минцифры — из
~/.config/gigachat/ca-bundle.pem (см. INSTALL.md, раздел «GigaChat API»).

Запуск:
    GIGACHAT_AUTH_KEY=<ключ> python3 tests/gigachat_run.py \
        --files data/ful_1c.xlsx data/cloud_registry.xlsx data/staff_1c.xlsx \
        --out out/gigachat --model GigaChat-2-Max
"""

import argparse
import json
import ssl
import subprocess
import sys
import time
import urllib.request
import uuid
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCRIPT = ROOT / "skill" / "ful-order-distribution" / "scripts" / "ful_distribution.py"
SKILL = ROOT / "skill" / "ful-order-distribution" / "SKILL.md"
SYSTEM_PROMPT = ROOT / "system_prompt.txt"

OAUTH_URL = "https://ngw.devices.sberbank.ru:9443/api/v2/oauth"
API_URL = "https://api.giga.chat/v1/chat/completions"
CA_BUNDLE = Path.home() / ".config/gigachat" / "ca-bundle.pem"


def auth_key():
    import os

    key = os.environ.get("GIGACHAT_AUTH_KEY", "").strip()
    if key:
        return key
    path = Path.home() / ".gigachat_key"
    if path.exists():
        return path.read_text().strip()
    raise SystemExit("не задан GIGACHAT_AUTH_KEY (или файл ~/.gigachat_key)")


def get_token(ctx):
    request = urllib.request.Request(
        OAUTH_URL,
        data=b"scope=GIGACHAT_API_PERS",
        headers={
            "Content-Type": "application/x-www-form-urlencoded",
            "Accept": "application/json",
            "RqUID": str(uuid.uuid4()),
            "Authorization": "Basic " + auth_key(),
        },
        method="POST",
    )
    with urllib.request.urlopen(request, context=ctx, timeout=40) as response:
        return json.loads(response.read().decode("utf-8"))["access_token"]


def chat(ctx, token, model, messages, temperature=0.0):
    payload = json.dumps({
        "model": model,
        "messages": messages,
        "temperature": temperature,
        "max_tokens": 600,
    }).encode("utf-8")
    request = urllib.request.Request(
        API_URL,
        data=payload,
        headers={
            "Content-Type": "application/json",
            "Accept": "application/json",
            "Authorization": "Bearer " + token,
        },
        method="POST",
    )
    with urllib.request.urlopen(request, context=ctx, timeout=90) as response:
        data = json.loads(response.read().decode("utf-8"))
    return data["choices"][0]["message"]["content"], data.get("usage", {})


def main(argv=None):
    parser = argparse.ArgumentParser()
    parser.add_argument("--files", nargs="+", required=True)
    parser.add_argument("--out", default="out/gigachat")
    parser.add_argument("--model", default="GigaChat-2-Max")
    args = parser.parse_args(argv)

    started = time.perf_counter()
    script_started = started
    run = subprocess.run(
        [sys.executable, str(SCRIPT), "--files"] + args.files + ["--out", args.out],
        capture_output=True, text=True,
    )
    if run.returncode != 0:
        print(run.stdout, run.stderr, file=sys.stderr)
        raise SystemExit("скрипт навыка завершился с ошибкой")
    script_elapsed = time.perf_counter() - script_started
    skill_json = run.stdout.strip().splitlines()[0]
    skill_summary = json.loads(skill_json)

    ctx = ssl.create_default_context(cafile=str(CA_BUNDLE)) if CA_BUNDLE.exists() else ssl.create_default_context()
    token = get_token(ctx)

    system = SYSTEM_PROMPT.read_text(encoding="utf-8")
    user = (
        "Пользователь приложил три файла Excel и отправил сообщение без текста.\n"
        "Скрипт навыка уже выполнен, его результат:\n"
        + json.dumps(skill_summary, ensure_ascii=False, indent=2)
        + "\nОтветь пользователю по шаблону из системного промпта. "
          "Не перечисляй строки, не выдумывай чисел."
    )

    answer, usage = chat(ctx, token, args.model, [
        {"role": "system", "content": system},
        {"role": "user", "content": user},
    ])
    elapsed = time.perf_counter() - started

    print("=" * 70)
    print("ОТВЕТ АГЕНТА В ЧАТЕ (модель {})".format(args.model))
    print("=" * 70)
    print(answer.strip())
    print("=" * 70)
    print("скрипт: {:.2f} с, всего с моделью: {:.2f} с, токены: {}".format(
        script_elapsed, elapsed, usage.get("total_tokens", "—")))
    print("файл результата:", skill_summary.get("file"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
