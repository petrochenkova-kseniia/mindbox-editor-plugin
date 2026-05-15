#!/usr/bin/env python3
"""Получение данных из Yandex Wordstat через Yandex Cloud Search API.

Скрипт используется скиллом `wordstat` плагина mindbox-editor-plugin.

Запуск:
    python3 wordstat.py "поисковая фраза"
    python3 wordstat.py "фраза 1" "фраза 2" "фраза 3"

Аутентификация: IAM-токен, выпущенный через `yc iam create-token`.
Скрипт сам ищет yc CLI в стандартных местах. Если yc не установлен или
не авторизован — печатает в stderr понятную инструкцию и завершается с
ненулевым кодом, чтобы вызывающий скилл мог отреагировать.

Folder ID и Federation ID одинаковы для всей команды Mindbox.
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import urllib.error
import urllib.request
from pathlib import Path

# Каталог mindbox/dev в Yandex Cloud Mindbox. На него выдаётся роль
# search-api.executor федеративным учёткам авторов Журнала.
FOLDER_ID = "b1g0hbo5f2a37kl3bi8b"

# Федерация Mindbox для входа через корпоративный SSO.
FEDERATION_ID = "ajeaq0at6ev5p9je07tl"

# Регион «вся Россия» в кодировке Яндекса. Подробнее в getRegionsTree
# (но для Журнала нам этого региона хватает).
REGION_ALL_RUSSIA = "225"

# По 10 похожих формулировок на фразу — обычно с хвостом совпадает с
# тем, что Wordstat показывает в веб-интерфейсе.
NUM_PHRASES_PER_REQUEST = 10

ENDPOINT = "https://searchapi.api.cloud.yandex.net/v2/wordstat/topRequests"


def find_yc() -> str | None:
    """Найти бинарь yc в PATH или типичных местах установки на macOS."""
    in_path = shutil.which("yc")
    if in_path:
        return in_path

    home = str(Path.home())
    for candidate in (
        f"{home}/yandex-cloud/bin/yc",
        "/usr/local/bin/yc",
        "/opt/homebrew/bin/yc",
    ):
        if os.path.exists(candidate):
            return candidate
    return None


def get_iam_token(yc: str) -> str | None:
    """Получить свежий IAM-токен через yc CLI.

    Токен живёт ~12 часов; кэшировать на стороне скрипта не нужно —
    yc сам делает это в своём профиле, а перевыпуск дешёвый.
    Возвращает None, если пользователь не авторизован, чтобы вызывающий
    скилл мог попросить запустить yc init.
    """
    try:
        result = subprocess.run(
            [yc, "iam", "create-token"],
            capture_output=True,
            timeout=15,
        )
    except (subprocess.TimeoutExpired, OSError):
        return None

    if result.returncode != 0:
        return None

    token = result.stdout.decode("utf-8", errors="replace").strip()
    return token or None


def call_wordstat(token: str, phrase: str) -> dict:
    """Вызвать topRequests для одной фразы и нормализовать ответ.

    Wordstat возвращает счётчики строками — приводим к int, чтобы
    дальше с ними было удобно работать в промпте.
    """
    body = json.dumps(
        {
            "folderId": FOLDER_ID,
            "phrase": phrase,
            "numPhrases": NUM_PHRASES_PER_REQUEST,
            "regions": [REGION_ALL_RUSSIA],
            "devices": ["DEVICE_ALL"],
        },
        ensure_ascii=False,
    ).encode("utf-8")

    req = urllib.request.Request(
        ENDPOINT,
        data=body,
        method="POST",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
    )

    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = json.loads(resp.read())
    except urllib.error.HTTPError as err:
        err_body = err.read().decode("utf-8", errors="replace")
        return {
            "phrase": phrase,
            "error": {"http_code": err.code, "body": err_body},
        }
    except (urllib.error.URLError, TimeoutError, OSError) as err:
        return {"phrase": phrase, "error": {"message": str(err)}}

    def to_int(value) -> int:
        try:
            return int(value)
        except (TypeError, ValueError):
            return 0

    return {
        "phrase": phrase,
        "totalCount": to_int(data.get("totalCount")),
        "topRequests": [
            {"phrase": item.get("phrase", ""), "count": to_int(item.get("count"))}
            for item in data.get("results", [])
        ],
        "associations": [
            {"phrase": item.get("phrase", ""), "count": to_int(item.get("count"))}
            for item in data.get("associations", [])
        ],
    }


def die(msg: str, code: int = 1) -> None:
    print(msg, file=sys.stderr)
    sys.exit(code)


def main() -> None:
    if len(sys.argv) < 2:
        die(
            'Usage: wordstat.py "phrase 1" ["phrase 2" ...]\n'
            "Передай одну или несколько поисковых фраз аргументами.",
            code=2,
        )

    phrases = [arg.strip() for arg in sys.argv[1:] if arg.strip()]
    if not phrases:
        die("Не передано ни одной непустой фразы.", code=2)

    yc = find_yc()
    if yc is None:
        die(
            "yc CLI не найден.\n"
            "Установить:\n"
            "  curl -sSL https://storage.yandexcloud.net/yandexcloud-yc/install.sh | bash\n"
            "После установки выполни:\n"
            "  source ~/.zshrc\n"
            "и запусти скилл ещё раз."
        )

    token = get_iam_token(yc)
    if token is None:
        die(
            "yc CLI не авторизован.\n"
            "Запусти один раз:\n"
            f"  {yc} init --federation-id={FEDERATION_ID}\n"
            "Откроется браузер с логином Mindbox SSO. После успешного входа\n"
            "запусти скилл ещё раз — авторизация сохранится."
        )

    results = [call_wordstat(token, phrase) for phrase in phrases]

    # Если ни одна фраза не прошла из-за 403 — даём конкретную инструкцию,
    # а не голый код ошибки.
    for result in results:
        err = result.get("error") if isinstance(result, dict) else None
        if err and err.get("http_code") == 403:
            die(
                "HTTP 403: нет доступа к Wordstat.\n"
                "Напиши Ксюше Петроченковой (petrochenkova@mindbox.cloud) —\n"
                "она заведёт заявку в хелпдеск на выдачу доступа.\n"
                "Как только доступ выдадут, попробуй ещё раз —\n"
                "авторизация подхватится сама.",
                code=3,
            )

    # Для одной фразы — один объект, для нескольких — массив. Так удобнее
    # дальше парсить в промпте, не приходится разворачивать массив из одного.
    payload = results[0] if len(results) == 1 else results
    print(json.dumps(payload, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
