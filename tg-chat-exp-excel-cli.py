#!/usr/bin/env python3
"""
=== Telegram Chat Exporter to Excel file v1.1 CLI ===
Клиент Telegram для экспорта содеожимого чата в фвйл Excel за указанный период времени
Версия CLI - работает в консоли.
"""

import os
from datetime import datetime
import dotenv

dotenv.load_dotenv()

from core import PROGRAM_NAME, PROGRAM_VERSION, SESSION_NAME, list_chats, export_messages, check_env_vars

API_ID = os.getenv("API_ID")
API_HASH = os.getenv("API_HASH")
PHONE = os.getenv("PHONE")
YEAR_DEFAULT = int(os.getenv("YEAR_DEFAULT", "2025"))
MONTH_DEFAULT = int(os.getenv("MONTH_DEFAULT", "9"))

def cli_log(s: str):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"{ts} — {s}")

def run_cli():
    print(f"=== {PROGRAM_NAME} {PROGRAM_VERSION} — CLI ===")
    print("Вывод сообщений по топикам и выгрузка в Excel\n")

    # выбор чата
    cli_log("Загружаю список чатов для выбора...")
    chats = list_chats(API_ID, API_HASH, SESSION_NAME, PHONE, log_callback=cli_log)
    if not chats:
        cli_log("Не удалось получить список чатов. Завершение.")
        return
    for i, (name, cid) in enumerate(chats):
        print(f"{i+1}) {name}  ({cid})")
    sel = input("Выберите номер чата (Enter = 1): ").strip()
    idx = int(sel) - 1 if sel.isdigit() and 1 <= int(sel) <= len(chats) else 0
    chat_id = chats[idx][1]

    year_in = input(f"Введите год [YYYY] (Enter = {YEAR_DEFAULT}): ").strip()
    month_in = input(f"Введите месяц [1-12] (Enter = {MONTH_DEFAULT}): ").strip()
    year = int(year_in) if year_in.isdigit() else YEAR_DEFAULT
    month = int(month_in) if month_in.isdigit() else MONTH_DEFAULT

    cli_log(f"Запуск экспорта для чата {chat_id} за {year}-{month:02d}")
    res = export_messages(API_ID, API_HASH, SESSION_NAME, PHONE, chat_id, year, month, log_callback=cli_log)
    if res.get("success"):
        cli_log(f"Успех. Сохранён файл: {res['filename']} (сообщений: {res['count']})")
    else:
        cli_log(f"Завершено с сообщением: {res.get('message')}")

if __name__ == "__main__":
    ok, missing = check_env_vars()
    if not ok:
        print("❌ Ошибка: отсутствуют параметры в .env:", ", ".join(missing))
    else:
        run_cli()
