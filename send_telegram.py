from __future__ import annotations

import os
from datetime import date
from pathlib import Path

import requests


BASE_DIR = Path(__file__).resolve().parent


def require_env(name: str) -> str:
    value = os.environ.get(name, "").strip()
    if not value:
        raise RuntimeError(f"Missing required environment variable: {name}")
    return value


def telegram_api_url(token: str, method: str) -> str:
    return f"https://api.telegram.org/bot{token}/{method}"


def send_message(token: str, chat_id: str, text: str) -> None:
    response = requests.post(
        telegram_api_url(token, "sendMessage"),
        data={"chat_id": chat_id, "text": text},
        timeout=30,
    )
    response.raise_for_status()


def send_document(token: str, chat_id: str, path: Path, caption: str) -> None:
    if not path.exists():
        return
    with path.open("rb") as file:
        response = requests.post(
            telegram_api_url(token, "sendDocument"),
            data={"chat_id": chat_id, "caption": caption},
            files={"document": (path.name, file)},
            timeout=120,
        )
    response.raise_for_status()


def main() -> int:
    token = require_env("TELEGRAM_BOT_TOKEN")
    chat_id = require_env("TELEGRAM_CHAT_ID")
    today = date.today().isoformat()
    outputs_dir = BASE_DIR / "outputs"

    daily_report = outputs_dir / f"job_matches_{today}.xlsx"
    history = outputs_dir / "job_history.xlsx"
    scan_log = outputs_dir / f"scan_log_{today}.json"

    send_message(token, chat_id, f"Job scan finished for {today}. Sending report files now.")
    send_document(token, chat_id, daily_report, "Daily job matches")
    send_document(token, chat_id, history, "Job history")
    send_document(token, chat_id, scan_log, "Scan log")
    print(f"Sent Telegram report to chat {chat_id}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
