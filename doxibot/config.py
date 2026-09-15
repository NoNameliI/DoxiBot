from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv

ROOT = Path(__file__).resolve().parents[1]


@dataclass(frozen=True)
class Config:
    token: str
    max_file_mb: int
    store_ttl_minutes: int
    workers: int


def load_config() -> Config:
    load_dotenv(ROOT / ".env")
    token = os.getenv("BOT_TOKEN", "").strip()
    if not token or token == "PASTE_YOUR_TOKEN_HERE":
        raise SystemExit(
            "Не задан токен бота. Создайте файл .env рядом с bot.py и впишите строку:\n"
            "BOT_TOKEN=123456:ABC-DEF...\n(токен выдаёт @BotFather)"
        )
    return Config(
        token=token,
        max_file_mb=int(os.getenv("MAX_FILE_MB", "20")),
        store_ttl_minutes=int(os.getenv("STORE_TTL_MINUTES", "120")),
        workers=int(os.getenv("WORKERS", "2")),
    )
