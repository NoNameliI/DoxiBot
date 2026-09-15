"""Временное хранилище присланных файлов, чтобы кнопки «Исправить»/«Примечания» работали без повторной загрузки."""

from __future__ import annotations

import secrets
import time
from dataclasses import dataclass, field


@dataclass
class StoredFile:
    data: bytes
    filename: str
    user_id: int
    created: float = field(default_factory=time.monotonic)


class FileStore:
    def __init__(self, ttl_minutes: int = 120, per_user: int = 5):
        self.ttl = ttl_minutes * 60
        self.per_user = per_user
        self._items: dict[str, StoredFile] = {}

    def put(self, data: bytes, filename: str, user_id: int) -> str:
        self._cleanup()
        own = sorted((k for k, v in self._items.items() if v.user_id == user_id), key=lambda k: self._items[k].created)
        for key in own[: max(0, len(own) - self.per_user + 1)]:
            del self._items[key]
        key = secrets.token_urlsafe(9)
        self._items[key] = StoredFile(data, filename, user_id)
        return key

    def get(self, key: str, user_id: int) -> StoredFile | None:
        self._cleanup()
        item = self._items.get(key)
        return item if item is not None and item.user_id == user_id else None

    def _cleanup(self) -> None:
        now = time.monotonic()
        for key in [k for k, v in self._items.items() if now - v.created > self.ttl]:
            del self._items[key]
