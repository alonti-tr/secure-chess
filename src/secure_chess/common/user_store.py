"""JSON-backed credential store with bcrypt-hashed passwords.

The persistent file `users.json` is rewritten atomically (`write-to-tempfile +
os.replace`) on every registration. A single `threading.Lock` guards every
mutating method.
"""

from __future__ import annotations

import json
import os
import re
import tempfile
import threading
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, Optional

from secure_chess.common.crypto import hash_password, verify_password
from secure_chess.common.errors import (
    DuplicateUserError,
    ValidationError,
    WeakPasswordError,
)


USERNAME_RE = re.compile(r"^[A-Za-z0-9_-]{3,32}$")
MIN_PASSWORD_LEN = 8


@dataclass
class Account:
    username: str
    password_hash: str
    created_at: datetime


def _now_utc() -> datetime:
    return datetime.now(timezone.utc).replace(microsecond=0)


def _validate_username(username: str) -> None:
    if not isinstance(username, str):
        raise ValidationError("username must be a string")
    if not USERNAME_RE.match(username):
        raise ValidationError(
            "username must be 3-32 chars from [A-Za-z0-9_-]"
        )


def _validate_password(password: str) -> None:
    if not isinstance(password, str):
        raise ValidationError("password must be a string")
    if len(password) < MIN_PASSWORD_LEN:
        raise WeakPasswordError(
            f"password must be at least {MIN_PASSWORD_LEN} characters"
        )


class UserStore:
    def __init__(self, path: str | Path) -> None:
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._lock = threading.Lock()
        self._accounts: Dict[str, Account] = {}
        self._load()

    def _load(self) -> None:
        if not self.path.exists():
            self._persist_locked()
            return
        raw = self.path.read_text(encoding="utf-8") or "{}"
        try:
            data = json.loads(raw)
        except json.JSONDecodeError:
            data = {}
        if not isinstance(data, dict):
            data = {}
        accounts: Dict[str, Account] = {}
        for username, entry in data.items():
            if not isinstance(entry, dict):
                continue
            hashed = entry.get("password_hash")
            created_raw = entry.get("created_at")
            if not isinstance(hashed, str):
                continue
            try:
                created = datetime.fromisoformat(created_raw) if created_raw else _now_utc()
            except (TypeError, ValueError):
                created = _now_utc()
            accounts[username] = Account(
                username=username, password_hash=hashed, created_at=created
            )
        self._accounts = accounts

    def _persist_locked(self) -> None:
        serial = {
            a.username: {
                "password_hash": a.password_hash,
                "created_at": a.created_at.isoformat().replace("+00:00", "Z"),
            }
            for a in self._accounts.values()
        }
        text = json.dumps(serial, indent=2, sort_keys=True)
        with tempfile.NamedTemporaryFile(
            mode="w", encoding="utf-8", dir=self.path.parent,
            delete=False, prefix=".users.", suffix=".tmp"
        ) as tmp:
            tmp.write(text)
            tmp.flush()
            os.fsync(tmp.fileno())
            tmp_name = tmp.name
        os.replace(tmp_name, self.path)

    def __len__(self) -> int:
        with self._lock:
            return len(self._accounts)

    def exists(self, username: str) -> bool:
        with self._lock:
            return username in self._accounts

    def register(self, username: str, password: str) -> Account:
        _validate_username(username)
        _validate_password(password)
        hashed = hash_password(password)
        with self._lock:
            if username in self._accounts:
                raise DuplicateUserError(f"username {username!r} already exists")
            account = Account(
                username=username, password_hash=hashed, created_at=_now_utc()
            )
            self._accounts[username] = account
            self._persist_locked()
            return account

    def authenticate(self, username: str, password: str) -> Optional[Account]:
        if not isinstance(username, str) or not isinstance(password, str):
            return None
        with self._lock:
            account = self._accounts.get(username)
        if account is None:
            return None
        if verify_password(password, account.password_hash):
            return account
        return None
