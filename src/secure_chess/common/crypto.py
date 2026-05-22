"""Password hashing primitives.

Wraps `bcrypt` so the rest of the codebase never imports `bcrypt` directly.
`bcrypt` automatically generates and embeds a per-password salt in the returned
string, so callers don't need separate salt management.
"""

from __future__ import annotations

import bcrypt


def hash_password(plain: str) -> str:
    if not isinstance(plain, str):
        raise TypeError("password must be a string")
    return bcrypt.hashpw(plain.encode("utf-8"), bcrypt.gensalt()).decode("ascii")


def verify_password(plain: str, hashed: str) -> bool:
    if not isinstance(plain, str) or not isinstance(hashed, str):
        return False
    try:
        return bcrypt.checkpw(plain.encode("utf-8"), hashed.encode("ascii"))
    except (ValueError, TypeError):
        return False
