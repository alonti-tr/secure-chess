"""T016 [US1] UserStore registration, persistence, authentication."""

from __future__ import annotations

import pytest

from secure_chess.common.errors import (
    DuplicateUserError,
    ValidationError,
    WeakPasswordError,
)
from secure_chess.common.user_store import UserStore


def test_register_writes_file(tmp_path):
    store = UserStore(tmp_path / "users.json")
    store.register("alice", "hunter2!")
    text = (tmp_path / "users.json").read_text(encoding="utf-8")
    assert "alice" in text
    assert "hunter2!" not in text


def test_persisted_file_has_no_plaintext(tmp_path):
    store = UserStore(tmp_path / "users.json")
    store.register("alice", "hunter2!")
    store.register("bob", "secretXYZ")
    text = (tmp_path / "users.json").read_text(encoding="utf-8")
    assert "hunter2!" not in text
    assert "secretXYZ" not in text


def test_duplicate_registration_rejected(tmp_path):
    store = UserStore(tmp_path / "users.json")
    store.register("alice", "hunter2!")
    with pytest.raises(DuplicateUserError):
        store.register("alice", "anotherpw")


def test_weak_password_rejected(tmp_path):
    store = UserStore(tmp_path / "users.json")
    with pytest.raises(WeakPasswordError):
        store.register("alice", "short")


def test_bad_username_rejected(tmp_path):
    store = UserStore(tmp_path / "users.json")
    with pytest.raises(ValidationError):
        store.register("a", "longenough123")
    with pytest.raises(ValidationError):
        store.register("has space", "longenough123")


def test_authenticate_success(tmp_path):
    store = UserStore(tmp_path / "users.json")
    store.register("alice", "hunter2!")
    account = store.authenticate("alice", "hunter2!")
    assert account is not None
    assert account.username == "alice"


def test_authenticate_wrong_password_returns_none(tmp_path):
    store = UserStore(tmp_path / "users.json")
    store.register("alice", "hunter2!")
    assert store.authenticate("alice", "wrong") is None


def test_authenticate_unknown_user_returns_none(tmp_path):
    store = UserStore(tmp_path / "users.json")
    assert store.authenticate("ghost", "anypassword") is None


def test_store_persists_across_instances(tmp_path):
    path = tmp_path / "users.json"
    first = UserStore(path)
    first.register("alice", "hunter2!")
    second = UserStore(path)
    account = second.authenticate("alice", "hunter2!")
    assert account is not None
    assert account.username == "alice"
