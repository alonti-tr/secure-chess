from __future__ import annotations

from secure_chess.common.crypto import hash_password, verify_password


def test_hash_is_not_plaintext() -> None:
    h = hash_password("hunter2!")
    assert h != "hunter2!"
    assert "hunter2!" not in h


def test_hash_format() -> None:
    h = hash_password("hunter2!")
    assert h.startswith("$2b$") or h.startswith("$2a$")


def test_verify_correct_password() -> None:
    h = hash_password("hunter2!")
    assert verify_password("hunter2!", h) is True


def test_verify_wrong_password() -> None:
    h = hash_password("hunter2!")
    assert verify_password("wrong", h) is False


def test_two_hashes_differ_due_to_salt() -> None:
    a = hash_password("samepw1234")
    b = hash_password("samepw1234")
    assert a != b
    assert verify_password("samepw1234", a)
    assert verify_password("samepw1234", b)


def test_verify_with_garbage_hash() -> None:
    assert verify_password("anything", "not-a-bcrypt-hash") is False
