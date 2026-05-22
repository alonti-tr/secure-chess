"""T017 [US1] Wire protocol framing for register/login messages."""

from __future__ import annotations

import socket

import pytest

from secure_chess.common import protocol
from secure_chess.common.errors import ProtocolError


def test_register_message_round_trip():
    a, b = socket.socketpair()
    try:
        protocol.send_message(a, {"type": "register", "username": "alice", "password": "hunter2!"})
        reader = protocol.make_reader(b)
        msg = protocol.recv_message(reader)
        assert msg == {"type": "register", "username": "alice", "password": "hunter2!"}
    finally:
        a.close()
        b.close()


def test_login_message_round_trip():
    a, b = socket.socketpair()
    try:
        protocol.send_message(a, {"type": "login", "username": "bob", "password": "s3cretpw"})
        reader = protocol.make_reader(b)
        msg = protocol.recv_message(reader)
        assert msg == {"type": "login", "username": "bob", "password": "s3cretpw"}
    finally:
        a.close()
        b.close()


def test_malformed_json_raises_protocol_error():
    a, b = socket.socketpair()
    try:
        a.sendall(b"not-json\n")
        reader = protocol.make_reader(b)
        with pytest.raises(ProtocolError):
            protocol.recv_message(reader)
    finally:
        a.close()
        b.close()


def test_unknown_message_type_raises_protocol_error():
    a, b = socket.socketpair()
    try:
        protocol.send_message(a, {"type": "unknown_thing"})
        reader = protocol.make_reader(b)
        with pytest.raises(ProtocolError):
            protocol.recv_message(reader)
    finally:
        a.close()
        b.close()


def test_ok_builder_shape():
    assert protocol.ok() == {"type": "ok"}
    assert protocol.ok(you_are="white") == {"type": "ok", "you_are": "white"}


def test_error_builder_shape():
    assert protocol.error("validation", "bad") == {
        "type": "error",
        "code": "validation",
        "message": "bad",
    }


def test_multiple_messages_back_to_back():
    a, b = socket.socketpair()
    try:
        protocol.send_message(a, {"type": "register", "username": "u1", "password": "pw11111111"})
        protocol.send_message(a, {"type": "login", "username": "u1", "password": "pw11111111"})
        reader = protocol.make_reader(b)
        m1 = protocol.recv_message(reader)
        m2 = protocol.recv_message(reader)
        assert m1["type"] == "register"
        assert m2["type"] == "login"
    finally:
        a.close()
        b.close()
