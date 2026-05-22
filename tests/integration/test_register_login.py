from __future__ import annotations

import socket
import time

from secure_chess.common import protocol
from secure_chess.server.server import ChessServer


def _connect(addr: tuple[str, int]) -> tuple[socket.socket, protocol._LineReader]:
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.settimeout(5.0)
    s.connect(addr)
    return s, protocol.make_reader(s)


def _wait_response(reader, max_msgs: int = 5):
    return protocol.recv_message(reader)


def test_register_then_restart_then_login(tmp_path):
    server = ChessServer(host="127.0.0.1", port=0, data_dir=str(tmp_path))
    addr = server.start()
    try:
        time.sleep(0.05)
        sock, reader = _connect(addr)
        protocol.send_message(sock, {"type": "register", "username": "alice", "password": "hunter2!"})
        msg = _wait_response(reader)
        assert msg["type"] == "ok", msg
        sock.close()
    finally:
        server.stop()

    users_file = tmp_path / "users.json"
    assert users_file.exists()
    assert "hunter2!" not in users_file.read_text(encoding="utf-8")

    server2 = ChessServer(host="127.0.0.1", port=0, data_dir=str(tmp_path))
    addr2 = server2.start()
    try:
        time.sleep(0.05)
        sock, reader = _connect(addr2)
        protocol.send_message(sock, {"type": "login", "username": "alice", "password": "hunter2!"})
        msg = _wait_response(reader)
        assert msg["type"] == "ok", msg
        sock.close()

        sock, reader = _connect(addr2)
        protocol.send_message(sock, {"type": "login", "username": "alice", "password": "wrong"})
        msg = _wait_response(reader)
        assert msg["type"] == "error"
        assert msg["code"] == "auth_failed"
        sock.close()

        sock, reader = _connect(addr2)
        protocol.send_message(sock, {"type": "login", "username": "ghost", "password": "anypassword"})
        msg = _wait_response(reader)
        assert msg["type"] == "error"
        assert msg["code"] == "auth_failed"
        sock.close()
    finally:
        server2.stop()


def test_duplicate_register_rejected_over_wire(tmp_path):
    server = ChessServer(host="127.0.0.1", port=0, data_dir=str(tmp_path))
    addr = server.start()
    try:
        time.sleep(0.05)
        s1, r1 = _connect(addr)
        protocol.send_message(s1, {"type": "register", "username": "bob", "password": "abcdefgh"})
        assert _wait_response(r1)["type"] == "ok"
        s1.close()

        s2, r2 = _connect(addr)
        protocol.send_message(s2, {"type": "register", "username": "bob", "password": "abcdefgh"})
        msg = _wait_response(r2)
        assert msg["type"] == "error"
        assert msg["code"] == "duplicate_user"
        s2.close()
    finally:
        server.stop()


def test_weak_password_rejected_over_wire(tmp_path):
    server = ChessServer(host="127.0.0.1", port=0, data_dir=str(tmp_path))
    addr = server.start()
    try:
        time.sleep(0.05)
        sock, reader = _connect(addr)
        protocol.send_message(sock, {"type": "register", "username": "alice", "password": "short"})
        msg = _wait_response(reader)
        assert msg["type"] == "error"
        assert msg["code"] == "weak_password"
        sock.close()
    finally:
        server.stop()
