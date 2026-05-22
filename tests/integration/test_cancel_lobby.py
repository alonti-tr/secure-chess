"""Integration tests for the `cancel_lobby` wire-protocol message.

`cancel_lobby` lets an authenticated session leave the matchmaking queue
without disconnecting, so the player can immediately switch to `play_ai` or
just return to the lobby screen without re-logging in.
"""

from __future__ import annotations

import socket
import time

from secure_chess.common import protocol
from secure_chess.server.server import ChessServer


def _connect(addr):
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.settimeout(10.0)
    s.connect(addr)
    return s, protocol.make_reader(s)


def test_cancel_lobby_then_play_ai(tmp_path):
    server = ChessServer(host="127.0.0.1", port=0, data_dir=str(tmp_path), ai_depth=1)
    addr = server.start()
    try:
        time.sleep(0.05)
        s, r = _connect(addr)
        protocol.send_message(s, {"type": "register", "username": "dave", "password": "supersecret"})
        assert protocol.recv_message(r)["type"] == "ok"

        protocol.send_message(s, {"type": "play_human"})
        assert protocol.recv_message(r)["type"] == "ok"

        protocol.send_message(s, {"type": "cancel_lobby"})
        assert protocol.recv_message(r)["type"] == "ok"

        protocol.send_message(s, {"type": "play_ai", "color": "white", "depth": 1})
        start = protocol.recv_message(r)
        assert start["type"] == "game_started"
        assert start["you_are"] == "white"

        s.close()
    finally:
        server.stop()


def test_cancel_lobby_without_being_in_lobby_errors(tmp_path):
    server = ChessServer(host="127.0.0.1", port=0, data_dir=str(tmp_path), ai_depth=1)
    addr = server.start()
    try:
        time.sleep(0.05)
        s, r = _connect(addr)
        protocol.send_message(s, {"type": "register", "username": "eve", "password": "supersecret"})
        assert protocol.recv_message(r)["type"] == "ok"

        protocol.send_message(s, {"type": "cancel_lobby"})
        err = protocol.recv_message(r)
        assert err["type"] == "error"
        assert err["code"] == "bad_state"

        s.close()
    finally:
        server.stop()


def test_cancel_lobby_removes_session_from_queue(tmp_path):
    server = ChessServer(host="127.0.0.1", port=0, data_dir=str(tmp_path), ai_depth=1)
    addr = server.start()
    try:
        time.sleep(0.05)
        s_a, r_a = _connect(addr)
        s_b, r_b = _connect(addr)

        protocol.send_message(s_a, {"type": "register", "username": "frank", "password": "p455word"})
        assert protocol.recv_message(r_a)["type"] == "ok"
        protocol.send_message(s_b, {"type": "register", "username": "grace", "password": "p455word"})
        assert protocol.recv_message(r_b)["type"] == "ok"

        protocol.send_message(s_a, {"type": "play_human"})
        assert protocol.recv_message(r_a)["type"] == "ok"
        protocol.send_message(s_a, {"type": "cancel_lobby"})
        assert protocol.recv_message(r_a)["type"] == "ok"

        protocol.send_message(s_b, {"type": "play_human"})
        assert protocol.recv_message(r_b)["type"] == "ok"

        s_b.settimeout(0.4)
        try:
            msg = protocol.recv_message(r_b)
            assert False, f"unexpected message while solo in lobby: {msg}"
        except Exception:
            pass

        s_a.close()
        s_b.close()
    finally:
        server.stop()
