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


def test_play_ai_full_session(tmp_path):
    server = ChessServer(host="127.0.0.1", port=0, data_dir=str(tmp_path), ai_depth=2)
    addr = server.start()
    try:
        time.sleep(0.05)
        s, r = _connect(addr)
        protocol.send_message(s, {"type": "register", "username": "carol", "password": "abcdefgh"})
        assert protocol.recv_message(r)["type"] == "ok"
        protocol.send_message(s, {"type": "play_ai", "color": "white", "depth": 2})
        start = protocol.recv_message(r)
        assert start["type"] == "game_started"
        assert start["you_are"] == "white"
        assert "AI(depth=" in start["opponent"]

        for human_move in ["e2e4", "g1f3", "f1c4"]:
            protocol.send_message(s, {"type": "move", "uci": human_move})
            ok = protocol.recv_message(r)
            assert ok["type"] == "ok", ok
            state_after_human = protocol.recv_message(r)
            assert state_after_human["type"] == "game_state"
            assert state_after_human["last_move"] == human_move
            state_after_ai = protocol.recv_message(r)
            assert state_after_ai["type"] == "game_state"

        s.close()
    finally:
        server.stop()
