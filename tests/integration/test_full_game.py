from __future__ import annotations

import socket
import time

from secure_chess.common import protocol
from secure_chess.server.server import ChessServer


def _connect(addr):
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.settimeout(5.0)
    s.connect(addr)
    return s, protocol.make_reader(s)


def _recv(reader, n: int = 1):
    return [protocol.recv_message(reader) for _ in range(n)]


def _login_and_join(sock, reader, username, password):
    protocol.send_message(sock, {"type": "register", "username": username, "password": password})
    assert _recv(reader, 1)[0]["type"] == "ok"
    protocol.send_message(sock, {"type": "play_human"})
    assert _recv(reader, 1)[0]["type"] == "ok"


def test_two_clients_play_scholars_mate(tmp_path):
    server = ChessServer(host="127.0.0.1", port=0, data_dir=str(tmp_path))
    addr = server.start()
    try:
        time.sleep(0.05)
        sa, ra = _connect(addr)
        sb, rb = _connect(addr)
        _login_and_join(sa, ra, "alice", "hunter2!")
        _login_and_join(sb, rb, "bob", "s3cretpw")

        time.sleep(0.1)

        a_start = _recv(ra, 1)[0]
        b_start = _recv(rb, 1)[0]
        assert a_start["type"] == "game_started"
        assert b_start["type"] == "game_started"
        assert a_start["you_are"] == "white"
        assert b_start["you_are"] == "black"
        assert a_start["game_id"] == b_start["game_id"]

        moves = [
            ("white", "e2e4", "e7e5"),
            ("white", "f1c4", "b8c6"),
            ("white", "d1h5", "g8f6"),
            ("white", "h5f7", None),
        ]

        for i, (_side, white_move, black_move) in enumerate(moves):
            protocol.send_message(sa, {"type": "move", "uci": white_move})
            assert _recv(ra, 1)[0]["type"] == "ok"

            push_a = _recv(ra, 1)[0]
            push_b = _recv(rb, 1)[0]
            if i == len(moves) - 1:
                assert push_a["type"] == "game_ended"
                assert push_b["type"] == "game_ended"
                assert push_a["result"] == "white_wins"
                assert push_a["reason"] == "checkmate"
                break
            assert push_a["type"] == "game_state"
            assert push_b["type"] == "game_state"
            assert push_a["last_move"] == white_move

            protocol.send_message(sb, {"type": "move", "uci": black_move})
            assert _recv(rb, 1)[0]["type"] == "ok"
            push_a = _recv(ra, 1)[0]
            push_b = _recv(rb, 1)[0]
            assert push_a["type"] == "game_state"
            assert push_b["type"] == "game_state"
            assert push_a["last_move"] == black_move

        sa.close()
        sb.close()
    finally:
        server.stop()


def test_illegal_move_rejected(tmp_path):
    server = ChessServer(host="127.0.0.1", port=0, data_dir=str(tmp_path))
    addr = server.start()
    try:
        time.sleep(0.05)
        sa, ra = _connect(addr)
        sb, rb = _connect(addr)
        _login_and_join(sa, ra, "alice", "hunter2!")
        _login_and_join(sb, rb, "bob", "s3cretpw")
        time.sleep(0.1)
        _recv(ra, 1)
        _recv(rb, 1)

        protocol.send_message(sa, {"type": "move", "uci": "e2e5"})
        msg = _recv(ra, 1)[0]
        assert msg["type"] == "error"
        assert msg["code"] == "illegal_move"

        sa.close()
        sb.close()
    finally:
        server.stop()


def test_resign_ends_game(tmp_path):
    server = ChessServer(host="127.0.0.1", port=0, data_dir=str(tmp_path))
    addr = server.start()
    try:
        time.sleep(0.05)
        sa, ra = _connect(addr)
        sb, rb = _connect(addr)
        _login_and_join(sa, ra, "alice", "hunter2!")
        _login_and_join(sb, rb, "bob", "s3cretpw")
        time.sleep(0.1)
        _recv(ra, 1)
        _recv(rb, 1)

        protocol.send_message(sa, {"type": "resign"})
        assert _recv(ra, 1)[0]["type"] == "ok"
        end_a = _recv(ra, 1)[0]
        end_b = _recv(rb, 1)[0]
        assert end_a["type"] == "game_ended"
        assert end_a["result"] == "black_wins"
        assert end_a["reason"] == "resignation"
        assert end_b["type"] == "game_ended"

        sa.close()
        sb.close()
    finally:
        server.stop()
