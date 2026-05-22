from __future__ import annotations

import socket
import time

from secure_chess.common import protocol
from secure_chess.common.errors import ConnectionClosed
from secure_chess.server.server import ChessServer


def _connect(addr):
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.settimeout(5.0)
    s.connect(addr)
    return s, protocol.make_reader(s)


def _recv(reader, n: int = 1):
    return [protocol.recv_message(reader) for _ in range(n)]


def _try_recv(sock, reader, timeout: float = 0.5):
    sock.settimeout(timeout)
    try:
        return protocol.recv_message(reader)
    except (socket.timeout, OSError, ConnectionClosed):
        return None
    finally:
        sock.settimeout(5.0)


def _login_and_join(sock, reader, username, password):
    protocol.send_message(sock, {"type": "register", "username": username, "password": password})
    assert _recv(reader, 1)[0]["type"] == "ok"
    protocol.send_message(sock, {"type": "play_human"})
    assert _recv(reader, 1)[0]["type"] == "ok"


def test_two_parallel_games_are_isolated(tmp_path):
    server = ChessServer(host="127.0.0.1", port=0, data_dir=str(tmp_path))
    addr = server.start()
    try:
        time.sleep(0.05)
        connections = []
        readers = []
        users = [("alice", "passa1234"), ("bob", "passb1234"),
                 ("carol", "passc1234"), ("dave", "passd1234")]
        for name, pw in users:
            s, r = _connect(addr)
            connections.append(s)
            readers.append(r)
            _login_and_join(s, r, name, pw)
        time.sleep(0.2)

        starts = [_recv(r, 1)[0] for r in readers]
        for msg in starts:
            assert msg["type"] == "game_started"
        game_ids = {msg["game_id"] for msg in starts}
        assert len(game_ids) == 2, f"expected 2 distinct games, got {game_ids}"

        assert starts[0]["game_id"] == starts[1]["game_id"]
        assert starts[2]["game_id"] == starts[3]["game_id"]
        assert starts[0]["game_id"] != starts[2]["game_id"]

        protocol.send_message(connections[0], {"type": "move", "uci": "e2e4"})
        assert _recv(readers[0], 1)[0]["type"] == "ok"

        a_state = _recv(readers[0], 1)[0]
        b_state = _recv(readers[1], 1)[0]
        assert a_state["type"] == "game_state"
        assert b_state["type"] == "game_state"
        assert a_state["game_id"] == starts[0]["game_id"]
        assert a_state["last_move"] == "e2e4"

        c_msg = _try_recv(connections[2], readers[2], timeout=0.3)
        d_msg = _try_recv(connections[3], readers[3], timeout=0.3)
        assert c_msg is None, f"game B unexpectedly received: {c_msg}"
        assert d_msg is None, f"game B unexpectedly received: {d_msg}"

        assert len(server.registry.active_games()) == 2

        for s in connections:
            s.close()
    finally:
        server.stop()


def test_active_games_count_drops_on_resign(tmp_path):
    server = ChessServer(host="127.0.0.1", port=0, data_dir=str(tmp_path))
    addr = server.start()
    try:
        time.sleep(0.05)
        sa, ra = _connect(addr)
        sb, rb = _connect(addr)
        _login_and_join(sa, ra, "alice", "passa1234")
        _login_and_join(sb, rb, "bob", "passb1234")
        time.sleep(0.1)
        _recv(ra, 1); _recv(rb, 1)
        assert len(server.registry.active_games()) == 1

        protocol.send_message(sa, {"type": "resign"})
        _recv(ra, 1); _recv(ra, 1); _recv(rb, 1)
        time.sleep(0.1)
        assert len(server.registry.active_games()) == 0
        sa.close(); sb.close()
    finally:
        server.stop()
