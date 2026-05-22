"""Smoke-test client used by quickstart validation (T053).

Connects to a running secure-chess server, registers a user, then plays one
move against the AI and prints what happens.
"""

from __future__ import annotations

import socket
import sys

from secure_chess.common import protocol


def main(host: str, port: int) -> int:
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.settimeout(10.0)
    s.connect((host, port))
    reader = protocol.make_reader(s)

    def send(msg):
        protocol.send_message(s, msg)

    def recv():
        return protocol.recv_message(reader)

    send({"type": "register", "username": "smoke_user", "password": "smokepass123"})
    print("register:", recv())

    send({"type": "play_ai", "color": "white", "depth": 2})
    print("game_started:", recv())

    send({"type": "move", "uci": "e2e4"})
    print("move ok:", recv())
    print("game_state (human move):", recv())
    print("game_state (AI move):", recv())

    send({"type": "resign"})
    print("resign ok:", recv())
    print("game_ended:", recv())

    s.close()
    return 0


if __name__ == "__main__":
    sys.exit(main("127.0.0.1", int(sys.argv[1]) if len(sys.argv) > 1 else 5151))
