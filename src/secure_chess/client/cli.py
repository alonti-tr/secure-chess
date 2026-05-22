from __future__ import annotations

import shlex
import socket
import sys
import threading
from typing import Any, Dict, Optional

from secure_chess.common import protocol
from secure_chess.common.errors import ConnectionClosed, ProtocolError


PIECE_TO_GLYPH = {
    "K": "K", "Q": "Q", "R": "R", "B": "B", "N": "N", "P": "P",
    "k": "k", "q": "q", "r": "r", "b": "b", "n": "n", "p": "p",
}


def render_board(fen_short: str) -> str:
    placement = fen_short.split()[0]
    lines: list[str] = []
    for rank_idx, row in enumerate(placement.split("/")):
        rank_label = 8 - rank_idx
        cells: list[str] = []
        for ch in row:
            if ch.isdigit():
                cells.extend(["."] * int(ch))
            else:
                cells.append(PIECE_TO_GLYPH.get(ch, ch))
        lines.append(f"{rank_label} {' '.join(cells)}")
    lines.append("  a b c d e f g h")
    return "\n".join(lines)


class ClientConnection:
    def __init__(self, host: str, port: int) -> None:
        self._sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self._sock.connect((host, port))
        self._reader = protocol.make_reader(self._sock)
        self._send_lock = threading.Lock()
        self._stop = threading.Event()
        self.you_are: Optional[str] = None
        self.opponent: Optional[str] = None
        self.last_board: Optional[str] = None
        self.game_id: Optional[str] = None

    def send(self, msg: Dict[str, Any]) -> None:
        with self._send_lock:
            protocol.send_message(self._sock, msg)

    def reader_loop(self) -> None:
        while not self._stop.is_set():
            try:
                msg = protocol.recv_message(self._reader)
            except ConnectionClosed:
                print("\n[connection closed by server]")
                self._stop.set()
                return
            except ProtocolError as exc:
                print(f"\n[protocol error from server: {exc}]")
                continue
            self._on_message(msg)

    def _on_message(self, msg: Dict[str, Any]) -> None:
        t = msg["type"]
        if t == "ok":
            extras = {k: v for k, v in msg.items() if k != "type"}
            if extras:
                print(f"ok  {extras}")
            else:
                print("ok")
        elif t == "error":
            print(f"error: {msg.get('code')} - {msg.get('message')}")
        elif t == "game_started":
            self.you_are = msg.get("you_are")
            self.opponent = msg.get("opponent")
            self.game_id = msg.get("game_id")
            self.last_board = msg.get("board")
            print()
            print("== game started ==")
            print(f"you are: {self.you_are}   opponent: {self.opponent}   to move: {msg.get('to_move')}")
            if self.last_board:
                print(render_board(self.last_board))
        elif t == "game_state":
            self.last_board = msg.get("board")
            print()
            print(f"last move: {msg.get('last_move')}   to move: {msg.get('to_move')}"
                  + ("   CHECK" if msg.get("in_check") else ""))
            if self.last_board:
                print(render_board(self.last_board))
        elif t == "game_ended":
            print()
            print(f"== game ended == result: {msg.get('result')}   reason: {msg.get('reason')}")
            self.you_are = None
            self.opponent = None
            self.last_board = None
            self.game_id = None
        else:
            print(f"[unknown push: {msg}]")

    def stop(self) -> None:
        self._stop.set()
        try:
            self._sock.shutdown(socket.SHUT_RDWR)
        except OSError:
            pass
        self._sock.close()


def _split_args(line: str) -> list[str]:
    try:
        return shlex.split(line)
    except ValueError:
        return line.split()


def run_cli(host: str, port: int) -> int:
    print(f"connected to {host}:{port}")
    conn = ClientConnection(host, port)
    reader_thread = threading.Thread(target=conn.reader_loop, daemon=True, name="client-reader")
    reader_thread.start()
    try:
        while not conn._stop.is_set():
            try:
                line = input("> ").strip()
            except (EOFError, KeyboardInterrupt):
                break
            if not line:
                continue
            parts = _split_args(line)
            cmd, *args = parts
            if cmd in {"quit", "exit"}:
                conn.send({"type": "quit"})
                break
            if cmd == "board":
                if conn.last_board:
                    print(render_board(conn.last_board))
                else:
                    print("(no board to show)")
                continue
            try:
                msg = _build_message(cmd, args)
            except ValueError as exc:
                print(f"error: {exc}")
                continue
            try:
                conn.send(msg)
            except (OSError, ConnectionClosed) as exc:
                print(f"[send failed: {exc}]")
                break
    finally:
        conn.stop()
    return 0


def _build_message(cmd: str, args: list[str]) -> Dict[str, Any]:
    if cmd == "register":
        if len(args) != 2:
            raise ValueError("usage: register <username> <password>")
        return {"type": "register", "username": args[0], "password": args[1]}
    if cmd == "login":
        if len(args) != 2:
            raise ValueError("usage: login <username> <password>")
        return {"type": "login", "username": args[0], "password": args[1]}
    if cmd == "play_human":
        return {"type": "play_human"}
    if cmd == "cancel_lobby":
        return {"type": "cancel_lobby"}
    if cmd == "play_ai":
        msg: Dict[str, Any] = {"type": "play_ai"}
        if len(args) >= 1:
            if args[0] not in ("white", "black"):
                raise ValueError("color must be 'white' or 'black'")
            msg["color"] = args[0]
        if len(args) >= 2:
            try:
                msg["depth"] = int(args[1])
            except ValueError:
                raise ValueError("depth must be an integer")
        return msg
    if cmd == "move":
        if len(args) != 1:
            raise ValueError("usage: move <uci>  e.g. move e2e4")
        return {"type": "move", "uci": args[0]}
    if cmd == "resign":
        return {"type": "resign"}
    raise ValueError(f"unknown command: {cmd!r}")
