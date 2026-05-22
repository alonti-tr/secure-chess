"""JSON-Lines wire protocol over a raw TCP socket.

Framing: one JSON object per line, UTF-8 encoded, terminated by `\\n`. The full
message schema lives in `specs/002-secure-chess/contracts/wire-protocol.md`.
"""

from __future__ import annotations

import json
import socket
from typing import Any, Dict

from secure_chess.common.errors import ConnectionClosed, ProtocolError


CLIENT_MESSAGE_TYPES = (
    "register",
    "login",
    "play_human",
    "play_ai",
    "cancel_lobby",
    "move",
    "resign",
    "quit",
)

SERVER_MESSAGE_TYPES = (
    "ok",
    "error",
    "game_started",
    "game_state",
    "game_ended",
)

MESSAGE_TYPES = CLIENT_MESSAGE_TYPES + SERVER_MESSAGE_TYPES


def ok(**extra: Any) -> Dict[str, Any]:
    msg: Dict[str, Any] = {"type": "ok"}
    msg.update(extra)
    return msg


def error(code: str, message: str) -> Dict[str, Any]:
    return {"type": "error", "code": code, "message": message}


def send_message(sock: socket.socket, msg: Dict[str, Any]) -> None:
    line = (json.dumps(msg, separators=(",", ":")) + "\n").encode("utf-8")
    try:
        sock.sendall(line)
    except OSError as exc:
        raise ConnectionClosed(str(exc)) from exc


class _LineReader:
    """Buffers a TCP socket and yields one JSON line per `read_line` call.

    Stored as an attribute on the `Session`/client so partial reads survive
    across calls.
    """

    def __init__(self, sock: socket.socket) -> None:
        self._sock = sock
        self._buf = bytearray()

    def read_line(self) -> bytes:
        while b"\n" not in self._buf:
            try:
                chunk = self._sock.recv(4096)
            except OSError as exc:
                raise ConnectionClosed(str(exc)) from exc
            if not chunk:
                raise ConnectionClosed("peer closed the connection")
            self._buf.extend(chunk)
        idx = self._buf.index(b"\n")
        line = bytes(self._buf[:idx])
        del self._buf[: idx + 1]
        return line


def make_reader(sock: socket.socket) -> _LineReader:
    return _LineReader(sock)


def recv_message(reader: _LineReader) -> Dict[str, Any]:
    line = reader.read_line()
    try:
        msg = json.loads(line.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ProtocolError(f"malformed JSON line: {exc}") from exc
    if not isinstance(msg, dict) or "type" not in msg:
        raise ProtocolError("message is not a JSON object with a 'type' field")
    if msg["type"] not in MESSAGE_TYPES:
        raise ProtocolError(f"unknown message type: {msg['type']!r}")
    return msg
