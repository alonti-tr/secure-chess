"""TCP accept loop for the secure-chess server.

`ChessServer` owns the listening socket and one shared `UserStore` / `Lobby` /
`GameRegistry`. Every accepted connection becomes a `Session` running on its
own daemon thread.
"""

from __future__ import annotations

import socket
import threading
from pathlib import Path
from typing import Dict, Optional

from secure_chess.common import protocol
from secure_chess.common.errors import ServerFullError
from secure_chess.common.log import get_logger
from secure_chess.common.user_store import UserStore
from secure_chess.server.lobby import GameRegistry, Lobby
from secure_chess.server.session import Session


log = get_logger("server")


class ChessServer:
    def __init__(
        self,
        host: str = "127.0.0.1",
        port: int = 5050,
        data_dir: str | Path = "data",
        ai_depth: int = 3,
        max_clients: int = 32,
    ) -> None:
        self.host = host
        self.port = port
        self.data_dir = Path(data_dir)
        self.data_dir.mkdir(parents=True, exist_ok=True)

        self.user_store = UserStore(self.data_dir / "users.json")
        self.lobby = Lobby()
        self.registry = GameRegistry()
        self.ai_depth = ai_depth
        self.max_clients = max_clients

        self._sock: Optional[socket.socket] = None
        self._stop_event = threading.Event()
        self._sessions: list[Session] = []
        self._sessions_lock = threading.Lock()
        self._active_logins: Dict[str, Session] = {}
        self._active_logins_lock = threading.Lock()
        self._accept_thread: Optional[threading.Thread] = None
        self.bound_address: tuple[str, int] | None = None

    def start(self) -> tuple[str, int]:
        self._sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self._sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        self._sock.bind((self.host, self.port))
        self._sock.listen(self.max_clients)
        self.bound_address = self._sock.getsockname()
        log.info("listening on %s:%d", *self.bound_address)
        log.info(
            "credential store: %s (%d accounts)", self.user_store.path, len(self.user_store)
        )
        self._accept_thread = threading.Thread(target=self._accept_loop, daemon=True)
        self._accept_thread.start()
        return self.bound_address

    def _accept_loop(self) -> None:
        assert self._sock is not None
        self._sock.settimeout(0.5)
        while not self._stop_event.is_set():
            try:
                client_sock, peer = self._sock.accept()
            except socket.timeout:
                continue
            except OSError:
                break

            with self._sessions_lock:
                live = [s for s in self._sessions if not getattr(s, "_closed", False)]
                self._sessions = live
                if len(live) >= self.max_clients:
                    log.warning("refusing connection from %s:%d - server full", *peer)
                    try:
                        protocol.send_message(
                            client_sock,
                            protocol.error(ServerFullError.code, "server is at capacity"),
                        )
                    except OSError:
                        pass
                    client_sock.close()
                    continue

            session = Session(
                sock=client_sock,
                peer=peer,
                user_store=self.user_store,
                lobby=self.lobby,
                registry=self.registry,
                active_logins=self._active_logins,
                active_logins_lock=self._active_logins_lock,
                ai_depth=self.ai_depth,
            )
            with self._sessions_lock:
                self._sessions.append(session)
            thread = threading.Thread(target=session.run, daemon=True, name=f"sess-{session.id[:6]}")
            thread.start()

    def stop(self) -> None:
        self._stop_event.set()
        if self._sock is not None:
            try:
                self._sock.close()
            except OSError:
                pass
        if self._accept_thread is not None:
            self._accept_thread.join(timeout=2.0)
        with self._sessions_lock:
            for s in self._sessions:
                s.close()
            self._sessions = []
        log.info("server stopped")
