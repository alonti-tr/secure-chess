"""Per-client server-side `Session` and the `SessionState` machine.

Each accepted TCP connection runs in its own thread driving `Session.run()`. The
handlers `handle_<type>` are added incrementally per user story (auth in US1,
game in US2, AI in US4).
"""

from __future__ import annotations

import socket
import threading
import uuid
from enum import Enum
from typing import TYPE_CHECKING, Any, Callable, Dict, Optional

from secure_chess.common import protocol
from secure_chess.common.errors import (
    AlreadyLoggedInError,
    AuthenticationError,
    BadStateError,
    ConnectionClosed,
    DuplicateUserError,
    IllegalMoveError,
    MoveParseError,
    NotYourTurnError,
    ProtocolError,
    SecureChessError,
    ValidationError,
    WeakPasswordError,
)
from secure_chess.common.log import get_logger


if TYPE_CHECKING:
    from secure_chess.common.user_store import Account, UserStore
    from secure_chess.common.game import Game
    from secure_chess.server.lobby import GameRegistry, Lobby


log = get_logger("session")


class SessionState(Enum):
    ANONYMOUS = "anonymous"
    AUTHENTICATED = "authenticated"
    IN_LOBBY = "in_lobby"
    IN_GAME = "in_game"


class Session:
    """One connected client. Owns its socket; survives as long as the TCP connection."""

    def __init__(
        self,
        sock: socket.socket,
        peer: tuple[str, int],
        user_store: "UserStore",
        lobby: "Lobby",
        registry: "GameRegistry",
        active_logins: Dict[str, "Session"],
        active_logins_lock: threading.Lock,
        ai_depth: int = 3,
    ) -> None:
        self.id: str = uuid.uuid4().hex
        self.socket: socket.socket = sock
        self.peer = peer
        self.account: Optional["Account"] = None
        self.current_game: Optional["Game"] = None
        self.state: SessionState = SessionState.ANONYMOUS

        self._reader = protocol.make_reader(sock)
        self._write_lock = threading.Lock()
        self._closed = False

        self.user_store = user_store
        self.lobby = lobby
        self.registry = registry
        self._active_logins = active_logins
        self._active_logins_lock = active_logins_lock
        self.ai_depth = ai_depth

    @property
    def username(self) -> str:
        return self.account.username if self.account else "<anonymous>"

    def send(self, msg: Dict[str, Any]) -> None:
        if self._closed:
            return
        with self._write_lock:
            try:
                protocol.send_message(self.socket, msg)
            except ConnectionClosed:
                self._closed = True

    def close(self) -> None:
        if self._closed:
            return
        self._closed = True
        try:
            self.socket.shutdown(socket.SHUT_RDWR)
        except OSError:
            pass
        try:
            self.socket.close()
        except OSError:
            pass

    def run(self) -> None:
        log.info("[%s] connected from %s:%d", self.id[:8], *self.peer)
        try:
            while not self._closed:
                try:
                    msg = protocol.recv_message(self._reader)
                except ConnectionClosed:
                    log.info("[%s] disconnected (%s)", self.id[:8], self.username)
                    break
                except ProtocolError as exc:
                    self.send(protocol.error(exc.code, str(exc)))
                    continue
                try:
                    self._dispatch(msg)
                except SecureChessError as exc:
                    self.send(protocol.error(exc.code, str(exc)))
                except Exception as exc:  # pragma: no cover - defensive
                    log.exception("[%s] unexpected error: %s", self.id[:8], exc)
                    self.send(protocol.error("internal_error", "internal server error"))
        finally:
            self._on_disconnect()
            self.close()

    def _dispatch(self, msg: Dict[str, Any]) -> None:
        kind = msg["type"]
        handler: Optional[Callable[[Dict[str, Any]], None]] = getattr(self, f"handle_{kind}", None)
        if handler is None:
            raise BadStateError(f"command {kind!r} is not legal in state {self.state.value!r}")
        handler(msg)

    def _on_disconnect(self) -> None:
        """Clean-up when the TCP socket goes away. Concrete behaviour added in US2."""
        if self.state is SessionState.IN_LOBBY:
            self.lobby.remove(self)
        elif self.state is SessionState.IN_GAME and self.current_game is not None:
            self.current_game.handle_disconnect(self)
            self.registry.end(self.current_game)
        if self.account is not None:
            with self._active_logins_lock:
                if self._active_logins.get(self.account.username) is self:
                    del self._active_logins[self.account.username]

    def handle_register(self, msg: Dict[str, Any]) -> None:
        if self.state is not SessionState.ANONYMOUS:
            raise BadStateError("already authenticated")
        username = msg.get("username", "")
        password = msg.get("password", "")
        if not isinstance(username, str) or not isinstance(password, str):
            raise ValidationError("username and password must be strings")
        try:
            account = self.user_store.register(username, password)
        except (DuplicateUserError, WeakPasswordError, ValidationError):
            raise
        self._login_as(account)
        self.send(protocol.ok())

    def handle_login(self, msg: Dict[str, Any]) -> None:
        if self.state is not SessionState.ANONYMOUS:
            raise BadStateError("already authenticated")
        username = msg.get("username", "")
        password = msg.get("password", "")
        if not isinstance(username, str) or not isinstance(password, str):
            raise ValidationError("username and password must be strings")
        account = self.user_store.authenticate(username, password)
        if account is None:
            raise AuthenticationError("invalid username or password")
        self._login_as(account)
        self.send(protocol.ok())

    def _login_as(self, account: "Account") -> None:
        with self._active_logins_lock:
            if account.username in self._active_logins:
                raise AlreadyLoggedInError(f"account {account.username!r} is already logged in")
            self._active_logins[account.username] = self
        self.account = account
        self.state = SessionState.AUTHENTICATED
        log.info("[%s] authenticated as %s", self.id[:8], account.username)

    def handle_quit(self, msg: Dict[str, Any]) -> None:
        self.send(protocol.ok())
        self._closed = True

    def handle_play_human(self, msg: Dict[str, Any]) -> None:
        if self.state is not SessionState.AUTHENTICATED:
            raise BadStateError("must be authenticated and not already in a game/lobby")
        self.state = SessionState.IN_LOBBY
        self.send(protocol.ok())
        pair = self.lobby.enqueue(self)
        if pair is not None:
            white, black = pair
            self.registry.create(white, black)

    def handle_cancel_lobby(self, msg: Dict[str, Any]) -> None:
        if self.state is not SessionState.IN_LOBBY:
            raise BadStateError("not currently waiting in a lobby")
        self.lobby.remove(self)
        self.state = SessionState.AUTHENTICATED
        self.send(protocol.ok())

    def handle_move(self, msg: Dict[str, Any]) -> None:
        if self.state is not SessionState.IN_GAME or self.current_game is None:
            raise BadStateError("no active game")
        uci = msg.get("uci", "")
        if not isinstance(uci, str):
            raise ValidationError("'uci' must be a string")
        from secure_chess.common.move import Move

        move = Move.parse(uci)
        game = self.current_game
        try:
            game.submit_move(self, move)
        except (IllegalMoveError, NotYourTurnError, MoveParseError):
            raise
        self.send(protocol.ok())
        if game.result is None:
            self.registry.broadcast_state(game, last_move=uci)
            next_session = game.current_session()
            if isinstance(next_session, AISession):
                self.registry.drive_ai(next_session)
        else:
            self.registry.end(game)

    def handle_resign(self, msg: Dict[str, Any]) -> None:
        if self.state is not SessionState.IN_GAME or self.current_game is None:
            raise BadStateError("no active game to resign")
        game = self.current_game
        game.resign(self)
        self.send(protocol.ok())
        self.registry.end(game)

    def handle_play_ai(self, msg: Dict[str, Any]) -> None:
        if self.state is not SessionState.AUTHENTICATED:
            raise BadStateError("must be authenticated and not already in a game/lobby")
        from secure_chess.common.pieces import Color

        requested_color = msg.get("color", "white")
        if requested_color not in ("white", "black"):
            raise ValidationError("'color' must be 'white' or 'black'")
        depth_raw = msg.get("depth", self.ai_depth)
        try:
            depth = int(depth_raw)
        except (TypeError, ValueError):
            raise ValidationError("'depth' must be an integer")
        depth = max(1, min(4, depth))

        ai_session = AISession(depth=depth, registry=self.registry)
        human_color = Color.WHITE if requested_color == "white" else Color.BLACK
        if human_color is Color.WHITE:
            white, black = self, ai_session
        else:
            white, black = ai_session, self
        self.registry.create(white, black)
        if isinstance(white, AISession):
            self.registry.drive_ai(white)


class AISession:
    """In-memory session impersonating an AI opponent.

    Implements the slice of the `Session` interface that `Game` and
    `GameRegistry` actually consume: `send`, `account`, `state`, `current_game`,
    and a `username`. Its `send` is a no-op because there is no socket.
    """

    def __init__(self, depth: int, registry: "GameRegistry") -> None:
        self.id: str = uuid.uuid4().hex
        self.account = _AIAccount(username=f"AI(depth={depth})")
        self.current_game: Optional["Game"] = None
        self.state: SessionState = SessionState.IN_GAME
        self.depth = depth
        self.registry = registry

    @property
    def username(self) -> str:
        return self.account.username

    def send(self, msg: Dict[str, Any]) -> None:  # pragma: no cover - intentionally inert
        return None

    def close(self) -> None:
        return None


class _AIAccount:
    """Tiny stand-in for an `Account` so `Game`/`GameRegistry` can treat the AI
    like any other player when reading `session.account.username`."""

    def __init__(self, username: str) -> None:
        self.username = username
