from __future__ import annotations

import threading
from typing import TYPE_CHECKING, Dict, List, Optional, Tuple

from secure_chess.common.game import Game, Result
from secure_chess.common.log import get_logger
from secure_chess.common.pieces import Color


if TYPE_CHECKING:
    from secure_chess.server.session import Session


log = get_logger("lobby")


class Lobby:
    def __init__(self) -> None:
        self._queue: List["Session"] = []
        self._lock = threading.Lock()

    def enqueue(self, session: "Session") -> Optional[Tuple["Session", "Session"]]:
        with self._lock:
            self._queue.append(session)
            if len(self._queue) >= 2:
                white = self._queue.pop(0)
                black = self._queue.pop(0)
                return white, black
        return None

    def remove(self, session: "Session") -> None:
        with self._lock:
            try:
                self._queue.remove(session)
            except ValueError:
                pass

    def __len__(self) -> int:
        with self._lock:
            return len(self._queue)


class GameRegistry:
    def __init__(self) -> None:
        self._games: Dict[str, Game] = {}
        self._lock = threading.Lock()

    def create(self, white: "Session", black: "Session") -> Game:
        from secure_chess.server.session import SessionState

        game = Game(white=white, black=black)
        with self._lock:
            self._games[game.id] = game

        white.current_game = game
        black.current_game = game
        white.state = SessionState.IN_GAME
        black.state = SessionState.IN_GAME

        board_fen = game.board.to_fen_short()
        white.send({
            "type": "game_started",
            "game_id": game.id,
            "you_are": "white",
            "opponent": black.username,
            "board": board_fen,
            "to_move": "white",
        })
        black.send({
            "type": "game_started",
            "game_id": game.id,
            "you_are": "black",
            "opponent": white.username,
            "board": board_fen,
            "to_move": "white",
        })
        log.info("game %s started: %s (white) vs %s (black)",
                 game.id[:8], white.username, black.username)
        return game

    def broadcast_state(self, game: Game, last_move: str) -> None:
        msg = {
            "type": "game_state",
            "game_id": game.id,
            "last_move": last_move,
            "board": game.board.to_fen_short(),
            "to_move": "white" if game.board.side_to_move is Color.WHITE else "black",
            "in_check": game.board.is_in_check(game.board.side_to_move),
        }
        game.white.send(msg)
        game.black.send(msg)

    def end(self, game: Game) -> None:
        from secure_chess.server.session import SessionState

        with self._lock:
            self._games.pop(game.id, None)

        if game.result is None:
            game.result = Result.draw("aborted")

        msg = {
            "type": "game_ended",
            "game_id": game.id,
            "result": game.result.outcome,
            "reason": game.result.reason,
        }
        game.white.send(msg)
        game.black.send(msg)
        game.white.current_game = None
        game.black.current_game = None
        if game.white.state is SessionState.IN_GAME:
            game.white.state = SessionState.AUTHENTICATED
        if game.black.state is SessionState.IN_GAME:
            game.black.state = SessionState.AUTHENTICATED
        log.info("game %s ended: %s (%s)",
                 game.id[:8], game.result.outcome, game.result.reason)

    def drive_ai(self, ai_session) -> None:
        from secure_chess.common.ai import AIPlayer

        game = ai_session.current_game
        if game is None or game.result is not None:
            return
        if game.current_session() is not ai_session:
            return
        move = AIPlayer().choose_move(game.board, depth=ai_session.depth)
        resolved = game.submit_move(ai_session, move)
        uci = resolved.to_uci()
        if game.result is None:
            self.broadcast_state(game, last_move=uci)
        else:
            self.end(game)

    def active_games(self) -> List[Game]:
        with self._lock:
            return list(self._games.values())
