from __future__ import annotations

import uuid
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import TYPE_CHECKING, Optional

from secure_chess.common.board import Board
from secure_chess.common.errors import IllegalMoveError, NotYourTurnError
from secure_chess.common.move import Move
from secure_chess.common.pieces import Color


if TYPE_CHECKING:
    from secure_chess.server.session import Session


@dataclass(frozen=True)
class Result:
    outcome: str
    reason: str

    @classmethod
    def white_wins(cls, reason: str) -> "Result":
        return cls("white_wins", reason)

    @classmethod
    def black_wins(cls, reason: str) -> "Result":
        return cls("black_wins", reason)

    @classmethod
    def draw(cls, reason: str) -> "Result":
        return cls("draw", reason)


class Game:
    def __init__(self, white: "Session", black: "Session") -> None:
        self.id: str = uuid.uuid4().hex
        self.white: "Session" = white
        self.black: "Session" = black
        self.board: Board = Board.initial()
        self.result: Optional[Result] = None
        self.created_at: datetime = datetime.now(timezone.utc)
        self.last_move_uci: Optional[str] = None

    def current_session(self) -> "Session":
        return self.white if self.board.side_to_move is Color.WHITE else self.black

    def color_of(self, session: "Session") -> Optional[Color]:
        if session is self.white:
            return Color.WHITE
        if session is self.black:
            return Color.BLACK
        return None

    def submit_move(self, session: "Session", move: Move) -> Move:
        if self.result is not None:
            raise IllegalMoveError("game has already ended")
        mover_color = self.color_of(session)
        if mover_color is None:
            raise NotYourTurnError("session is not a participant in this game")
        if mover_color is not self.board.side_to_move:
            raise NotYourTurnError("it is not your turn")

        resolved = self.board.apply(move)
        self.last_move_uci = resolved.to_uci()

        if self.board.is_checkmate():
            winner_color = mover_color
            self.result = (
                Result.white_wins("checkmate") if winner_color is Color.WHITE
                else Result.black_wins("checkmate")
            )
        elif self.board.is_stalemate():
            self.result = Result.draw("stalemate")
        elif self.board.is_threefold_repetition():
            self.result = Result.draw("threefold")
        elif self.board.is_fifty_move_draw():
            self.result = Result.draw("fifty-move")
        return resolved

    def resign(self, session: "Session") -> None:
        if self.result is not None:
            return
        color = self.color_of(session)
        if color is Color.WHITE:
            self.result = Result.black_wins("resignation")
        else:
            self.result = Result.white_wins("resignation")

    def handle_disconnect(self, session: "Session") -> None:
        if self.result is not None:
            return
        color = self.color_of(session)
        if color is Color.WHITE:
            self.result = Result.black_wins("disconnect")
        elif color is Color.BLACK:
            self.result = Result.white_wins("disconnect")
