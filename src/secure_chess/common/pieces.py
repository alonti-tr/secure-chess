"""Chess piece types, colours, and per-piece pseudo-legal move generation.

This module defines:
  - `Color`             : WHITE / BLACK with an `opposite()` helper.
  - `PieceType`         : the six chess piece kinds.
  - `SYMBOLS`           : single-letter symbols for ASCII rendering (uppercase = white).
  - `Piece`             : a (kind, color) pair plus pseudo-legal move generation.

Pseudo-legal generation does NOT filter out moves that leave the mover's own king in check
- that filter lives on `Board` (see `board.py`) so the same per-piece logic is reused both
for legal move generation and for the king-attack scan used by `Board.is_in_check`.
"""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from typing import TYPE_CHECKING, Iterator

if TYPE_CHECKING:
    from secure_chess.common.board import Board
    from secure_chess.common.move import Move, Square


class Color(Enum):
    WHITE = "white"
    BLACK = "black"

    def opposite(self) -> "Color":
        return Color.BLACK if self is Color.WHITE else Color.WHITE


class PieceType(Enum):
    PAWN = "pawn"
    KNIGHT = "knight"
    BISHOP = "bishop"
    ROOK = "rook"
    QUEEN = "queen"
    KING = "king"


SYMBOLS: dict[PieceType, str] = {
    PieceType.PAWN: "P",
    PieceType.KNIGHT: "N",
    PieceType.BISHOP: "B",
    PieceType.ROOK: "R",
    PieceType.QUEEN: "Q",
    PieceType.KING: "K",
}


KNIGHT_OFFSETS: tuple[tuple[int, int], ...] = (
    (1, 2), (2, 1), (2, -1), (1, -2),
    (-1, -2), (-2, -1), (-2, 1), (-1, 2),
)

KING_OFFSETS: tuple[tuple[int, int], ...] = (
    (1, 0), (1, 1), (0, 1), (-1, 1),
    (-1, 0), (-1, -1), (0, -1), (1, -1),
)

BISHOP_RAYS: tuple[tuple[int, int], ...] = ((1, 1), (1, -1), (-1, 1), (-1, -1))
ROOK_RAYS:   tuple[tuple[int, int], ...] = ((1, 0), (-1, 0), (0, 1), (0, -1))
QUEEN_RAYS:  tuple[tuple[int, int], ...] = BISHOP_RAYS + ROOK_RAYS


def _in_bounds(sq: "Square") -> bool:
    f, r = sq
    return 0 <= f < 8 and 0 <= r < 8


@dataclass(frozen=True)
class Piece:
    kind: PieceType
    color: Color

    def symbol(self) -> str:
        s = SYMBOLS[self.kind]
        return s if self.color is Color.WHITE else s.lower()

    def pseudo_legal_targets(self, board: "Board", origin: "Square") -> Iterator["Move"]:
        from secure_chess.common.move import Move

        if self.kind is PieceType.KNIGHT:
            yield from _jump_moves(self, board, origin, KNIGHT_OFFSETS)
        elif self.kind is PieceType.KING:
            yield from _jump_moves(self, board, origin, KING_OFFSETS)
            yield from _castling_moves(self, board, origin)
        elif self.kind is PieceType.BISHOP:
            yield from _ray_moves(self, board, origin, BISHOP_RAYS)
        elif self.kind is PieceType.ROOK:
            yield from _ray_moves(self, board, origin, ROOK_RAYS)
        elif self.kind is PieceType.QUEEN:
            yield from _ray_moves(self, board, origin, QUEEN_RAYS)
        elif self.kind is PieceType.PAWN:
            yield from _pawn_moves(self, board, origin)
        else:
            raise AssertionError(f"unknown piece kind: {self.kind}")
        del Move


def _jump_moves(piece: Piece, board: "Board", origin: "Square", offsets) -> Iterator["Move"]:
    from secure_chess.common.move import Move

    f, r = origin
    for df, dr in offsets:
        target = (f + df, r + dr)
        if not _in_bounds(target):
            continue
        occupant = board.piece_at(target)
        if occupant is None:
            yield Move(origin=origin, target=target)
        elif occupant.color is not piece.color:
            yield Move(origin=origin, target=target, is_capture=True)


def _ray_moves(piece: Piece, board: "Board", origin: "Square", rays) -> Iterator["Move"]:
    from secure_chess.common.move import Move

    f, r = origin
    for df, dr in rays:
        step = 1
        while True:
            target = (f + df * step, r + dr * step)
            if not _in_bounds(target):
                break
            occupant = board.piece_at(target)
            if occupant is None:
                yield Move(origin=origin, target=target)
            else:
                if occupant.color is not piece.color:
                    yield Move(origin=origin, target=target, is_capture=True)
                break
            step += 1


def _pawn_moves(piece: Piece, board: "Board", origin: "Square") -> Iterator["Move"]:
    from secure_chess.common.move import Move

    f, r = origin
    direction = 1 if piece.color is Color.WHITE else -1
    start_rank = 1 if piece.color is Color.WHITE else 6
    promotion_rank = 7 if piece.color is Color.WHITE else 0

    one_step = (f, r + direction)
    if _in_bounds(one_step) and board.piece_at(one_step) is None:
        if one_step[1] == promotion_rank:
            for promo in (PieceType.QUEEN, PieceType.ROOK, PieceType.BISHOP, PieceType.KNIGHT):
                yield Move(origin=origin, target=one_step, promotion=promo)
        else:
            yield Move(origin=origin, target=one_step)

            two_step = (f, r + 2 * direction)
            if r == start_rank and board.piece_at(two_step) is None:
                yield Move(origin=origin, target=two_step)

    for df in (-1, 1):
        diag = (f + df, r + direction)
        if not _in_bounds(diag):
            continue
        occupant = board.piece_at(diag)
        if occupant is not None and occupant.color is not piece.color:
            if diag[1] == promotion_rank:
                for promo in (PieceType.QUEEN, PieceType.ROOK, PieceType.BISHOP, PieceType.KNIGHT):
                    yield Move(origin=origin, target=diag, promotion=promo, is_capture=True)
            else:
                yield Move(origin=origin, target=diag, is_capture=True)
        elif occupant is None and board.en_passant_target == diag:
            yield Move(origin=origin, target=diag, is_capture=True, is_en_passant=True)


def _castling_moves(piece: Piece, board: "Board", origin: "Square") -> Iterator["Move"]:
    from secure_chess.common.move import Move

    if piece.color is Color.WHITE:
        if origin != (4, 0):
            return
        king_side_flag, queen_side_flag = "K", "Q"
        rank = 0
    else:
        if origin != (4, 7):
            return
        king_side_flag, queen_side_flag = "k", "q"
        rank = 7

    if king_side_flag in board.castling_rights:
        if board.piece_at((5, rank)) is None and board.piece_at((6, rank)) is None:
            if not board.is_square_attacked((4, rank), piece.color.opposite()) \
               and not board.is_square_attacked((5, rank), piece.color.opposite()) \
               and not board.is_square_attacked((6, rank), piece.color.opposite()):
                yield Move(origin=origin, target=(6, rank), is_castle="K")
    if queen_side_flag in board.castling_rights:
        if (board.piece_at((1, rank)) is None
                and board.piece_at((2, rank)) is None
                and board.piece_at((3, rank)) is None):
            if not board.is_square_attacked((4, rank), piece.color.opposite()) \
               and not board.is_square_attacked((3, rank), piece.color.opposite()) \
               and not board.is_square_attacked((2, rank), piece.color.opposite()):
                yield Move(origin=origin, target=(2, rank), is_castle="Q")
