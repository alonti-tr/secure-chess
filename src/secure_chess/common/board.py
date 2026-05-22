"""The `Board` class - the chess position plus the auxiliary state required by
the rules (castling rights, en-passant target square, halfmove clock, fullmove
number, repetition history).

`Board` is the heart of the game and the single source of truth for legality.
Every server-side move is fed through `Board.apply`, which is the only place
that mutates the position.
"""

from __future__ import annotations

from typing import Iterable, List, Optional, Tuple

from secure_chess.common.errors import IllegalMoveError
from secure_chess.common.move import Move, Square
from secure_chess.common.pieces import (
    BISHOP_RAYS,
    Color,
    KING_OFFSETS,
    KNIGHT_OFFSETS,
    Piece,
    PieceType,
    QUEEN_RAYS,
    ROOK_RAYS,
    SYMBOLS,
)


_PIECE_FROM_FEN = {
    "P": (PieceType.PAWN, Color.WHITE),
    "N": (PieceType.KNIGHT, Color.WHITE),
    "B": (PieceType.BISHOP, Color.WHITE),
    "R": (PieceType.ROOK, Color.WHITE),
    "Q": (PieceType.QUEEN, Color.WHITE),
    "K": (PieceType.KING, Color.WHITE),
    "p": (PieceType.PAWN, Color.BLACK),
    "n": (PieceType.KNIGHT, Color.BLACK),
    "b": (PieceType.BISHOP, Color.BLACK),
    "r": (PieceType.ROOK, Color.BLACK),
    "q": (PieceType.QUEEN, Color.BLACK),
    "k": (PieceType.KING, Color.BLACK),
}


class Board:
    """8x8 chess board. `squares[file][rank]` is the piece on that square."""

    __slots__ = (
        "squares",
        "side_to_move",
        "castling_rights",
        "en_passant_target",
        "halfmove_clock",
        "fullmove_number",
        "position_history",
    )

    def __init__(
        self,
        squares: Optional[List[List[Optional[Piece]]]] = None,
        side_to_move: Color = Color.WHITE,
        castling_rights: Optional[set[str]] = None,
        en_passant_target: Optional[Square] = None,
        halfmove_clock: int = 0,
        fullmove_number: int = 1,
    ) -> None:
        self.squares: List[List[Optional[Piece]]] = squares or [[None] * 8 for _ in range(8)]
        self.side_to_move: Color = side_to_move
        self.castling_rights: set[str] = castling_rights if castling_rights is not None else set()
        self.en_passant_target: Optional[Square] = en_passant_target
        self.halfmove_clock: int = halfmove_clock
        self.fullmove_number: int = fullmove_number
        self.position_history: List[str] = []

    @classmethod
    def initial(cls) -> "Board":
        squares: List[List[Optional[Piece]]] = [[None] * 8 for _ in range(8)]
        back = [
            PieceType.ROOK, PieceType.KNIGHT, PieceType.BISHOP, PieceType.QUEEN,
            PieceType.KING, PieceType.BISHOP, PieceType.KNIGHT, PieceType.ROOK,
        ]
        for f in range(8):
            squares[f][0] = Piece(back[f], Color.WHITE)
            squares[f][1] = Piece(PieceType.PAWN, Color.WHITE)
            squares[f][6] = Piece(PieceType.PAWN, Color.BLACK)
            squares[f][7] = Piece(back[f], Color.BLACK)
        b = cls(
            squares=squares,
            side_to_move=Color.WHITE,
            castling_rights={"K", "Q", "k", "q"},
            en_passant_target=None,
            halfmove_clock=0,
            fullmove_number=1,
        )
        b.position_history.append(b._fingerprint())
        return b

    def piece_at(self, sq: Square) -> Optional[Piece]:
        f, r = sq
        return self.squares[f][r]

    def _find_king(self, color: Color) -> Square:
        for f in range(8):
            for r in range(8):
                p = self.squares[f][r]
                if p is not None and p.kind is PieceType.KING and p.color is color:
                    return (f, r)
        raise AssertionError(f"no {color} king on the board")

    def is_square_attacked(self, sq: Square, by_color: Color) -> bool:
        """True if any `by_color` piece can capture onto `sq` (independent of
        whose turn it is). Used both for is_in_check and for castling tests."""
        f, r = sq
        for df, dr in KNIGHT_OFFSETS:
            t = (f + df, r + dr)
            if 0 <= t[0] < 8 and 0 <= t[1] < 8:
                p = self.squares[t[0]][t[1]]
                if p is not None and p.color is by_color and p.kind is PieceType.KNIGHT:
                    return True

        for df, dr in KING_OFFSETS:
            t = (f + df, r + dr)
            if 0 <= t[0] < 8 and 0 <= t[1] < 8:
                p = self.squares[t[0]][t[1]]
                if p is not None and p.color is by_color and p.kind is PieceType.KING:
                    return True

        pawn_dir = -1 if by_color is Color.WHITE else 1
        for df in (-1, 1):
            t = (f + df, r + pawn_dir)
            if 0 <= t[0] < 8 and 0 <= t[1] < 8:
                p = self.squares[t[0]][t[1]]
                if p is not None and p.color is by_color and p.kind is PieceType.PAWN:
                    return True

        for df, dr in BISHOP_RAYS:
            step = 1
            while True:
                t = (f + df * step, r + dr * step)
                if not (0 <= t[0] < 8 and 0 <= t[1] < 8):
                    break
                p = self.squares[t[0]][t[1]]
                if p is None:
                    step += 1
                    continue
                if p.color is by_color and p.kind in (PieceType.BISHOP, PieceType.QUEEN):
                    return True
                break

        for df, dr in ROOK_RAYS:
            step = 1
            while True:
                t = (f + df * step, r + dr * step)
                if not (0 <= t[0] < 8 and 0 <= t[1] < 8):
                    break
                p = self.squares[t[0]][t[1]]
                if p is None:
                    step += 1
                    continue
                if p.color is by_color and p.kind in (PieceType.ROOK, PieceType.QUEEN):
                    return True
                break

        return False

    def is_in_check(self, color: Color) -> bool:
        return self.is_square_attacked(self._find_king(color), color.opposite())

    def _pseudo_legal_moves(self, color: Color) -> Iterable[Move]:
        for f in range(8):
            for r in range(8):
                p = self.squares[f][r]
                if p is None or p.color is not color:
                    continue
                yield from p.pseudo_legal_targets(self, (f, r))

    def legal_moves(self, color: Optional[Color] = None) -> List[Move]:
        side = color or self.side_to_move
        result: List[Move] = []
        for move in list(self._pseudo_legal_moves(side)):
            snapshot = self._snapshot()
            try:
                self._apply_unchecked(move)
                in_check_after = self.is_in_check(side)
            finally:
                self._restore(snapshot)
            if not in_check_after:
                result.append(move)
        return result

    def is_legal(self, move: Move) -> bool:
        for legal in self.legal_moves(self.side_to_move):
            if legal.matches(move):
                return True
        return False

    def apply(self, move: Move) -> Move:
        """Validate `move` and mutate the board. Returns the resolved move with
        flags (capture, castle, en-passant, promotion) populated."""
        resolved: Optional[Move] = None
        for legal in self.legal_moves(self.side_to_move):
            if legal.matches(move):
                resolved = legal
                break
        if resolved is None:
            raise IllegalMoveError(f"move {move.to_uci()!r} is not legal in the current position")
        self._apply_unchecked(resolved)
        return resolved

    def _snapshot(self):
        return (
            [row[:] for row in self.squares],
            self.side_to_move,
            set(self.castling_rights),
            self.en_passant_target,
            self.halfmove_clock,
            self.fullmove_number,
            list(self.position_history),
        )

    def _restore(self, snap) -> None:
        (
            self.squares,
            self.side_to_move,
            self.castling_rights,
            self.en_passant_target,
            self.halfmove_clock,
            self.fullmove_number,
            self.position_history,
        ) = (
            [row[:] for row in snap[0]],
            snap[1],
            set(snap[2]),
            snap[3],
            snap[4],
            snap[5],
            list(snap[6]),
        )

    def _apply_unchecked(self, move: Move) -> None:
        origin = move.origin
        target = move.target
        piece = self.squares[origin[0]][origin[1]]
        assert piece is not None, f"no piece on origin square {origin}"

        captured = self.squares[target[0]][target[1]]
        is_pawn_move = piece.kind is PieceType.PAWN
        is_capture = captured is not None or move.is_en_passant

        self.squares[origin[0]][origin[1]] = None

        if move.is_en_passant:
            cap_rank = target[1] + (-1 if piece.color is Color.WHITE else 1)
            self.squares[target[0]][cap_rank] = None
            self.squares[target[0]][target[1]] = piece
        elif move.promotion is not None:
            self.squares[target[0]][target[1]] = Piece(move.promotion, piece.color)
        else:
            self.squares[target[0]][target[1]] = piece

        if move.is_castle == "K":
            rank = 0 if piece.color is Color.WHITE else 7
            rook = self.squares[7][rank]
            self.squares[7][rank] = None
            self.squares[5][rank] = rook
        elif move.is_castle == "Q":
            rank = 0 if piece.color is Color.WHITE else 7
            rook = self.squares[0][rank]
            self.squares[0][rank] = None
            self.squares[3][rank] = rook

        if piece.kind is PieceType.KING:
            if piece.color is Color.WHITE:
                self.castling_rights.discard("K")
                self.castling_rights.discard("Q")
            else:
                self.castling_rights.discard("k")
                self.castling_rights.discard("q")
        if piece.kind is PieceType.ROOK:
            if origin == (0, 0):
                self.castling_rights.discard("Q")
            elif origin == (7, 0):
                self.castling_rights.discard("K")
            elif origin == (0, 7):
                self.castling_rights.discard("q")
            elif origin == (7, 7):
                self.castling_rights.discard("k")
        if target == (0, 0):
            self.castling_rights.discard("Q")
        elif target == (7, 0):
            self.castling_rights.discard("K")
        elif target == (0, 7):
            self.castling_rights.discard("q")
        elif target == (7, 7):
            self.castling_rights.discard("k")

        if is_pawn_move and abs(target[1] - origin[1]) == 2:
            self.en_passant_target = (origin[0], (origin[1] + target[1]) // 2)
        else:
            self.en_passant_target = None

        if is_pawn_move or is_capture:
            self.halfmove_clock = 0
        else:
            self.halfmove_clock += 1

        if self.side_to_move is Color.BLACK:
            self.fullmove_number += 1
        self.side_to_move = self.side_to_move.opposite()

        self.position_history.append(self._fingerprint())

    def is_checkmate(self) -> bool:
        return self.is_in_check(self.side_to_move) and not self.legal_moves(self.side_to_move)

    def is_stalemate(self) -> bool:
        return not self.is_in_check(self.side_to_move) and not self.legal_moves(self.side_to_move)

    def is_fifty_move_draw(self) -> bool:
        return self.halfmove_clock >= 100

    def is_threefold_repetition(self) -> bool:
        if not self.position_history:
            return False
        last = self.position_history[-1]
        return self.position_history.count(last) >= 3

    def _fingerprint(self) -> str:
        return self._placement_only() + " " + ("w" if self.side_to_move is Color.WHITE else "b") \
            + " " + self._castling_string() + " " + self._ep_string()

    def _placement_only(self) -> str:
        rows: List[str] = []
        for rank in range(7, -1, -1):
            row = ""
            empty = 0
            for f in range(8):
                p = self.squares[f][rank]
                if p is None:
                    empty += 1
                else:
                    if empty:
                        row += str(empty)
                        empty = 0
                    sym = SYMBOLS[p.kind]
                    row += sym if p.color is Color.WHITE else sym.lower()
            if empty:
                row += str(empty)
            rows.append(row)
        return "/".join(rows)

    def _castling_string(self) -> str:
        if not self.castling_rights:
            return "-"
        order = "KQkq"
        return "".join(ch for ch in order if ch in self.castling_rights)

    def _ep_string(self) -> str:
        from secure_chess.common.move import to_algebraic
        return to_algebraic(self.en_passant_target) if self.en_passant_target else "-"

    def to_fen_short(self) -> str:
        return f"{self._placement_only()} {'w' if self.side_to_move is Color.WHITE else 'b'} {self._castling_string()} {self._ep_string()}"

    @classmethod
    def from_fen_short(cls, fen: str) -> "Board":
        from secure_chess.common.move import from_algebraic

        parts = fen.strip().split()
        if len(parts) < 4:
            raise ValueError(f"invalid short FEN: {fen!r}")
        placement, side, castling, ep = parts[0], parts[1], parts[2], parts[3]
        squares: List[List[Optional[Piece]]] = [[None] * 8 for _ in range(8)]
        rows = placement.split("/")
        if len(rows) != 8:
            raise ValueError(f"invalid placement: {placement!r}")
        for rank_idx, row in enumerate(rows):
            rank = 7 - rank_idx
            f = 0
            for ch in row:
                if ch.isdigit():
                    f += int(ch)
                else:
                    kind, color = _PIECE_FROM_FEN[ch]
                    squares[f][rank] = Piece(kind, color)
                    f += 1
        b = cls(
            squares=squares,
            side_to_move=Color.WHITE if side == "w" else Color.BLACK,
            castling_rights=set() if castling == "-" else set(castling),
            en_passant_target=None if ep == "-" else from_algebraic(ep),
        )
        b.position_history.append(b._fingerprint())
        return b
