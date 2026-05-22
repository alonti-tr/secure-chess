from __future__ import annotations

from dataclasses import dataclass
from typing import Literal, Optional, Tuple

from secure_chess.common.errors import MoveParseError
from secure_chess.common.pieces import PieceType


Square = Tuple[int, int]


def to_algebraic(sq: Square) -> str:
    f, r = sq
    if not (0 <= f < 8 and 0 <= r < 8):
        raise ValueError(f"square out of range: {sq}")
    return f"{chr(ord('a') + f)}{r + 1}"


def from_algebraic(text: str) -> Square:
    if len(text) != 2 or text[0] not in "abcdefgh" or text[1] not in "12345678":
        raise MoveParseError(f"invalid square: {text!r}")
    return (ord(text[0]) - ord('a'), int(text[1]) - 1)


_PROMO_MAP: dict[str, PieceType] = {
    "q": PieceType.QUEEN,
    "r": PieceType.ROOK,
    "b": PieceType.BISHOP,
    "n": PieceType.KNIGHT,
}

_PROMO_REV: dict[PieceType, str] = {v: k for k, v in _PROMO_MAP.items()}


@dataclass(frozen=True)
class Move:
    origin: Square
    target: Square
    promotion: Optional[PieceType] = None
    is_capture: bool = False
    is_castle: Optional[Literal["K", "Q"]] = None
    is_en_passant: bool = False

    @classmethod
    def parse(cls, text: str) -> "Move":
        text = text.strip().lower()
        if not (4 <= len(text) <= 5):
            raise MoveParseError(f"move must be 4 or 5 characters, got {text!r}")
        origin = from_algebraic(text[0:2])
        target = from_algebraic(text[2:4])
        promotion: Optional[PieceType] = None
        if len(text) == 5:
            ch = text[4]
            if ch not in _PROMO_MAP:
                raise MoveParseError(f"invalid promotion piece: {ch!r}")
            promotion = _PROMO_MAP[ch]
        return cls(origin=origin, target=target, promotion=promotion)

    def to_uci(self) -> str:
        s = to_algebraic(self.origin) + to_algebraic(self.target)
        if self.promotion is not None:
            s += _PROMO_REV[self.promotion]
        return s

    def matches(self, other: "Move") -> bool:
        return (
            self.origin == other.origin
            and self.target == other.target
            and self.promotion == other.promotion
        )
