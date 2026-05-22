"""T025 [US2] Chess rules end-to-end on the `Board` class."""

from __future__ import annotations

import pytest

from secure_chess.common.board import Board
from secure_chess.common.errors import IllegalMoveError
from secure_chess.common.move import Move
from secure_chess.common.pieces import Color, Piece, PieceType


def _play(board: Board, *ucis: str) -> Board:
    for uci in ucis:
        board.apply(Move.parse(uci))
    return board


def test_initial_position_to_fen():
    b = Board.initial()
    assert b.to_fen_short() == "rnbqkbnr/pppppppp/8/8/8/8/PPPPPPPP/RNBQKBNR w KQkq -"


def test_initial_legal_move_count_white_is_20():
    assert len(Board.initial().legal_moves(Color.WHITE)) == 20


def test_apply_e2e4_updates_position():
    b = Board.initial()
    b.apply(Move.parse("e2e4"))
    assert b.to_fen_short() == "rnbqkbnr/pppppppp/8/8/4P3/8/PPPP1PPP/RNBQKBNR b KQkq e3"


def test_illegal_move_raises():
    b = Board.initial()
    with pytest.raises(IllegalMoveError):
        b.apply(Move.parse("e2e5"))


def test_scholars_mate_is_checkmate():
    b = Board.initial()
    _play(b, "e2e4", "e7e5", "f1c4", "b8c6", "d1h5", "g8f6", "h5f7")
    assert b.is_checkmate()
    assert not b.is_stalemate()


def test_kingside_castling_white():
    b = Board.from_fen_short("r3k2r/pppppppp/8/8/8/8/PPPPPPPP/R3K2R w KQkq -")
    b.apply(Move.parse("e1g1"))
    assert b.piece_at((6, 0)) == Piece(PieceType.KING, Color.WHITE)
    assert b.piece_at((5, 0)) == Piece(PieceType.ROOK, Color.WHITE)
    assert "K" not in b.castling_rights
    assert "Q" not in b.castling_rights


def test_queenside_castling_white():
    b = Board.from_fen_short("r3k2r/pppppppp/8/8/8/8/PPPPPPPP/R3K2R w KQkq -")
    b.apply(Move.parse("e1c1"))
    assert b.piece_at((2, 0)) == Piece(PieceType.KING, Color.WHITE)
    assert b.piece_at((3, 0)) == Piece(PieceType.ROOK, Color.WHITE)


def test_castling_denied_when_through_check():
    b = Board.from_fen_short("4k3/8/8/8/8/8/5r2/R3K2R w KQ -")
    legal = [m.to_uci() for m in b.legal_moves(Color.WHITE)]
    assert "e1g1" not in legal


def test_castling_denied_after_king_moves():
    b = Board.from_fen_short("r3k2r/8/8/8/8/8/8/R3K2R w KQkq -")
    b.apply(Move.parse("e1e2"))
    b.apply(Move.parse("e8e7"))
    b.apply(Move.parse("e2e1"))
    b.apply(Move.parse("e7e8"))
    assert "K" not in b.castling_rights
    assert "Q" not in b.castling_rights
    assert "k" not in b.castling_rights
    assert "q" not in b.castling_rights


def test_en_passant_capture():
    b = Board.from_fen_short("rnbqkbnr/pp1ppppp/8/2p5/4P3/8/PPPP1PPP/RNBQKBNR w KQkq -")
    b.apply(Move.parse("e4e5"))
    b.apply(Move.parse("d7d5"))
    assert b.en_passant_target == (3, 5)
    b.apply(Move.parse("e5d6"))
    assert b.piece_at((3, 4)) is None
    assert b.piece_at((3, 5)) == Piece(PieceType.PAWN, Color.WHITE)


def test_pawn_promotion_to_queen():
    b = Board.from_fen_short("8/P7/8/8/8/8/8/4k2K w - -")
    b.apply(Move.parse("a7a8q"))
    assert b.piece_at((0, 7)) == Piece(PieceType.QUEEN, Color.WHITE)


def test_stalemate_position():
    b = Board.from_fen_short("7k/5Q2/6K1/8/8/8/8/8 b - -")
    assert not b.is_in_check(Color.BLACK)
    assert b.is_stalemate()
    assert not b.is_checkmate()


def test_check_detection():
    b = Board.from_fen_short("4k3/8/8/8/8/8/4R3/4K3 b - -")
    assert b.is_in_check(Color.BLACK)
    assert not b.is_checkmate()


def test_fifty_move_draw_flag():
    b = Board.initial()
    b.halfmove_clock = 100
    assert b.is_fifty_move_draw()


def test_threefold_repetition():
    b = Board.initial()
    for _ in range(2):
        b.apply(Move.parse("g1f3"))
        b.apply(Move.parse("g8f6"))
        b.apply(Move.parse("f3g1"))
        b.apply(Move.parse("f6g8"))
    assert b.is_threefold_repetition()
