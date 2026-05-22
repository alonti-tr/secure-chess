from __future__ import annotations

from secure_chess.common.board import Board
from secure_chess.common.move import from_algebraic
from secure_chess.common.pieces import Color, Piece, PieceType


def _empty_board() -> Board:
    return Board(side_to_move=Color.WHITE, castling_rights=set())


def _targets(board: Board, sq) -> set:
    piece = board.piece_at(sq)
    assert piece is not None
    return {m.target for m in piece.pseudo_legal_targets(board, sq)}


def test_knight_l_jumps_from_centre():
    b = _empty_board()
    b.squares[3][3] = Piece(PieceType.KNIGHT, Color.WHITE)
    expected = {(1, 2), (2, 1), (4, 1), (5, 2), (5, 4), (4, 5), (2, 5), (1, 4)}
    assert _targets(b, (3, 3)) == expected


def test_knight_jumps_capture_or_skip_own_pieces():
    b = _empty_board()
    b.squares[3][3] = Piece(PieceType.KNIGHT, Color.WHITE)
    b.squares[1][2] = Piece(PieceType.PAWN, Color.WHITE)
    b.squares[5][2] = Piece(PieceType.PAWN, Color.BLACK)
    targets = _targets(b, (3, 3))
    assert (1, 2) not in targets
    assert (5, 2) in targets


def test_bishop_diagonal_blocked_by_own_piece():
    b = _empty_board()
    b.squares[3][3] = Piece(PieceType.BISHOP, Color.WHITE)
    b.squares[5][5] = Piece(PieceType.PAWN, Color.WHITE)
    targets = _targets(b, (3, 3))
    assert (4, 4) in targets
    assert (5, 5) not in targets
    assert (6, 6) not in targets


def test_bishop_can_capture_enemy_piece_blocking():
    b = _empty_board()
    b.squares[3][3] = Piece(PieceType.BISHOP, Color.WHITE)
    b.squares[5][5] = Piece(PieceType.PAWN, Color.BLACK)
    targets = _targets(b, (3, 3))
    assert (5, 5) in targets
    assert (6, 6) not in targets


def test_pawn_single_and_double_push_from_start():
    b = _empty_board()
    b.squares[4][1] = Piece(PieceType.PAWN, Color.WHITE)
    targets = _targets(b, (4, 1))
    assert (4, 2) in targets
    assert (4, 3) in targets


def test_pawn_no_double_push_when_blocked():
    b = _empty_board()
    b.squares[4][1] = Piece(PieceType.PAWN, Color.WHITE)
    b.squares[4][2] = Piece(PieceType.PAWN, Color.BLACK)
    targets = _targets(b, (4, 1))
    assert (4, 2) not in targets
    assert (4, 3) not in targets


def test_pawn_diagonal_capture():
    b = _empty_board()
    b.squares[4][1] = Piece(PieceType.PAWN, Color.WHITE)
    b.squares[3][2] = Piece(PieceType.PAWN, Color.BLACK)
    b.squares[5][2] = Piece(PieceType.KNIGHT, Color.BLACK)
    targets = _targets(b, (4, 1))
    assert (3, 2) in targets
    assert (5, 2) in targets


def test_pawn_promotion_expansion():
    b = _empty_board()
    b.squares[0][6] = Piece(PieceType.PAWN, Color.WHITE)
    moves = list(b.piece_at((0, 6)).pseudo_legal_targets(b, (0, 6)))
    promos = [m.promotion for m in moves if m.target == (0, 7)]
    assert set(promos) == {PieceType.QUEEN, PieceType.ROOK, PieceType.BISHOP, PieceType.KNIGHT}


def test_rook_rays():
    b = _empty_board()
    b.squares[3][3] = Piece(PieceType.ROOK, Color.WHITE)
    targets = _targets(b, (3, 3))
    assert {(0, 3), (1, 3), (2, 3), (4, 3), (5, 3), (6, 3), (7, 3)} <= targets
    assert {(3, 0), (3, 1), (3, 2), (3, 4), (3, 5), (3, 6), (3, 7)} <= targets
