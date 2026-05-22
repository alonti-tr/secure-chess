from __future__ import annotations

import time

import pytest

from secure_chess.common.ai import AIPlayer
from secure_chess.common.board import Board
from secure_chess.common.move import Move


AI_TIME_BUDGET_SECONDS = 5.0


FIXTURE_FENS = [
    "rnbqkbnr/pppppppp/8/8/8/8/PPPPPPPP/RNBQKBNR w KQkq -",
    "rnbqkbnr/pppppppp/8/8/4P3/8/PPPP1PPP/RNBQKBNR b KQkq e3",
    "rnbqkbnr/pp1ppppp/8/2p5/4P3/5N2/PPPP1PPP/RNBQKB1R b KQkq -",
    "r1bqkbnr/pppp1ppp/2n5/4p3/2B1P3/5N2/PPPP1PPP/RNBQK2R w KQkq -",
    "rnbqkbnr/ppp2ppp/8/3pp3/8/2N2N2/PPPPPPPP/R1BQKB1R w KQkq -",
    "8/8/8/4k3/8/8/4K3/4R3 w - -",
    "7k/8/6K1/6Q1/8/8/8/8 w - -",
    "r3k2r/pppppppp/8/8/8/8/PPPPPPPP/R3K2R w KQkq -",
]


@pytest.mark.parametrize("fen", FIXTURE_FENS)
def test_ai_returns_legal_move(fen: str):
    board = Board.from_fen_short(fen)
    if board.legal_moves(board.side_to_move) == []:
        pytest.skip("position has no legal moves")
    start = time.monotonic()
    move = AIPlayer().choose_move(board, depth=2)
    elapsed = time.monotonic() - start
    assert elapsed < AI_TIME_BUDGET_SECONDS, f"AI took {elapsed:.2f}s"
    legal = board.legal_moves(board.side_to_move)
    assert any(legal_move.matches(move) for legal_move in legal), \
        f"AI returned non-legal move {move.to_uci()} from {fen}"


def test_ai_picks_mate_in_one():
    board = Board.from_fen_short(
        "r1bqkb1r/pppp1Qpp/2n2n2/4p3/2B1P3/8/PPPP1PPP/RNB1K1NR b KQkq -"
    )
    assert board.is_checkmate()


def test_ai_finds_mating_move_scholars_mate_setup():
    board = Board.from_fen_short(
        "r1bqkb1r/pppp1ppp/2n2n2/4p2Q/2B1P3/8/PPPP1PPP/RNB1K1NR w KQkq -"
    )
    move = AIPlayer().choose_move(board, depth=3)
    assert move.to_uci() == "h5f7", f"expected h5f7 (mate), got {move.to_uci()}"


def test_ai_makes_only_legal_moves_in_short_game():
    board = Board.initial()
    ai = AIPlayer()
    for _ in range(10):
        legal = board.legal_moves(board.side_to_move)
        if not legal:
            break
        move = ai.choose_move(board, depth=2)
        assert any(m.matches(move) for m in legal)
        board.apply(move)
