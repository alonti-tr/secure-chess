from __future__ import annotations

import math
import random
from typing import Optional

from secure_chess.common.board import Board
from secure_chess.common.move import Move
from secure_chess.common.pieces import Color, PieceType


PIECE_VALUE = {
    PieceType.PAWN: 100,
    PieceType.KNIGHT: 300,
    PieceType.BISHOP: 300,
    PieceType.ROOK: 500,
    PieceType.QUEEN: 900,
    PieceType.KING: 0,
}

CHECKMATE_SCORE = 1_000_000


def evaluate(board: Board) -> int:
    score = 0
    for f in range(8):
        for r in range(8):
            p = board.squares[f][r]
            if p is None:
                continue
            val = PIECE_VALUE[p.kind]
            score += val if p.color is Color.WHITE else -val
    return score


class AIPlayer:

    def choose_move(self, board: Board, depth: int = 3) -> Move:
        depth = max(1, min(4, int(depth)))
        side = board.side_to_move
        legal = board.legal_moves(side)
        if not legal:
            raise RuntimeError("AI asked to move in a terminal position")

        best_score: Optional[int] = None
        best_moves: list[Move] = []
        alpha, beta = -math.inf, math.inf

        for move in legal:
            snap = board._snapshot()
            try:
                board._apply_unchecked(move)
                score = -self._negamax(
                    board, depth - 1, -beta, -alpha,
                    perspective=side.opposite()
                )
            finally:
                board._restore(snap)

            if best_score is None or score > best_score:
                best_score = score
                best_moves = [move]
            elif score == best_score:
                best_moves.append(move)
            if score > alpha:
                alpha = score
        assert best_moves, "negamax produced no candidate"
        return random.choice(best_moves)

    def _negamax(self, board: Board, depth: int, alpha: float, beta: float, perspective: Color) -> int:
        if depth == 0:
            return self._perspective_eval(board, perspective)
        legal = board.legal_moves(board.side_to_move)
        if not legal:
            if board.is_in_check(board.side_to_move):
                return -CHECKMATE_SCORE + (10 - depth)
            return 0

        best = -math.inf
        for move in legal:
            snap = board._snapshot()
            try:
                board._apply_unchecked(move)
                score = -self._negamax(
                    board, depth - 1, -beta, -alpha,
                    perspective=perspective.opposite()
                )
            finally:
                board._restore(snap)
            if score > best:
                best = score
            if best > alpha:
                alpha = best
            if alpha >= beta:
                break
        return int(best)

    @staticmethod
    def _perspective_eval(board: Board, perspective: Color) -> int:
        white_score = evaluate(board)
        return white_score if perspective is Color.WHITE else -white_score
