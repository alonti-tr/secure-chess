from __future__ import annotations

import pytest

from secure_chess.client.gui import (
    PIECE_TO_UNICODE,
    grid_to_square,
    parse_fen_placement,
    square_to_grid,
)


class TestParseFenPlacement:
    def test_initial_position(self) -> None:
        fen = "rnbqkbnr/pppppppp/8/8/8/8/PPPPPPPP/RNBQKBNR w KQkq - 0 1"
        grid = parse_fen_placement(fen)

        assert grid[0][0] == "r"
        assert grid[7][0] == "r"
        assert grid[4][0] == "k"
        assert grid[3][0] == "q"
        assert grid[4][7] == "K"
        assert grid[3][7] == "Q"

        for f in range(8):
            assert grid[f][1] == "p"
            assert grid[f][6] == "P"

        for f in range(8):
            for rt in (2, 3, 4, 5):
                assert grid[f][rt] is None

    def test_empty_squares_run(self) -> None:
        fen = "8/8/8/4P3/8/8/8/8 w - - 0 1"
        grid = parse_fen_placement(fen)
        assert grid[4][3] == "P"
        non_empty = [(f, rt) for f in range(8) for rt in range(8) if grid[f][rt] is not None]
        assert non_empty == [(4, 3)]

    def test_short_fen_without_extra_fields(self) -> None:
        grid = parse_fen_placement("8/8/8/8/8/8/8/8")
        for f in range(8):
            for rt in range(8):
                assert grid[f][rt] is None

    def test_empty_string_yields_empty_grid(self) -> None:
        grid = parse_fen_placement("")
        assert all(cell is None for row in grid for cell in row)

    def test_malformed_row_does_not_crash(self) -> None:
        grid = parse_fen_placement("rnbqkbnr/pppppppp/8/8/8/8/8/8 w - - 0 1")
        assert grid[0][0] == "r"
        assert grid[0][1] == "p"


class TestSquareToGrid:
    def test_corners(self) -> None:
        assert square_to_grid("a1") == (0, 7)
        assert square_to_grid("h1") == (7, 7)
        assert square_to_grid("a8") == (0, 0)
        assert square_to_grid("h8") == (7, 0)

    def test_pawn_starting_squares(self) -> None:
        assert square_to_grid("e2") == (4, 6)
        assert square_to_grid("e7") == (4, 1)

    @pytest.mark.parametrize(
        "bad",
        ["", "a", "abc", "z1", "a0", "a9", "i1", "11", "aa", None, 42],
    )
    def test_invalid_returns_none(self, bad: object) -> None:
        assert square_to_grid(bad) is None


class TestGridToSquare:
    def test_corners(self) -> None:
        assert grid_to_square(0, 7) == "a1"
        assert grid_to_square(7, 7) == "h1"
        assert grid_to_square(0, 0) == "a8"
        assert grid_to_square(7, 0) == "h8"

    def test_round_trip(self) -> None:
        for sq in ["e4", "d5", "g7", "b2", "h1", "a8"]:
            fr = square_to_grid(sq)
            assert fr is not None
            assert grid_to_square(*fr) == sq


class TestPieceGlyphs:
    def test_every_fen_letter_has_a_glyph(self) -> None:
        for letter in "KQRBNPkqrbnp":
            assert letter in PIECE_TO_UNICODE
            assert len(PIECE_TO_UNICODE[letter]) == 1

    def test_white_and_black_glyphs_differ(self) -> None:
        for white, black in [("K", "k"), ("Q", "q"), ("R", "r"), ("B", "b"), ("N", "n"), ("P", "p")]:
            assert PIECE_TO_UNICODE[white] != PIECE_TO_UNICODE[black]
