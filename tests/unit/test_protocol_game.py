from __future__ import annotations

import socket

from secure_chess.common import protocol
from secure_chess.common.board import Board


def _round_trip(msg):
    a, b = socket.socketpair()
    try:
        protocol.send_message(a, msg)
        reader = protocol.make_reader(b)
        return protocol.recv_message(reader)
    finally:
        a.close()
        b.close()


def test_play_human_round_trip():
    assert _round_trip({"type": "play_human"}) == {"type": "play_human"}


def test_cancel_lobby_round_trip():
    assert _round_trip({"type": "cancel_lobby"}) == {"type": "cancel_lobby"}


def test_move_round_trip():
    assert _round_trip({"type": "move", "uci": "e2e4"}) == {"type": "move", "uci": "e2e4"}


def test_resign_round_trip():
    assert _round_trip({"type": "resign"}) == {"type": "resign"}


def test_game_started_round_trip():
    msg = {
        "type": "game_started",
        "game_id": "abc",
        "you_are": "white",
        "opponent": "bob",
        "board": "rnbqkbnr/pppppppp/8/8/8/8/PPPPPPPP/RNBQKBNR w KQkq -",
        "to_move": "white",
    }
    assert _round_trip(msg) == msg


def test_game_state_round_trip():
    msg = {
        "type": "game_state",
        "game_id": "abc",
        "last_move": "e2e4",
        "board": "rnbqkbnr/pppppppp/8/8/4P3/8/PPPP1PPP/RNBQKBNR b KQkq e3",
        "to_move": "black",
        "in_check": False,
    }
    assert _round_trip(msg) == msg


def test_game_ended_round_trip():
    msg = {"type": "game_ended", "game_id": "abc", "result": "white_wins", "reason": "checkmate"}
    assert _round_trip(msg) == msg


def test_starting_position_fen_matches_contract():
    b = Board.initial()
    assert b.to_fen_short() == "rnbqkbnr/pppppppp/8/8/8/8/PPPPPPPP/RNBQKBNR w KQkq -"
