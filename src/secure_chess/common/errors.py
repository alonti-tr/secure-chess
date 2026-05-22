from __future__ import annotations


class SecureChessError(Exception):

    code: str = "internal_error"


class MoveParseError(SecureChessError):
    code = "move_parse_error"


class IllegalMoveError(SecureChessError):
    code = "illegal_move"


class NotYourTurnError(SecureChessError):
    code = "not_your_turn"


class ProtocolError(SecureChessError):
    code = "bad_request"


class ConnectionClosed(SecureChessError):
    code = "connection_closed"


class DuplicateUserError(SecureChessError):
    code = "duplicate_user"


class WeakPasswordError(SecureChessError):
    code = "weak_password"


class AuthenticationError(SecureChessError):
    code = "auth_failed"


class AlreadyLoggedInError(SecureChessError):
    code = "already_logged_in"


class ValidationError(SecureChessError):
    code = "validation"


class BadStateError(SecureChessError):
    code = "bad_state"


class ServerFullError(SecureChessError):
    code = "server_full"
