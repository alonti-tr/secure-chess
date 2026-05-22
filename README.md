# Secure Two-Player Chess

A networked chess game in which two participants authenticate to a central
server with password-protected accounts and play a full game of chess against
each other through a plain-text JSON-Lines protocol over TCP. Passwords are
hashed at rest with `bcrypt`. The server can host multiple parallel games and
includes an alpha-beta minimax AI opponent. Two client front-ends are bundled:
a Tkinter desktop GUI (default) and a text-mode REPL (`--cli`); both speak the
same wire protocol so a GUI client can play against a CLI client.

## Assignment requirements coverage

| Requirement | Where it is satisfied |
|---|---|
| ≥ 2 different classes with objects + operations | `src/secure_chess/common/`: `Piece`, `Board`, `Move`, `Game`, `Account`, `UserStore`, plus server-side `Session`, `Lobby`, `GameRegistry` — 9 domain classes total, each with state and non-trivial operations. |
| Client–server system | `src/secure_chess/server/` + `src/secure_chess/client/`, TCP transport, JSON-Lines wire protocol (see `specs/002-secure-chess/contracts/wire-protocol.md`). |
| Multiple parallel clients (**bonus**) | One server thread per accepted connection + thread-safe `Lobby` and `GameRegistry`. Verified by `tests/integration/test_parallel_games.py`. |
| AI usage (**bonus**) | `src/secure_chess/common/ai.py` — alpha-beta minimax with material evaluation. `play_ai` wire command. Verified by `tests/unit/test_ai.py` and `tests/integration/test_play_ai.py`. |
| OS — no change | Pure Python 3.11+; runs in the default Windows or Linux shell with `python -m secure_chess.server` / `python -m secure_chess.client`. |
| Encrypt **passwords only** (no comms encryption) | `bcrypt` in `src/secure_chess/common/crypto.py`; on-disk store `data/users.json` contains only hashes. The TCP transport is deliberately plain text per the assignment. |
| UI | Tkinter desktop GUI (`secure_chess.client.gui`) with login/lobby/game screens and a clickable 8×8 board, plus the original text-mode REPL (`--cli`). Tkinter is part of the Python standard library, so no extra UI framework is pulled in. |

## Quick start

```powershell
cd secure-chess
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -e ".[dev]"
```

### Run the server

```powershell
python -m secure_chess.server --host 127.0.0.1 --port 5050 --data-dir ./data
```

### Run two GUI clients (two new windows)

```powershell
python -m secure_chess.client --host 127.0.0.1 --port 5050
```

The Tkinter window opens on the Login screen. Register or log in, then on the
Lobby screen click **Join Lobby** (for human-vs-human) or **Start AI Game**
(for the bonus AI opponent). The Game screen renders an 8×8 board — click
your piece, then click its destination square, to send a move.

Open a second `python -m secure_chess.client` window for the opposing player,
register a different username, and click **Join Lobby** to be matched.

### Run two CLI clients (no display required)

```powershell
python -m secure_chess.client --host 127.0.0.1 --port 5050 --cli
> register alice hunter2!
ok
> play_human
ok
== game started ==
...
> move e2e4
```

```powershell
python -m secure_chess.client --host 127.0.0.1 --port 5050 --cli
> register bob s3cretpw
ok
> play_human
ok
...
```

### Play against the AI (CLI)

```powershell
python -m secure_chess.client --host 127.0.0.1 --port 5050 --cli
> register carol abcdefgh
ok
> play_ai
== game started ==
you are: white   opponent: AI(depth=3)
...
> move e2e4
```

### Verify password encryption

```powershell
Get-Content .\data\users.json
# bcrypt hashes only, no plaintext passwords

Select-String -Path .\data\users.json -Pattern "hunter2!|s3cretpw|abcdefgh"
# expected: no matches
```

### Run the test suite

```powershell
pytest -v
```

Expected: 60+ tests pass (unit + integration) in under 60 seconds.

## Project layout

```text
secure-chess/
├── pyproject.toml
├── pytest.ini
├── README.md
├── src/secure_chess/
│   ├── common/          # pure-logic library (no I/O)
│   │   ├── pieces.py    # Color, PieceType, Piece + pseudo-legal moves
│   │   ├── board.py     # Board, legality, check / mate / draws, FEN
│   │   ├── move.py      # Square + Move dataclass + UCI parsing
│   │   ├── game.py      # Game + Result
│   │   ├── crypto.py    # bcrypt wrappers
│   │   ├── user_store.py# Account + JSON-backed UserStore
│   │   ├── protocol.py  # JSON-Lines framing
│   │   ├── ai.py        # alpha-beta minimax (BONUS)
│   │   ├── errors.py    # exception hierarchy
│   │   └── log.py       # stderr logging helper
│   ├── server/          # TCP socket + threading
│   │   ├── __main__.py
│   │   ├── server.py    # accept loop + ChessServer
│   │   ├── session.py   # Session, SessionState, AISession
│   │   └── lobby.py     # Lobby + GameRegistry
│   └── client/          # both front-ends share the same wire protocol
│       ├── __main__.py  # default GUI, --cli switches to text REPL
│       ├── gui.py       # Tkinter desktop GUI
│       └── cli.py       # text-mode REPL
└── tests/
    ├── unit/            # pure-logic tests
    └── integration/     # server + clients over real sockets
```

For the full specification, plan, and design artifacts see
`specs/002-secure-chess/`.

## Protocol

Plain JSON, one object per `\n`-terminated line, over a raw TCP socket. See
`specs/002-secure-chess/contracts/wire-protocol.md` for the exact schema and
the list of error codes.

```
C -> S  {"type":"register","username":"alice","password":"hunter2!"}
S -> C  {"type":"ok"}
C -> S  {"type":"play_human"}
S -> C  {"type":"ok"}
S -> C  {"type":"game_started","you_are":"white","opponent":"bob", ...}
C -> S  {"type":"move","uci":"e2e4"}
S -> C  {"type":"ok"}
S -> C  {"type":"game_state", ...}
...
S -> C  {"type":"game_ended","result":"white_wins","reason":"checkmate"}
```
