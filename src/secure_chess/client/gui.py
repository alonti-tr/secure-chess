"""Tkinter GUI client for the secure-chess server.

Three screens managed by a single Tk root: Login → Lobby → Game. The GUI speaks
exactly the same JSON-Lines wire protocol as the CLI client in
`secure_chess.client.cli`, so any combination of GUI and CLI clients can play
against each other or against the AI.

Threading: a single background reader thread drains the TCP socket and pushes
every server message onto a `queue.Queue`. The Tk main loop polls that queue
every 50ms via `root.after`, ensuring every widget mutation happens on the Tk
main thread.
"""

from __future__ import annotations

import queue
import socket
import threading
import tkinter as tk
from tkinter import messagebox, ttk
from typing import Any, Dict, List, Optional, Tuple

from secure_chess.common import protocol
from secure_chess.common.errors import ConnectionClosed, ProtocolError


SQUARE_SIZE = 64
BOARD_MARGIN = 18
BOARD_PIXELS = SQUARE_SIZE * 8

LIGHT_COLOR = "#EEEED2"
DARK_COLOR = "#769656"
SELECTED_COLOR = "#F7EC74"
LAST_MOVE_COLOR = "#F0D75C"

PIECE_TO_UNICODE = {
    "K": "\u2654", "Q": "\u2655", "R": "\u2656", "B": "\u2657", "N": "\u2658", "P": "\u2659",
    "k": "\u265A", "q": "\u265B", "r": "\u265C", "b": "\u265D", "n": "\u265E", "p": "\u265F",
}


def parse_fen_placement(fen_short: str) -> List[List[Optional[str]]]:
    """Parse the placement field of a FEN-short string into an 8x8 grid.

    Returns a 2D list `grid[file][rank_from_top]` where:
    - file 0 = a-file, file 7 = h-file
    - rank_from_top 0 = rank 8, rank_from_top 7 = rank 1
    - each cell is either `None` or a one-character FEN piece code
      (uppercase = white, lowercase = black)
    """
    placement = fen_short.split()[0] if fen_short else ""
    grid: List[List[Optional[str]]] = [[None] * 8 for _ in range(8)]
    for rank_idx, row in enumerate(placement.split("/")):
        if rank_idx >= 8:
            break
        f = 0
        for ch in row:
            if ch.isdigit():
                f += int(ch)
            elif f < 8:
                grid[f][rank_idx] = ch
                f += 1
    return grid


def square_to_grid(sq: str) -> Optional[Tuple[int, int]]:
    """Convert algebraic notation (e.g. 'e2') to grid coords `(file, rank_from_top)`."""
    if not isinstance(sq, str) or len(sq) != 2:
        return None
    f = ord(sq[0]) - ord("a")
    if not (0 <= f < 8):
        return None
    try:
        r = int(sq[1])
    except ValueError:
        return None
    if not (1 <= r <= 8):
        return None
    return f, 8 - r


def grid_to_square(file: int, rank_from_top: int) -> str:
    """Convert grid coords to algebraic notation (e.g. (4,6) -> 'e2')."""
    return f"{chr(ord('a') + file)}{8 - rank_from_top}"


class ServerConnection:
    """Background TCP reader publishing every message onto a thread-safe queue.

    The queue is consumed by `ChessGui._poll_events` on the Tk main thread, so
    no Tk widget is ever touched by the reader thread.
    """

    def __init__(self, host: str, port: int, timeout: float = 5.0) -> None:
        self._sock = socket.create_connection((host, port), timeout=timeout)
        self._sock.settimeout(None)
        self._reader = protocol.make_reader(self._sock)
        self._send_lock = threading.Lock()
        self.events: "queue.Queue[Dict[str, Any]]" = queue.Queue()
        self._stopping = threading.Event()
        self._reader_thread = threading.Thread(
            target=self._read_loop,
            name="gui-reader",
            daemon=True,
        )
        self._reader_thread.start()

    def send(self, msg: Dict[str, Any]) -> None:
        with self._send_lock:
            protocol.send_message(self._sock, msg)

    def _read_loop(self) -> None:
        try:
            while not self._stopping.is_set():
                msg = protocol.recv_message(self._reader)
                self.events.put(msg)
        except ConnectionClosed:
            self.events.put({"type": "_disconnected", "reason": "server closed the connection"})
        except ProtocolError as exc:
            self.events.put({"type": "_disconnected", "reason": f"protocol error: {exc}"})

    def close(self) -> None:
        self._stopping.set()
        try:
            self._sock.shutdown(socket.SHUT_RDWR)
        except OSError:
            pass
        try:
            self._sock.close()
        except OSError:
            pass


class ChessGui:
    """Top-level Tk application managing the three screens and the server connection."""

    def __init__(self, host: str, port: int) -> None:
        self.host = host
        self.port = port
        self.conn: Optional[ServerConnection] = None

        self.username: Optional[str] = None
        self.pending_auth_username: Optional[str] = None
        self.pending_action: Optional[str] = None
        self.waiting_in_lobby: bool = False

        self.you_are: Optional[str] = None
        self.opponent: Optional[str] = None
        self.board_grid: List[List[Optional[str]]] = [[None] * 8 for _ in range(8)]
        self.to_move: str = "white"
        self.in_check: bool = False
        self.last_move: Optional[str] = None
        self.selected_square: Optional[str] = None
        self.move_log: List[str] = []
        self.game_active: bool = False

        self.root = tk.Tk()
        self.root.title("Secure Chess")
        self.root.geometry("780x680")
        self.root.minsize(780, 680)

        self.container = tk.Frame(self.root)
        self.container.pack(fill="both", expand=True)

        self.login_frame = LoginFrame(self.container, self)
        self.lobby_frame = LobbyFrame(self.container, self)
        self.game_frame = GameFrame(self.container, self)

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.show_login()

    def _hide_all(self) -> None:
        for f in (self.login_frame, self.lobby_frame, self.game_frame):
            f.pack_forget()

    def show_login(self) -> None:
        self._hide_all()
        self.login_frame.pack(fill="both", expand=True)
        self.login_frame.focus_username()

    def show_lobby(self) -> None:
        self._hide_all()
        self.lobby_frame.refresh(self.username)
        self.lobby_frame.pack(fill="both", expand=True)

    def show_game(self) -> None:
        self._hide_all()
        self.game_frame.refresh_all()
        self.game_frame.pack(fill="both", expand=True)

    def ensure_connected(self) -> bool:
        if self.conn is not None:
            return True
        try:
            self.conn = ServerConnection(self.host, self.port)
        except OSError as exc:
            messagebox.showerror(
                "Connection failed",
                f"Cannot connect to {self.host}:{self.port}\n{exc}",
            )
            return False
        self.root.after(50, self._poll_events)
        return True

    def _poll_events(self) -> None:
        if self.conn is None:
            return
        try:
            while True:
                msg = self.conn.events.get_nowait()
                self._handle_message(msg)
        except queue.Empty:
            pass
        if self.conn is not None:
            self.root.after(50, self._poll_events)

    def _handle_message(self, msg: Dict[str, Any]) -> None:
        t = msg.get("type")
        if t == "_disconnected":
            self._on_disconnected(msg.get("reason", "connection closed"))
            return
        if t == "ok":
            self._on_ok(msg)
        elif t == "error":
            self._on_error(msg)
        elif t == "game_started":
            self._on_game_started(msg)
        elif t == "game_state":
            self._on_game_state(msg)
        elif t == "game_ended":
            self._on_game_ended(msg)

    def _on_ok(self, _msg: Dict[str, Any]) -> None:
        action = self.pending_action
        self.pending_action = None
        if self.pending_auth_username is not None:
            self.username = self.pending_auth_username
            self.pending_auth_username = None
            self.show_lobby()
            return
        if action == "play_human":
            self.waiting_in_lobby = True
            self.lobby_frame.refresh_waiting()
        elif action == "cancel_lobby":
            self.waiting_in_lobby = False
            self.lobby_frame.refresh_waiting()

    def _on_error(self, msg: Dict[str, Any]) -> None:
        code = msg.get("code", "?")
        message = msg.get("message", "?")
        action = self.pending_action
        self.pending_action = None
        if self.pending_auth_username is not None:
            self.pending_auth_username = None
            messagebox.showerror(f"Login error: {code}", message)
            return
        messagebox.showerror(f"Error: {code}", message)
        if action == "cancel_lobby":
            self.lobby_frame.refresh_waiting()
            return
        if self.game_active:
            self.selected_square = None
            self.game_frame.refresh_board()

    def _on_game_started(self, msg: Dict[str, Any]) -> None:
        self.you_are = msg.get("you_are")
        self.opponent = msg.get("opponent")
        self.board_grid = parse_fen_placement(msg.get("board", ""))
        self.to_move = msg.get("to_move", "white")
        self.in_check = False
        self.last_move = None
        self.selected_square = None
        self.move_log = []
        self.game_active = True
        self.waiting_in_lobby = False
        self.pending_action = None
        self.show_game()

    def _on_game_state(self, msg: Dict[str, Any]) -> None:
        board = msg.get("board")
        if board:
            self.board_grid = parse_fen_placement(board)
        self.to_move = msg.get("to_move", self.to_move)
        self.in_check = bool(msg.get("in_check", False))
        last = msg.get("last_move")
        if last:
            self.last_move = last
            self.move_log.append(last)
        self.selected_square = None
        self.game_frame.refresh_all()

    def _on_game_ended(self, msg: Dict[str, Any]) -> None:
        result = msg.get("result", "?")
        reason = msg.get("reason", "?")
        self.game_active = False
        messagebox.showinfo("Game ended", f"Result: {result}\nReason: {reason}")
        self.show_lobby()

    def _on_disconnected(self, reason: str) -> None:
        self.conn = None
        was_authed = self.username is not None
        self._reset_state()
        if was_authed:
            messagebox.showwarning("Disconnected", reason)
        self.show_login()

    def _reset_state(self) -> None:
        if self.conn is not None:
            self.conn.close()
            self.conn = None
        self.username = None
        self.pending_auth_username = None
        self.pending_action = None
        self.waiting_in_lobby = False
        self.you_are = None
        self.opponent = None
        self.selected_square = None
        self.game_active = False
        self.move_log = []
        self.board_grid = [[None] * 8 for _ in range(8)]

    def send_command(self, msg: Dict[str, Any]) -> bool:
        if self.conn is None:
            return False
        action_type = msg.get("type")
        if action_type in ("play_human", "play_ai", "cancel_lobby"):
            self.pending_action = action_type
        try:
            self.conn.send(msg)
            return True
        except (OSError, ConnectionClosed) as exc:
            self.pending_action = None
            messagebox.showerror("Send failed", str(exc))
            self._on_disconnected(str(exc))
            return False

    def on_close(self) -> None:
        if self.conn is not None:
            try:
                self.conn.send({"type": "quit"})
            except (OSError, ConnectionClosed):
                pass
            self.conn.close()
            self.conn = None
        self.root.destroy()

    def run(self) -> int:
        self.root.mainloop()
        return 0


class LoginFrame(tk.Frame):
    """First screen: connect + authenticate (register or login)."""

    def __init__(self, parent: tk.Widget, app: ChessGui) -> None:
        super().__init__(parent)
        self.app = app

        title = tk.Label(self, text="Secure Chess", font=("Arial", 32, "bold"))
        title.pack(pady=(90, 6))

        subtitle = tk.Label(
            self,
            text="Sign in or create an account to start playing",
            font=("Arial", 12),
            fg="#444",
        )
        subtitle.pack(pady=(0, 30))

        form = tk.Frame(self)
        form.pack()

        tk.Label(form, text="Username:", font=("Arial", 11)).grid(
            row=0, column=0, sticky="e", padx=5, pady=6
        )
        self.user_entry = tk.Entry(form, font=("Arial", 11), width=26)
        self.user_entry.grid(row=0, column=1, padx=5, pady=6)

        tk.Label(form, text="Password:", font=("Arial", 11)).grid(
            row=1, column=0, sticky="e", padx=5, pady=6
        )
        self.pwd_entry = tk.Entry(form, font=("Arial", 11), width=26, show="\u2022")
        self.pwd_entry.grid(row=1, column=1, padx=5, pady=6)
        self.pwd_entry.bind("<Return>", lambda _e: self._login())

        btn_row = tk.Frame(self)
        btn_row.pack(pady=22)

        tk.Button(
            btn_row, text="Login", font=("Arial", 11), width=14, command=self._login
        ).pack(side="left", padx=6)
        tk.Button(
            btn_row, text="Register", font=("Arial", 11), width=14, command=self._register
        ).pack(side="left", padx=6)

        self.hint = tk.Label(
            self,
            text=f"Server: {app.host}:{app.port}",
            font=("Arial", 9),
            fg="#888",
        )
        self.hint.pack(side="bottom", pady=8)

    def focus_username(self) -> None:
        self.user_entry.focus_set()

    def _validate(self) -> Optional[Tuple[str, str]]:
        u = self.user_entry.get().strip()
        p = self.pwd_entry.get()
        if not u or not p:
            messagebox.showerror("Missing fields", "Both username and password are required.")
            return None
        return u, p

    def _login(self) -> None:
        vals = self._validate()
        if vals is None:
            return
        if not self.app.ensure_connected():
            return
        u, p = vals
        self.app.pending_auth_username = u
        self.app.send_command({"type": "login", "username": u, "password": p})

    def _register(self) -> None:
        vals = self._validate()
        if vals is None:
            return
        if not self.app.ensure_connected():
            return
        u, p = vals
        self.app.pending_auth_username = u
        self.app.send_command({"type": "register", "username": u, "password": p})


class LobbyFrame(tk.Frame):
    """Second screen: choose to play another human or play against the AI.

    Tracks two visual sub-states driven by `ChessGui.waiting_in_lobby`:
    - normal: both Join Lobby and Start AI Game are enabled.
    - waiting: a yellow banner with "Waiting for opponent..." is shown and the
      two action buttons are greyed out; only Cancel + Logout remain clickable.
    """

    def __init__(self, parent: tk.Widget, app: ChessGui) -> None:
        super().__init__(parent)
        self.app = app

        self.title_label = tk.Label(self, text="", font=("Arial", 22, "bold"))
        self.title_label.pack(pady=(40, 6))

        tk.Label(
            self,
            text="Choose your opponent",
            font=("Arial", 12),
            fg="#444",
        ).pack(pady=(0, 12))

        self.waiting_banner = tk.Frame(self, bg="#FFF4C2", bd=1, relief="solid")
        self.waiting_label = tk.Label(
            self.waiting_banner,
            text="\u23F3  Waiting for another player to join the lobby...",
            font=("Arial", 11, "bold"),
            bg="#FFF4C2",
            fg="#664D03",
        )
        self.waiting_label.pack(side="left", padx=12, pady=8)
        self.cancel_btn = tk.Button(
            self.waiting_banner,
            text="Cancel",
            font=("Arial", 10),
            command=self._cancel_lobby,
        )
        self.cancel_btn.pack(side="right", padx=10, pady=6)

        human_frame = tk.LabelFrame(
            self, text=" Play Human ", font=("Arial", 11, "bold"), padx=18, pady=14
        )
        human_frame.pack(pady=10, padx=80, fill="x")
        tk.Label(
            human_frame,
            text=(
                "Wait in the lobby until another authenticated player joins."
                "\nColors are assigned automatically (first to wait plays White)."
            ),
            justify="left",
        ).pack(anchor="w", pady=(0, 8))
        self.join_btn = tk.Button(
            human_frame,
            text="Join Lobby",
            font=("Arial", 11),
            width=18,
            command=self._play_human,
        )
        self.join_btn.pack()

        ai_frame = tk.LabelFrame(
            self, text=" Play AI (bonus) ", font=("Arial", 11, "bold"), padx=18, pady=14
        )
        ai_frame.pack(pady=10, padx=80, fill="x")

        opts = tk.Frame(ai_frame)
        opts.pack(anchor="w", pady=(0, 8))

        tk.Label(opts, text="Your color:").grid(row=0, column=0, sticky="e", padx=4, pady=2)
        self.color_var = tk.StringVar(value="white")
        tk.Radiobutton(opts, text="White", variable=self.color_var, value="white").grid(
            row=0, column=1, sticky="w"
        )
        tk.Radiobutton(opts, text="Black", variable=self.color_var, value="black").grid(
            row=0, column=2, sticky="w"
        )

        tk.Label(opts, text="AI search depth:").grid(row=1, column=0, sticky="e", padx=4, pady=2)
        self.depth_var = tk.IntVar(value=3)
        ttk.Combobox(
            opts,
            textvariable=self.depth_var,
            values=[1, 2, 3, 4],
            width=4,
            state="readonly",
        ).grid(row=1, column=1, sticky="w")
        tk.Label(opts, text="(higher = stronger but slower)", fg="#666").grid(
            row=1, column=2, columnspan=2, sticky="w", padx=(8, 0)
        )

        self.ai_btn = tk.Button(
            ai_frame,
            text="Start AI Game",
            font=("Arial", 11),
            width=18,
            command=self._play_ai,
        )
        self.ai_btn.pack()

        tk.Button(self, text="Logout", font=("Arial", 10), command=self._logout).pack(
            side="bottom", pady=16
        )

    def refresh(self, username: Optional[str]) -> None:
        self.title_label.config(text=f"Welcome, {username or 'guest'}")
        self.refresh_waiting()

    def refresh_waiting(self) -> None:
        if self.app.waiting_in_lobby:
            self.waiting_banner.pack(after=self.title_label, padx=80, pady=(4, 16), fill="x")
            self.join_btn.config(state="disabled")
            self.ai_btn.config(state="disabled")
        else:
            self.waiting_banner.pack_forget()
            self.join_btn.config(state="normal")
            self.ai_btn.config(state="normal")

    def _play_human(self) -> None:
        if self.app.waiting_in_lobby:
            return
        self.app.send_command({"type": "play_human"})

    def _play_ai(self) -> None:
        if self.app.waiting_in_lobby:
            messagebox.showinfo(
                "Already in lobby",
                "You're currently waiting for a human opponent.\n"
                "Click Cancel first if you'd rather play against the AI.",
            )
            return
        self.app.send_command(
            {
                "type": "play_ai",
                "color": self.color_var.get(),
                "depth": int(self.depth_var.get()),
            }
        )

    def _cancel_lobby(self) -> None:
        if not self.app.waiting_in_lobby:
            return
        self.app.send_command({"type": "cancel_lobby"})

    def _logout(self) -> None:
        if self.app.conn is not None:
            try:
                self.app.conn.send({"type": "quit"})
            except (OSError, ConnectionClosed):
                pass
        self.app._reset_state()
        self.app.show_login()


class GameFrame(tk.Frame):
    """Third screen: status bar, clickable 8x8 board, move log, resign button."""

    def __init__(self, parent: tk.Widget, app: ChessGui) -> None:
        super().__init__(parent)
        self.app = app

        self.status_label = tk.Label(
            self, text="", font=("Arial", 11), anchor="w", justify="left"
        )
        self.status_label.pack(fill="x", padx=12, pady=(10, 4))

        body = tk.Frame(self)
        body.pack(padx=12, pady=4, fill="both", expand=True)

        canvas_width = BOARD_PIXELS + BOARD_MARGIN * 2
        canvas_height = BOARD_PIXELS + BOARD_MARGIN * 2
        self.canvas = tk.Canvas(
            body,
            width=canvas_width,
            height=canvas_height,
            bg="#f5f5f0",
            highlightthickness=0,
        )
        self.canvas.pack(side="left")
        self.canvas.bind("<Button-1>", self._on_click)

        side = tk.Frame(body)
        side.pack(side="left", padx=(16, 0), fill="y")

        tk.Label(side, text="Move log", font=("Arial", 10, "bold")).pack(anchor="w")
        log_box = tk.Frame(side)
        log_box.pack(pady=(2, 10))
        self.log_text = tk.Text(
            log_box, width=18, height=22, state="disabled", font=("Courier", 10)
        )
        self.log_text.pack(side="left")
        scroll = tk.Scrollbar(log_box, command=self.log_text.yview)
        scroll.pack(side="right", fill="y")
        self.log_text.config(yscrollcommand=scroll.set)

        tk.Button(side, text="Resign", width=16, command=self._resign).pack(pady=(2, 4))
        tk.Button(side, text="Quit to Lobby", width=16, command=self._lobby).pack()

    def refresh_all(self) -> None:
        self.refresh_status()
        self.refresh_board()
        self.refresh_log()

    def refresh_status(self) -> None:
        you = self.app.you_are or "?"
        opp = self.app.opponent or "?"
        to_move = self.app.to_move
        last = self.app.last_move or "—"
        your_turn = to_move == self.app.you_are
        turn_marker = "  ← YOUR TURN" if your_turn else ""
        check_text = "    \u26A0 CHECK" if self.app.in_check else ""
        text = (
            f"You: {you}    Opponent: {opp}    To move: {to_move}{turn_marker}{check_text}\n"
            f"Last move: {last}"
        )
        self.status_label.config(text=text, fg="#A00" if self.app.in_check else "#000")

    def _board_orientation_flipped(self) -> bool:
        return self.app.you_are == "black"

    def _logical_to_display(self, f: int, rt: int) -> Tuple[int, int]:
        if self._board_orientation_flipped():
            return 7 - f, 7 - rt
        return f, rt

    def _display_to_logical(self, df: int, drt: int) -> Tuple[int, int]:
        if self._board_orientation_flipped():
            return 7 - df, 7 - drt
        return df, drt

    def refresh_board(self) -> None:
        c = self.canvas
        c.delete("all")
        s = SQUARE_SIZE
        m = BOARD_MARGIN

        for f in range(8):
            for rt in range(8):
                df, drt = self._logical_to_display(f, rt)
                x0 = m + df * s
                y0 = m + drt * s
                light = (f + rt) % 2 == 0
                color = LIGHT_COLOR if light else DARK_COLOR
                c.create_rectangle(x0, y0, x0 + s, y0 + s, fill=color, outline="")

        if self.app.last_move and len(self.app.last_move) >= 4:
            for sq in (self.app.last_move[:2], self.app.last_move[2:4]):
                fr = square_to_grid(sq)
                if fr is None:
                    continue
                f, rt = fr
                df, drt = self._logical_to_display(f, rt)
                x0 = m + df * s
                y0 = m + drt * s
                c.create_rectangle(
                    x0 + 2, y0 + 2, x0 + s - 2, y0 + s - 2,
                    outline=LAST_MOVE_COLOR, width=3,
                )

        if self.app.selected_square is not None:
            fr = square_to_grid(self.app.selected_square)
            if fr is not None:
                f, rt = fr
                df, drt = self._logical_to_display(f, rt)
                x0 = m + df * s
                y0 = m + drt * s
                c.create_rectangle(x0, y0, x0 + s, y0 + s, fill=SELECTED_COLOR, outline="")

        for f in range(8):
            for rt in range(8):
                piece = self.app.board_grid[f][rt]
                if piece is None:
                    continue
                df, drt = self._logical_to_display(f, rt)
                cx = m + df * s + s // 2
                cy = m + drt * s + s // 2
                glyph = PIECE_TO_UNICODE.get(piece, piece)
                if piece.isupper():
                    c.create_text(cx + 1, cy + 1, text=glyph, font=("Arial", 40), fill="#000")
                    c.create_text(cx, cy, text=glyph, font=("Arial", 40), fill="#FFF")
                else:
                    c.create_text(cx, cy, text=glyph, font=("Arial", 40), fill="#000")

        for i in range(8):
            df, _ = self._logical_to_display(i, 0)
            fx = m + df * s + s // 2
            c.create_text(
                fx, m + 8 * s + 9,
                text=chr(ord("a") + i),
                font=("Arial", 10), fill="#333",
            )
        for i in range(8):
            _, drt = self._logical_to_display(0, i)
            ry = m + drt * s + s // 2
            c.create_text(m - 10, ry, text=str(8 - i), font=("Arial", 10), fill="#333")

    def refresh_log(self) -> None:
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", tk.END)
        for i, move in enumerate(self.app.move_log):
            move_no = i // 2 + 1
            if i % 2 == 0:
                self.log_text.insert(tk.END, f"{move_no:>3}. {move:<6} ")
            else:
                self.log_text.insert(tk.END, f"{move}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")

    def _on_click(self, event: tk.Event) -> None:
        if not self.app.game_active:
            return
        x = event.x - BOARD_MARGIN
        y = event.y - BOARD_MARGIN
        if not (0 <= x < BOARD_PIXELS and 0 <= y < BOARD_PIXELS):
            return
        df = x // SQUARE_SIZE
        drt = y // SQUARE_SIZE
        f, rt = self._display_to_logical(df, drt)
        sq = grid_to_square(f, rt)

        if self.app.to_move != self.app.you_are:
            return

        piece_here = self.app.board_grid[f][rt]
        own_piece_here = piece_here is not None and self._is_own_piece(piece_here)

        if self.app.selected_square is None:
            if own_piece_here:
                self.app.selected_square = sq
                self.refresh_board()
            return

        if sq == self.app.selected_square:
            self.app.selected_square = None
            self.refresh_board()
            return

        if own_piece_here:
            self.app.selected_square = sq
            self.refresh_board()
            return

        origin = self.app.selected_square
        uci = origin + sq

        origin_fr = square_to_grid(origin)
        if origin_fr is not None:
            of, ort = origin_fr
            moving = self.app.board_grid[of][ort]
            if moving == "P" and rt == 0:
                uci += "q"
            elif moving == "p" and rt == 7:
                uci += "q"

        self.app.send_command({"type": "move", "uci": uci})

    def _is_own_piece(self, piece: str) -> bool:
        if self.app.you_are == "white":
            return piece.isupper()
        if self.app.you_are == "black":
            return piece.islower()
        return False

    def _resign(self) -> None:
        if not self.app.game_active:
            return
        if messagebox.askyesno("Resign", "Are you sure you want to resign?"):
            self.app.send_command({"type": "resign"})

    def _lobby(self) -> None:
        if not self.app.game_active:
            self.app.show_lobby()
            return
        if messagebox.askyesno("Quit Game", "Quitting will count as resigning. Continue?"):
            self.app.send_command({"type": "resign"})


def run_gui(host: str, port: int) -> int:
    """Entry point used by `secure_chess.client.__main__`."""
    return ChessGui(host, port).run()
