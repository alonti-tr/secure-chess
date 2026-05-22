from __future__ import annotations

import argparse
import sys


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="secure-chess-client")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=5050)
    parser.add_argument(
        "--cli",
        action="store_true",
        help="Use the text-mode REPL instead of the Tkinter GUI",
    )
    args = parser.parse_args(argv)

    if args.cli:
        from secure_chess.client.cli import run_cli

        return run_cli(args.host, args.port)

    try:
        from secure_chess.client.gui import run_gui
    except ImportError as exc:
        print(
            f"GUI is unavailable ({exc}); falling back to CLI. "
            f"Re-run with --cli to suppress this warning.",
            file=sys.stderr,
        )
        from secure_chess.client.cli import run_cli

        return run_cli(args.host, args.port)

    return run_gui(args.host, args.port)


if __name__ == "__main__":
    sys.exit(main())
