"""CLI entry point: `python -m secure_chess.server`."""

from __future__ import annotations

import argparse
import sys
import threading

from secure_chess.common.log import get_logger
from secure_chess.server.server import ChessServer


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="secure-chess-server")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=5050)
    parser.add_argument("--data-dir", default="data")
    parser.add_argument("--ai-depth", type=int, default=3)
    parser.add_argument("--max-clients", type=int, default=32)
    parser.add_argument("--log-level", default="INFO",
                        choices=["DEBUG", "INFO", "WARNING", "ERROR"])
    args = parser.parse_args(argv)

    get_logger("main", level=args.log_level)

    server = ChessServer(
        host=args.host,
        port=args.port,
        data_dir=args.data_dir,
        ai_depth=args.ai_depth,
        max_clients=args.max_clients,
    )
    server.start()

    ev = threading.Event()
    try:
        ev.wait()
    except KeyboardInterrupt:
        pass
    finally:
        server.stop()
    return 0


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
