from __future__ import annotations

import logging
import sys


_FORMAT = "%(asctime)s %(levelname)s %(name)s | %(message)s"
_configured = False


def get_logger(name: str, level: str = "INFO") -> logging.Logger:
    global _configured
    if not _configured:
        handler = logging.StreamHandler(sys.stderr)
        handler.setFormatter(logging.Formatter(_FORMAT))
        root = logging.getLogger("secure_chess")
        root.addHandler(handler)
        root.setLevel(level)
        _configured = True
    return logging.getLogger(f"secure_chess.{name}" if not name.startswith("secure_chess") else name)
