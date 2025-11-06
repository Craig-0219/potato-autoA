"""Utility helpers for logging, retries, and timekeeping."""
from __future__ import annotations

from pathlib import Path
import logging


def setup_logging(log_path: Path) -> logging.Logger:
    """Configure a basic rotating logger placeholder."""
    logger = logging.getLogger("autoa")
    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)
    handler = logging.FileHandler(log_path, encoding="utf-8")
    formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    return logger
