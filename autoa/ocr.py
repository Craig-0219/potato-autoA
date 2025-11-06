"""Optional Tesseract OCR integration."""
from __future__ import annotations

from pathlib import Path


class OCRClient:
    """Performs lightweight OCR for validation steps."""

    def __init__(self, tesseract_cmd: Path | None = None) -> None:
        self.tesseract_cmd = tesseract_cmd

    def read_text(self, image_path: Path) -> str:
        raise NotImplementedError("OCR stub")
