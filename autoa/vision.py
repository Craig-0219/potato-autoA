"""Computer vision helpers for template matching."""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Sequence


@dataclass
class TemplateConfig:
    templates: Sequence[Path]
    threshold: float = 0.92
    timeout_sec: float = 10.0


class TemplateMatcher:
    """Loads templates and performs best-effort OpenCV matching."""

    def __init__(self, search_paths: Iterable[Path] | None = None) -> None:
        self.search_paths = list(search_paths or [])

    def locate(self, config: TemplateConfig) -> tuple[int, int] | None:
        raise NotImplementedError("Template locate stub")
