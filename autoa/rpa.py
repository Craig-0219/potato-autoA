"""High-level mouse/keyboard automation wrappers."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable, Tuple


@dataclass
class RPASettings:
    click_pause: float = 0.1
    move_duration: float = 0.2
    type_interval: float = 0.05


class RPAController:
    """Encapsulates keyboard and mouse primitives."""

    def __init__(self, settings: RPASettings | None = None) -> None:
        self.settings = settings or RPASettings()

    def click(self, x: int, y: int, button: str = "left") -> None:
        raise NotImplementedError("Mouse click automation not yet implemented")

    def move(self, x: int, y: int, duration: float | None = None) -> None:
        raise NotImplementedError("Mouse movement not yet implemented")

    def drag_drop(self, start: Tuple[int, int], end: Tuple[int, int], duration: float | None = None) -> None:
        raise NotImplementedError("Drag and drop automation not yet implemented")

    def type_text(self, text: str) -> None:
        raise NotImplementedError("Keyboard typing not yet implemented")

    def hotkey(self, keys: Iterable[str]) -> None:
        raise NotImplementedError("Hotkey automation not yet implemented")
