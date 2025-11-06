"""Window management utilities for the LINE desktop app."""
from __future__ import annotations

from dataclasses import dataclass


@dataclass
class AppWindow:
    title: str


class WindowController:
    """Provides hooks for focusing or relocating application windows."""

    def focus(self, window: AppWindow) -> None:
        raise NotImplementedError("Window focus stub")
