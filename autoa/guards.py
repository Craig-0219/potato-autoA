"""Rate limiting and human-like pacing helpers."""
from __future__ import annotations

from dataclasses import dataclass
from random import uniform
from time import sleep


@dataclass
class DelayWindow:
    min_seconds: float
    max_seconds: float

    def sample(self) -> float:
        return uniform(self.min_seconds, self.max_seconds)


class GuardRail:
    """Encapsulates throttling policies and blacklists."""

    def __init__(self, *, click_delay: DelayWindow | None = None) -> None:
        self.click_delay = click_delay or DelayWindow(0.06, 0.12)

    def pause_for_click(self) -> None:
        sleep(self.click_delay.sample())
