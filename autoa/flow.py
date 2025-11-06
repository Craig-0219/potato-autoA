"""YAML-driven task orchestration."""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

import yaml


@dataclass
class FlowContext:
    task_file: Path
    variables: dict[str, Any]


class FlowExecutor:
    """Executes high-level steps described in YAML."""

    def __init__(self, *, rpa, matcher, guards) -> None:
        self.rpa = rpa
        self.matcher = matcher
        self.guards = guards

    def load(self, path: Path) -> dict[str, Any]:
        with path.open("r", encoding="utf-8") as stream:
            return yaml.safe_load(stream)

    def run(self, path: Path, *, dry_run: bool = False) -> None:
        raise NotImplementedError("Flow execution not yet implemented")
