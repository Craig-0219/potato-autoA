"""Command-line entry point for AUTOA RPA."""
from __future__ import annotations

import argparse
from pathlib import Path

from autoa.flow import FlowExecutor
from autoa.rpa import RPAController
from autoa.vision import TemplateMatcher
from autoa.guards import GuardRail


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Automate LINE desktop interactions")
    parser.add_argument("mode", choices=["run", "dryrun"], help="Execution mode")
    parser.add_argument("task", type=Path, help="Path to a YAML task file")
    parser.add_argument("--config", type=Path, default=Path("config.yaml"), help="Config file")
    parser.add_argument("--from", dest="from_step", default=None, help="Resume from a specific step label")
    return parser.parse_args()


def build_executor() -> FlowExecutor:
    rpa = RPAController()
    matcher = TemplateMatcher()
    guards = GuardRail()
    return FlowExecutor(rpa=rpa, matcher=matcher, guards=guards)


def main() -> None:
    args = parse_args()
    executor = build_executor()
    dry_run = args.mode == "dryrun"

    task_path = args.task
    if not task_path.exists():
        raise FileNotFoundError(f"Task file not found: {task_path}")

    # Flow execution is not implemented; raise the same error in one place for clarity.
    executor.run(task_path, dry_run=dry_run)


if __name__ == "__main__":
    main()
