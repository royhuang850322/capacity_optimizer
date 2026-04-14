"""
Centralized run-time file logging for CLI and launcher flows.

Goals:
- keep one predictable log file per run
- preserve debug-level details for support diagnostics
- allow launcher and CLI to share the same log path
"""
from __future__ import annotations

import logging
import os
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from app.runtime_paths import RuntimePaths


APP_LOGGER_NAME = "capacity_optimizer"
RUN_LOG_PATH_ENV_VAR = "CAPACITY_OPTIMIZER_RUN_LOG_PATH"


@dataclass(frozen=True)
class RunLoggingContext:
    log_path: Path
    logger_name: str = APP_LOGGER_NAME


def setup_run_file_logging(
    runtime_paths: RuntimePaths,
    *,
    run_label: str = "optimizer_run",
) -> RunLoggingContext:
    """
    Configure a debug-level file logger and return the active log path.

    If RUN_LOG_PATH_ENV_VAR is set, that path is used so launcher + CLI can
    write into one shared run log file. Otherwise, a timestamped file under
    runtime_paths.logs_dir is created.
    """
    log_path = _resolve_log_path(runtime_paths, run_label=run_label)
    logger = logging.getLogger(APP_LOGGER_NAME)
    logger.setLevel(logging.DEBUG)
    logger.propagate = False

    for handler in list(logger.handlers):
        logger.removeHandler(handler)
        try:
            handler.close()
        except Exception:
            pass

    file_handler = logging.FileHandler(log_path, mode="a", encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(
        logging.Formatter(
            fmt="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
    )
    logger.addHandler(file_handler)

    logger.debug("Logger initialized.")
    logger.debug("Log file path: %s", log_path)
    logger.debug("Workspace: %s", runtime_paths.user_workspace_dir)
    logger.debug("Install dir: %s", runtime_paths.app_install_dir)
    logger.debug("Frozen mode: %s", runtime_paths.is_frozen)

    return RunLoggingContext(log_path=log_path)


def get_app_logger() -> logging.Logger:
    return logging.getLogger(APP_LOGGER_NAME)


def format_user_error(
    *,
    code: str,
    summary: str,
    log_path: Path | None,
    details: str | None = None,
    hints: list[str] | None = None,
) -> str:
    lines = [f"  ERROR [{code}]: {summary}"]
    if details:
        lines.append(f"  Details: {details}")
    if hints:
        lines.append("  Suggested actions:")
        for hint in hints:
            lines.append(f"    - {hint}")
    if log_path:
        lines.append(f"  Log file: {log_path}")
    return "\n".join(lines)


def _resolve_log_path(runtime_paths: RuntimePaths, *, run_label: str) -> Path:
    override = os.environ.get(RUN_LOG_PATH_ENV_VAR, "").strip()
    if override:
        candidate = Path(override).expanduser()
    else:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        candidate = runtime_paths.logs_dir / f"{run_label}_{timestamp}.log"

    if candidate.suffix.lower() != ".log":
        candidate = candidate / f"{run_label}.log"

    candidate.parent.mkdir(parents=True, exist_ok=True)
    return candidate.resolve()
