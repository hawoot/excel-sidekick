"""Logging configuration for Excel Sidekick."""

import logging
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional

try:
    from rich.logging import RichHandler
    from rich.traceback import install as install_rich_traceback
    RICH_AVAILABLE = True
except ImportError:
    RICH_AVAILABLE = False


def setup_logging(
    level: str = "INFO",
    log_file: Optional[str] = None,
    log_format: Optional[str] = None,
) -> logging.Logger:
    """
    Set up logging configuration for Excel Sidekick.

    Console output and rich tracebacks are always enabled.
    The {date} placeholder in log_file is replaced with current date (YYYY-MM-DD).

    Args:
        level: Logging level (DEBUG, INFO, WARNING, ERROR)
        log_file: Path to log file with optional {date} placeholder
        log_format: Custom log format string

    Returns:
        Configured logger instance
    """
    # Install rich traceback globally for better error formatting
    if RICH_AVAILABLE:
        install_rich_traceback(show_locals=True)

    # Default format if not specified
    if log_format is None:
        log_format = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"

    # Create root logger
    logger = logging.getLogger("excel_sidekick")
    logger.setLevel(getattr(logging, level.upper()))

    # Remove existing handlers to avoid duplicates
    logger.handlers.clear()

    # Console handler (always enabled, use RichHandler if available)
    if RICH_AVAILABLE:
        console_handler = RichHandler(
            rich_tracebacks=True,
            tracebacks_show_locals=True,
            markup=True,
        )
    else:
        console_handler = logging.StreamHandler(sys.stdout)
        formatter = logging.Formatter(log_format)
        console_handler.setFormatter(formatter)

    console_handler.setLevel(getattr(logging, level.upper()))
    logger.addHandler(console_handler)

    # File handler
    if log_file:
        # Replace {date} placeholder with current date
        current_date = datetime.now().strftime("%Y-%m-%d")
        log_file_resolved = log_file.replace("{date}", current_date)

        log_path = Path(log_file_resolved)
        log_path.parent.mkdir(parents=True, exist_ok=True)

        file_handler = logging.FileHandler(log_file_resolved)
        file_handler.setLevel(getattr(logging, level.upper()))

        # File handler always uses standard formatter (not RichHandler)
        file_formatter = logging.Formatter(log_format)
        file_handler.setFormatter(file_formatter)
        logger.addHandler(file_handler)

    # Prevent propagation to root logger
    logger.propagate = False

    return logger


def get_logger(name: str) -> logging.Logger:
    """
    Get a logger instance for a specific module.

    Args:
        name: Name of the module (typically __name__)

    Returns:
        Logger instance
    """
    return logging.getLogger(f"excel_sidekick.{name}")
