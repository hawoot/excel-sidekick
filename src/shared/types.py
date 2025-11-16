"""Shared type definitions and enums for Excel Sidekick."""

from enum import Enum
from typing import Literal


class TraceDirection(str, Enum):
    """Direction for dependency tracing."""

    UPSTREAM = "upstream"  # What feeds into this cell
    DOWNSTREAM = "downstream"  # What uses this cell
    BOTH = "both"  # Both directions


class SnapshotFormat(str, Enum):
    """Format for snapshot output."""

    MARKDOWN = "markdown"
    ASCII = "ascii"
    JSON = "json"


class LogLevel(str, Enum):
    """Logging levels."""

    DEBUG = "DEBUG"
    INFO = "INFO"
    WARNING = "WARNING"
    ERROR = "ERROR"


class MatchType(str, Enum):
    """Type of annotation match."""

    EXACT = "exact"  # Exact range match
    SUPERSET = "superset"  # Query contains annotation
    SUBSET = "subset"  # Query inside annotation
    PARTIAL = "partial"  # Partial overlap
    ADJACENT = "adjacent"  # Adjacent to annotation


# Type aliases for clarity
CellAddress = str  # e.g., "A1", "Sheet1!B5"
RangeAddress = str  # e.g., "A1:B10", "Sheet1!A1:B10"
SheetName = str
FormulaString = str
