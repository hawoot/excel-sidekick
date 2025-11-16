"""Custom exceptions for Excel Sidekick."""


class ExcelSidekickError(Exception):
    """Base exception for all Excel Sidekick errors."""

    pass


class ExcelConnectionError(ExcelSidekickError):
    """Raised when Excel connection fails."""

    pass


class WorkbookNotFoundError(ExcelSidekickError):
    """Raised when workbook is not found or not open."""

    pass


class SheetNotFoundError(ExcelSidekickError):
    """Raised when specified sheet doesn't exist."""

    pass


class InvalidRangeError(ExcelSidekickError):
    """Raised when range specification is invalid."""

    pass


class DependencyGraphError(ExcelSidekickError):
    """Raised when dependency graph operations fail."""

    pass


class CacheError(ExcelSidekickError):
    """Raised when cache operations fail."""

    pass


class AnnotationError(ExcelSidekickError):
    """Raised when annotation operations fail."""

    pass


class LLMProviderError(ExcelSidekickError):
    """Raised when LLM provider operations fail."""

    pass


class ConfigurationError(ExcelSidekickError):
    """Raised when configuration is invalid or missing."""

    pass


class AgentError(ExcelSidekickError):
    """Raised when agent execution fails."""

    pass
