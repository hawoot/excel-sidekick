"""Configuration loader for Excel Sidekick."""

import os
from pathlib import Path
from typing import Any, Dict, Optional, Union

import yaml
from pydantic import BaseModel, Field, field_validator

from src.shared.exceptions import ConfigurationError
from src.shared.types import DependencyMode, LogLevel, SnapshotFormat, TraceDirection


# Project root for resolving relative paths
def get_project_root() -> Path:
    """Get project root directory (where config/ folder lives)."""
    # This file is at src/infrastructure/config/config_loader.py
    # Project root is 4 levels up
    return Path(__file__).parent.parent.parent.parent


def resolve_path(path_value: Union[str, Path]) -> Path:
    """
    Resolve path, supporting both relative and absolute paths.

    Relative paths are resolved from project root.
    Absolute paths are used as-is.

    Args:
        path_value: Path string or Path object

    Returns:
        Resolved Path object
    """
    path = Path(path_value)

    if path.is_absolute():
        # Use absolute path as-is
        return path
    else:
        # Resolve relative path from project root
        return (get_project_root() / path).resolve()


class ExcelConfig(BaseModel):
    """Excel connection configuration."""

    platform: str = "windows"
    auto_connect: bool = True


class ConnectionConfig(BaseModel):
    """Connection workflow configuration."""

    auto_list_on_error: bool = True  # Show workbook list when connection fails
    auto_build_graph: str = "prompt"  # prompt | always | never
    build_graph_timeout: int = 300  # seconds before timeout warning
    prefer_most_recent: bool = False  # Auto-select most recently modified if ambiguous


class SelectionConfig(BaseModel):
    """Selection behaviour configuration."""

    auto_expand_context: bool = True
    expand_rows: int = 5
    expand_cols: int = 3


class SnapshotCollapseConfig(BaseModel):
    """Snapshot collapse configuration."""

    empty_rows: bool = True
    empty_columns: bool = True
    show_summary: bool = True


class SnapshotSamplingConfig(BaseModel):
    """Snapshot sampling configuration."""

    enabled: bool = True
    threshold_cells: int = 3000
    sample_every_n_rows: int = 10
    always_show_first_n: int = 10
    always_show_last_n: int = 5


class SnapshotConfig(BaseModel):
    """Snapshot generation configuration."""

    format: SnapshotFormat = SnapshotFormat.MARKDOWN
    max_cells_per_snapshot: int = 10000
    collapse: SnapshotCollapseConfig = Field(default_factory=SnapshotCollapseConfig)
    sampling: SnapshotSamplingConfig = Field(default_factory=SnapshotSamplingConfig)


class DependencyCacheConfig(BaseModel):
    """Dependency cache configuration."""

    enabled: bool = True
    location: Path = Path(".cache")
    auto_rebuild_on_change: bool = True

    @field_validator('location', mode='before')
    @classmethod
    def resolve_cache_location(cls, v):
        """Resolve path (supports relative and absolute)."""
        return resolve_path(v)


class DependenciesConfig(BaseModel):
    """Dependency analysis configuration."""

    mode: DependencyMode = DependencyMode.ON_DEMAND
    batch_size: int = 1000  # Rows per batch for full_graph mode
    default_depth: int = 3
    max_depth: int = 10
    default_direction: TraceDirection = TraceDirection.BOTH
    cross_sheet: bool = True
    cache: DependencyCacheConfig = Field(default_factory=DependencyCacheConfig)


class AnnotationsConfig(BaseModel):
    """Annotations configuration."""

    storage: str = "separate_file"
    file_location: Path = Path(".cache")
    default_sheet_scope: bool = True

    @field_validator('file_location', mode='before')
    @classmethod
    def resolve_annotation_location(cls, v):
        """Resolve path (supports relative and absolute)."""
        return resolve_path(v)


class ManualProviderConfig(BaseModel):
    """Manual LLM provider configuration."""

    enabled: bool = True
    input_file: Path = Path("llm_input.txt")
    output_file: Path = Path("llm_output.txt")

    @field_validator('input_file', 'output_file', mode='before')
    @classmethod
    def resolve_llm_file(cls, v):
        """Resolve path (supports relative and absolute)."""
        return resolve_path(v)


class InternalAPIProviderConfig(BaseModel):
    """Internal API provider configuration."""

    enabled: bool = False
    endpoint: str = ""
    auth_type: str = ""
    tenant_id: str = ""
    model: str = ""


class MockProviderConfig(BaseModel):
    """Mock provider configuration."""

    enabled: bool = True


class LLMProvidersConfig(BaseModel):
    """LLM providers configuration."""

    manual: ManualProviderConfig = Field(default_factory=ManualProviderConfig)
    internal_api: InternalAPIProviderConfig = Field(default_factory=InternalAPIProviderConfig)
    mock: MockProviderConfig = Field(default_factory=MockProviderConfig)


class LLMConfig(BaseModel):
    """LLM configuration."""

    default_provider: str = "manual"
    providers: LLMProvidersConfig = Field(default_factory=LLMProvidersConfig)


class AgentConfig(BaseModel):
    """Agent configuration."""

    max_iterations: int = 15
    verbose: bool = True
    tools: list[str] = Field(
        default_factory=lambda: [
            "get_workbook_overview",
            "get_snapshot",
            "search_cells",
            "get_cell_details",
            "get_formulas",
            "trace_dependencies",
            "get_current_selection",
            "get_annotations",
            "add_annotation",
            "search_annotations",
        ]
    )


class LoggingConfig(BaseModel):
    """Logging configuration.

    Note: Console output and rich tracebacks are always enabled.
    The {date} placeholder in file path is replaced with YYYY-MM-DD at runtime.
    """

    level: LogLevel = LogLevel.INFO
    file: Path = Path("logs/excel_sidekick_{date}.log")
    format: str = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"

    @field_validator('file', mode='before')
    @classmethod
    def resolve_log_file(cls, v):
        """Resolve path (supports relative and absolute)."""
        return resolve_path(v)


class CLIConfig(BaseModel):
    """CLI configuration.

    Note: Welcome message is always shown in REPL mode.
    """

    prompt: str = "excel-sidekick> "
    history_file: Path = Path("logs/cli_history.txt")

    @field_validator('history_file', mode='before')
    @classmethod
    def resolve_history_file(cls, v):
        """Resolve path (supports relative and absolute)."""
        return resolve_path(v)


class Config(BaseModel):
    """Main configuration class for Excel Sidekick."""

    excel: ExcelConfig = Field(default_factory=ExcelConfig)
    connection: ConnectionConfig = Field(default_factory=ConnectionConfig)
    selection: SelectionConfig = Field(default_factory=SelectionConfig)
    snapshot: SnapshotConfig = Field(default_factory=SnapshotConfig)
    dependencies: DependenciesConfig = Field(default_factory=DependenciesConfig)
    annotations: AnnotationsConfig = Field(default_factory=AnnotationsConfig)
    llm: LLMConfig = Field(default_factory=LLMConfig)
    agent: AgentConfig = Field(default_factory=AgentConfig)
    logging: LoggingConfig = Field(default_factory=LoggingConfig)
    cli: CLIConfig = Field(default_factory=CLIConfig)


def load_config(config_path: Optional[Path] = None) -> Config:
    """
    Load configuration from YAML file.

    Args:
        config_path: Path to config file. If None, uses default 'config/config.yaml'

    Returns:
        Loaded configuration object

    Raises:
        ConfigurationError: If config file is not found or invalid
    """
    if config_path is None:
        # Default to config/config.yaml relative to project root
        project_root = Path(__file__).parent.parent.parent.parent
        config_path = project_root / "config" / "config.yaml"

    if not config_path.exists():
        raise ConfigurationError(f"Configuration file not found: {config_path}")

    try:
        with open(config_path, "r") as f:
            config_dict = yaml.safe_load(f)

        if config_dict is None:
            config_dict = {}

        return Config(**config_dict)

    except yaml.YAMLError as e:
        raise ConfigurationError(f"Invalid YAML in config file: {e}")
    except Exception as e:
        raise ConfigurationError(f"Failed to load configuration: {e}")


def get_config(config_path: Optional[Path] = None) -> Config:
    """
    Get configuration instance (convenience function).

    Args:
        config_path: Path to config file

    Returns:
        Configuration object
    """
    return load_config(config_path)
