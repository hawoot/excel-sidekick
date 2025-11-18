"""CLI command handlers."""

from src.presentation.cli.commands.annotate_command import AnnotateCommand
from src.presentation.cli.commands.ask_command import AskCommand
from src.presentation.cli.commands.build_command import BuildCommand
from src.presentation.cli.commands.cache_command import CacheCommand
from src.presentation.cli.commands.connect_command import ConnectCommand
from src.presentation.cli.commands.discover_command import DiscoverCommand
from src.presentation.cli.commands.explain_command import ExplainCommand
from src.presentation.cli.commands.list_command import ListCommand  # Deprecated, use DiscoverCommand
from src.presentation.cli.commands.search_command import SearchCommand
from src.presentation.cli.commands.trace_command import TraceCommand

__all__ = [
    "ConnectCommand",
    "DiscoverCommand",
    "ListCommand",  # Deprecated
    "BuildCommand",
    "AskCommand",
    "ExplainCommand",
    "TraceCommand",
    "AnnotateCommand",
    "CacheCommand",
    "SearchCommand",
]
