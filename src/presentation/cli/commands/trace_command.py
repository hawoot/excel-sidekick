"""Trace command - Trace cell dependencies."""

from rich.console import Console

from src.application.excel_assistant_service import ExcelAssistantService
from src.presentation.cli.formatters import ResponseFormatter, TreeFormatter
from src.shared.types import TraceDirection

from src.shared.logging import get_logger

logger = get_logger(__name__)

class TraceCommand:
    """Handle trace command."""

    def __init__(
        self,
        service: ExcelAssistantService,
        console: Console,
        formatter: ResponseFormatter,
        tree_formatter: TreeFormatter,
    ):
        """
        Initialize trace command.

        Args:
            service: Application service
            console: Rich console
            formatter: Response formatter
            tree_formatter: Tree formatter
        """
        self.service = service
        self.console = console
        self.formatter = formatter
        self.tree_formatter = tree_formatter

    def execute(
        self,
        cell_address: str,
        direction: str = "both",
        depth: int = 5,
    ) -> bool:
        """
        Execute trace command.

        Args:
            cell_address: Cell address to trace from
            direction: Direction to trace (precedents, dependents, both)
            depth: Maximum depth to trace

        Returns:
            True if successful, False otherwise
        """
        try:
            # Check connection
            if not self.service.is_connected():
                self.formatter.format_error(
                    ValueError("Not connected to a workbook. Use 'connect' first.")
                )
                return False

            # Parse direction
            direction_map = {
                "precedents": TraceDirection.PRECEDENTS,
                "dependents": TraceDirection.DEPENDENTS,
                "both": TraceDirection.BOTH,
            }

            if direction.lower() not in direction_map:
                self.formatter.format_error(
                    ValueError(f"Invalid direction: {direction}. Use precedents, dependents, or both.")
                )
                return False

            trace_direction = direction_map[direction.lower()]

            # Trace dependencies
            self.console.print(
                f"\n[dim]Tracing {direction} for {cell_address} (depth={depth})...[/dim]\n"
            )

            dep_tree = self.service.dependency_analysis.trace_dependencies(
                cell_address=cell_address,
                direction=trace_direction,
                depth=depth,
            )

            # Format and display tree
            self.tree_formatter.format_tree(dep_tree, max_depth=depth)

            return True

        except Exception as e:
            logger.exception("Command execution failed")
            self.formatter.format_error(e)
            return False
