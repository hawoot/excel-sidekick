"""Build command - Build dependency graph for connected workbook."""

from rich.console import Console

from src.application.excel_assistant_service import ExcelAssistantService
from src.presentation.cli.formatters import ResponseFormatter

from src.shared.logging import get_logger

logger = get_logger(__name__)

class BuildCommand:
    """Handle build command to build/rebuild dependency graph."""

    def __init__(
        self,
        service: ExcelAssistantService,
        console: Console,
        formatter: ResponseFormatter,
    ):
        """
        Initialize build command.

        Args:
            service: Application service
            console: Rich console
            formatter: Response formatter
        """
        self.service = service
        self.console = console
        self.formatter = formatter

    def execute(self, force: bool = False) -> bool:
        """
        Execute build command.

        Args:
            force: Whether to force rebuild even if graph exists

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

            workbook = self.service.get_current_workbook()
            formula_count = workbook.total_formula_count()

            # Check if graph already exists
            status = self.service.get_cache_status()

            if status.get("graph_cached") and not force:
                self.console.print(
                    "[yellow]Dependency graph already exists[/yellow]"
                )
                self.console.print(
                    f"[dim]Use 'build --force' to rebuild ({formula_count} formulas)[/dim]"
                )
                return True

            # Build the graph
            self.console.print(
                f"\nBuilding dependency graph for [bold]{workbook.name}[/bold]..."
            )

            if formula_count > 1000:
                estimated_time = formula_count // 20  # Rough estimate
                self.console.print(
                    f"[dim]Analyzing ~{formula_count} formulas (~{estimated_time} seconds)...[/dim]\n"
                )

            self.service.build_graph()

            # Show results
            graph = self.service.dependency_analysis.get_current_graph()
            self.formatter.format_success(
                f"Dependency graph built: {graph.node_count()} nodes, "
                f"{graph.edge_count()} edges"
            )

            self.console.print("[dim]Cache saved for future sessions[/dim]")

            return True

        except Exception as e:
            logger.exception("Command execution failed")
            self.formatter.format_error(e)
            return False
