"""Cache command - Manage dependency graph cache."""

from rich.console import Console
from rich.table import Table

from src.application.excel_assistant_service import ExcelAssistantService
from src.presentation.cli.formatters import ResponseFormatter


class CacheCommand:
    """Handle cache command."""

    def __init__(
        self,
        service: ExcelAssistantService,
        console: Console,
        formatter: ResponseFormatter,
    ):
        """
        Initialize cache command.

        Args:
            service: Application service
            console: Rich console
            formatter: Response formatter
        """
        self.service = service
        self.console = console
        self.formatter = formatter

    def status(self) -> bool:
        """
        Show cache status.

        Returns:
            True if successful, False otherwise
        """
        try:
            status = self.service.get_cache_status()

            if not status["connected"]:
                self.console.print(
                    "[yellow]Not connected to a workbook[/yellow]"
                )
                return True

            # Display as table
            table = Table(title="Cache Status")
            table.add_column("Property", style="cyan")
            table.add_column("Value", style="green")

            table.add_row("Workbook", status["workbook"])
            table.add_row(
                "Graph Cached",
                " Yes" if status["graph_cached"] else " No",
            )
            table.add_row("Node Count", str(status["node_count"]))
            table.add_row("Formula Count", str(status["formula_count"]))
            table.add_row(
                "Has Annotations",
                " Yes" if status["has_annotations"] else " No",
            )

            self.console.print(table)

            return True

        except Exception as e:
            self.formatter.format_error(e)
            return False

    def rebuild(self) -> bool:
        """
        Rebuild dependency graph cache.

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

            self.console.print("[dim]Rebuilding dependency graph...[/dim]")

            self.service.rebuild_cache()

            self.formatter.format_success("Cache rebuilt")

            return True

        except Exception as e:
            self.formatter.format_error(e)
            return False

    def clear(self) -> bool:
        """
        Clear dependency graph cache.

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

            self.service.clear_cache()

            self.formatter.format_success("Cache cleared")

            return True

        except Exception as e:
            self.formatter.format_error(e)
            return False
