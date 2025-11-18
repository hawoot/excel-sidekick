"""Search command - Search annotations (placeholder for Phase 2)."""

from typing import Optional

from rich.console import Console

from src.application.excel_assistant_service import ExcelAssistantService
from src.presentation.cli.formatters import ResponseFormatter

from src.shared.logging import get_logger

logger = get_logger(__name__)

class SearchCommand:
    """Handle search command."""

    def __init__(
        self,
        service: ExcelAssistantService,
        console: Console,
        formatter: ResponseFormatter,
    ):
        """
        Initialize search command.

        Args:
            service: Application service
            console: Rich console
            formatter: Response formatter
        """
        self.service = service
        self.console = console
        self.formatter = formatter

    def execute(self, query: str, sheet: Optional[str] = None) -> bool:
        """
        Execute search command.

        Args:
            query: Search query
            sheet: Optional sheet to search in

        Returns:
            True if successful, False otherwise
        """
        # For Phase 1, just search annotations
        # Phase 2 will add cell content search
        try:
            # Check connection
            if not self.service.is_connected():
                self.formatter.format_error(
                    ValueError("Not connected to a workbook. Use 'connect' first.")
                )
                return False

            # Get all annotations for sheet
            annotations = self.service.get_annotations(sheet=sheet)

            # Filter by query (simple substring match)
            query_lower = query.lower()
            matches = [
                ann
                for ann in annotations
                if query_lower in ann.label.lower()
                or (ann.description and query_lower in ann.description.lower())
            ]

            if not matches:
                self.console.print(
                    f"[yellow]No annotations found matching '{query}'[/yellow]"
                )
                return True

            # Display results
            self.console.print(
                f"\n[bold]Found {len(matches)} annotation(s) matching '{query}':[/bold]\n"
            )

            for ann in matches:
                self.console.print(
                    f"  [cyan]{ann.range.to_address()}[/cyan] - "
                    f"[bold]{ann.label}[/bold]"
                )
                if ann.description:
                    self.console.print(f"    {ann.description}")

            return True

        except Exception as e:
            logger.exception("Command execution failed")
            self.formatter.format_error(e)
            return False
