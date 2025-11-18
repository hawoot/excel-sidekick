"""Explain command - Explain current selection."""

from typing import Optional

from rich.console import Console

from src.application.excel_assistant_service import ExcelAssistantService
from src.domain.models.selection import Selection
from src.presentation.cli.formatters import ResponseFormatter

from src.shared.logging import get_logger

logger = get_logger(__name__)

class ExplainCommand:
    """Handle explain command."""

    def __init__(
        self,
        service: ExcelAssistantService,
        console: Console,
        formatter: ResponseFormatter,
    ):
        """
        Initialize explain command.

        Args:
            service: Application service
            console: Rich console
            formatter: Response formatter
        """
        self.service = service
        self.console = console
        self.formatter = formatter

    def execute(
        self,
        selection: Optional[Selection] = None,
        mode: str = "educational",
        show_context: bool = False,
    ) -> bool:
        """
        Execute explain command.

        Args:
            selection: Selection to explain (None for current)
            mode: Query mode (educational, technical, concise)
            show_context: Whether to show context used

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

            # Get selection if not provided
            if selection is None:
                self.console.print("[dim]Using current Excel selection...[/dim]\n")

            # Explain selection
            response = self.service.explain_selection(
                selection=selection,
                mode=mode,
            )

            # Format and display response
            self.formatter.format_response(
                response=response,
                show_context=show_context,
            )

            return True

        except Exception as e:
            logger.exception("Command execution failed")
            self.formatter.format_error(e)
            return False
