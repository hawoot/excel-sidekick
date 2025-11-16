"""Ask command - Ask a question about the workbook."""

from typing import Optional

from rich.console import Console

from src.application.excel_assistant_service import ExcelAssistantService
from src.domain.models.selection import Selection
from src.presentation.cli.formatters import ResponseFormatter


class AskCommand:
    """Handle ask command."""

    def __init__(
        self,
        service: ExcelAssistantService,
        console: Console,
        formatter: ResponseFormatter,
    ):
        """
        Initialize ask command.

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
        question: str,
        selection: Optional[Selection] = None,
        mode: str = "educational",
        show_context: bool = False,
    ) -> bool:
        """
        Execute ask command.

        Args:
            question: Question to ask
            selection: Optional selection to use as context
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

            # Ask question
            self.console.print(f"\n[dim]Exploring workbook to answer: {question}[/dim]\n")

            response = self.service.ask_question(
                question=question,
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
            self.formatter.format_error(e)
            return False
