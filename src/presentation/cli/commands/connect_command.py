"""Connect command - Connect to Excel workbook."""

from typing import Optional

from rich.console import Console

from src.application.excel_assistant_service import ExcelAssistantService
from src.presentation.cli.formatters import ResponseFormatter
from src.shared.exceptions import ExcelConnectionError


class ConnectCommand:
    """Handle connect command."""

    def __init__(
        self,
        service: ExcelAssistantService,
        console: Console,
        formatter: ResponseFormatter,
    ):
        """
        Initialize connect command.

        Args:
            service: Application service
            console: Rich console
            formatter: Response formatter
        """
        self.service = service
        self.console = console
        self.formatter = formatter

    def execute(self, workbook_name: Optional[str] = None) -> bool:
        """
        Execute connect command.

        Args:
            workbook_name: Name of workbook (None for active)

        Returns:
            True if successful, False otherwise
        """
        try:
            # Disconnect if already connected
            if self.service.is_connected():
                self.console.print("[yellow]Disconnecting from current workbook...[/yellow]")
                self.service.disconnect()

            # Connect to workbook
            self.console.print(
                f"Connecting to workbook: {workbook_name or 'active'}..."
            )

            workbook = self.service.connect(workbook_name)

            # Show success message
            self.formatter.format_success(
                f"Connected to '{workbook.name}' "
                f"({len(workbook.sheets)} sheets, "
                f"{workbook.total_formula_count()} formulas)"
            )

            # Show sheet list
            self.console.print("\n[bold]Sheets:[/bold]")
            for sheet in workbook.sheets:
                self.console.print(
                    f"  " {sheet.name} "
                    f"({sheet.formula_count} formulas, "
                    f"{sheet.cell_count} cells)"
                )

            return True

        except ExcelConnectionError as e:
            self.formatter.format_error(e)
            self.formatter.format_info(
                "Make sure Excel is running with a workbook open"
            )
            return False

        except Exception as e:
            self.formatter.format_error(e)
            return False
