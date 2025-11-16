"""Annotate command - Manage annotations."""

from typing import Optional

from rich.console import Console
from rich.table import Table

from src.application.excel_assistant_service import ExcelAssistantService
from src.presentation.cli.formatters import ResponseFormatter


class AnnotateCommand:
    """Handle annotate command."""

    def __init__(
        self,
        service: ExcelAssistantService,
        console: Console,
        formatter: ResponseFormatter,
    ):
        """
        Initialize annotate command.

        Args:
            service: Application service
            console: Rich console
            formatter: Response formatter
        """
        self.service = service
        self.console = console
        self.formatter = formatter

    def add(
        self,
        range_address: str,
        label: str,
        description: Optional[str] = None,
    ) -> bool:
        """
        Add annotation to range.

        Args:
            range_address: Range address (e.g., "Sheet1!A1:B10")
            label: Short label
            description: Optional detailed description

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

            # Add annotation
            self.service.add_annotation(
                range_address=range_address,
                label=label,
                description=description,
            )

            self.formatter.format_success(
                f"Added annotation '{label}' for {range_address}"
            )

            return True

        except Exception as e:
            self.formatter.format_error(e)
            return False

    def list(self, sheet: Optional[str] = None) -> bool:
        """
        List annotations for sheet or all sheets.

        Args:
            sheet: Sheet name (None for all)

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

            # Get annotations
            annotations = self.service.get_annotations(sheet=sheet)

            if not annotations:
                self.console.print(
                    f"[yellow]No annotations found{' for ' + sheet if sheet else ''}[/yellow]"
                )
                return True

            # Display as table
            table = Table(title=f"Annotations{' for ' + sheet if sheet else ''}")
            table.add_column("Sheet", style="cyan")
            table.add_column("Range", style="green")
            table.add_column("Label", style="bold")
            table.add_column("Description", style="dim")

            for annotation in annotations:
                table.add_row(
                    annotation.sheet,
                    annotation.range.to_address(include_sheet=False),
                    annotation.label,
                    annotation.description or "",
                )

            self.console.print(table)

            return True

        except Exception as e:
            self.formatter.format_error(e)
            return False
