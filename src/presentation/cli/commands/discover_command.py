"""Discover command - Discover all open Excel workbooks."""

from rich.console import Console
from rich.table import Table

from src.infrastructure.excel.workbook_discovery import WorkbookDiscovery
from src.presentation.cli.formatters import ResponseFormatter

from src.shared.logging import get_logger

logger = get_logger(__name__)

class DiscoverCommand:
    """Handle discover command to show all open workbooks."""

    def __init__(self, console: Console, formatter: ResponseFormatter):
        """
        Initialize discover command.

        Args:
            console: Rich console
            formatter: Response formatter
        """
        self.console = console
        self.formatter = formatter

    def execute(self) -> bool:
        """
        Execute discover command to show all open workbooks.

        Returns:
            True if successful, False otherwise
        """
        try:
            self.console.print("\n[dim]Discovering open workbooks...[/dim]\n")

            # Discover all workbooks
            workbooks = WorkbookDiscovery.list_all_workbooks()

            if not workbooks:
                self.console.print("[yellow]No open workbooks found[/yellow]")
                self.formatter.format_info("Open an Excel workbook and try again")
                return True

            # Display workbooks in table
            table = Table(title=f"Open Workbooks ({len(workbooks)} found)")

            table.add_column("#", style="dim", justify="right", width=3)
            table.add_column("PID", style="cyan", width=7)
            table.add_column("Workbook", style="bold")
            table.add_column("Full Path", style="dim", overflow="fold")
            table.add_column("Sheets", justify="right", width=7)

            for idx, wb_info in enumerate(workbooks, start=1):
                # Add unsaved indicator
                workbook_name = wb_info.workbook_name
                if not wb_info.is_saved:
                    workbook_name += " [red]*[/red]"

                table.add_row(
                    str(idx),
                    str(wb_info.excel_pid),
                    workbook_name,
                    wb_info.full_path,
                    str(wb_info.sheet_count),
                )

            self.console.print(table)

            # Check for duplicates and show warning
            duplicate_paths = WorkbookDiscovery.get_duplicate_paths(workbooks)
            if duplicate_paths:
                self.console.print(
                    f"\n[yellow]⚠ Note:[/yellow] The following workbook(s) are open in multiple Excel instances:"
                )
                for path in duplicate_paths:
                    matching_pids = [
                        str(wb.excel_pid) for wb in workbooks if wb.full_path == path
                    ]
                    # Get filename from first match
                    filename = next(
                        (wb.workbook_name for wb in workbooks if wb.full_path == path),
                        "Unknown",
                    )
                    self.console.print(
                        f"  • {filename} (PIDs: {', '.join(matching_pids)})"
                    )

            self.console.print(
                f"\n[dim]Use 'connect' to connect to a workbook[/dim]"
            )
            self.console.print(
                f"[dim]Use 'connect <full_path>' to connect to a specific file[/dim]"
            )

            return True

        except Exception as e:
            logger.exception("Command execution failed")
            self.formatter.format_error(e)
            return False
