"""Interactive workbook selector for CLI."""

from typing import List, Optional

from prompt_toolkit import prompt
from prompt_toolkit.validation import ValidationError, Validator
from rich.console import Console
from rich.table import Table

from src.infrastructure.excel.workbook_discovery import WorkbookDiscovery, WorkbookInfo


class NumberValidator(Validator):
    """Validate that input is a valid selection number."""

    def __init__(self, max_value: int):
        """
        Initialize validator.

        Args:
            max_value: Maximum valid number
        """
        self.max_value = max_value

    def validate(self, document):
        """Validate the input."""
        text = document.text.strip().lower()

        # Allow 'q' to quit
        if text == "q":
            return

        # Check if it's a valid number
        try:
            num = int(text)
            if num < 1 or num > self.max_value:
                raise ValidationError(
                    message=f"Please enter a number between 1 and {self.max_value}, or 'q' to cancel"
                )
        except ValueError:
            raise ValidationError(
                message=f"Please enter a number between 1 and {self.max_value}, or 'q' to cancel"
            )


class InteractiveWorkbookSelector:
    """Interactive workbook selection interface."""

    def __init__(self, console: Console):
        """
        Initialize selector.

        Args:
            console: Rich console for output
        """
        self.console = console

    def select_workbook(
        self, workbooks: Optional[List[WorkbookInfo]] = None
    ) -> Optional[WorkbookInfo]:
        """
        Display workbook list and let user select one.

        Args:
            workbooks: List of workbooks to choose from. If None, discovers all workbooks.

        Returns:
            Selected WorkbookInfo, or None if cancelled
        """
        # Discover workbooks if not provided
        if workbooks is None:
            self.console.print("\n[dim]Discovering open workbooks across all Excel instances...[/dim]\n")
            try:
                workbooks = WorkbookDiscovery.list_all_workbooks()
            except Exception as e:
                self.console.print(f"[red]Error:[/red] {e}")
                return None

        if not workbooks:
            self.console.print("[yellow]No open workbooks found.[/yellow]")
            return None

        # Display workbook table
        self._display_workbook_table(workbooks)

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
                filename = workbooks[0].workbook_name if workbooks else "Unknown"
                self.console.print(
                    f"  • {filename} (PIDs: {', '.join(matching_pids)})"
                )
            self.console.print()

        # Get user selection
        validator = NumberValidator(len(workbooks))
        selection_text = prompt(
            f"Select workbook (1-{len(workbooks)}), or 'q' to cancel: ",
            validator=validator,
        )

        selection_text = selection_text.strip().lower()

        if selection_text == "q":
            return None

        # Get selected workbook (subtract 1 for 0-based indexing)
        selection_index = int(selection_text) - 1
        return workbooks[selection_index]

    def _display_workbook_table(self, workbooks: List[WorkbookInfo]) -> None:
        """
        Display workbooks in a formatted table.

        Args:
            workbooks: List of workbooks to display
        """
        table = Table(title="Open Workbooks")

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

    def confirm_selection(
        self, workbook_info: WorkbookInfo, message: Optional[str] = None
    ) -> bool:
        """
        Ask user to confirm workbook selection.

        Args:
            workbook_info: Workbook to confirm
            message: Optional custom message

        Returns:
            True if confirmed, False otherwise
        """
        if message is None:
            message = f"Connect to {workbook_info.workbook_name} (PID: {workbook_info.excel_pid})?"

        self.console.print(f"\n{message}")
        self.console.print(f"[dim]Path: {workbook_info.full_path}[/dim]")

        response = prompt("Confirm? (Y/n): ").strip().lower()

        return response in ("", "y", "yes")

    def prompt_build_graph(self, workbook_info: WorkbookInfo, formula_count: int) -> bool:
        """
        Ask user if they want to build the dependency graph.

        Args:
            workbook_info: Workbook to build graph for
            formula_count: Estimated number of formulas

        Returns:
            True if user wants to build graph
        """
        self.console.print(
            f"\n[bold]Build dependency graph?[/bold] "
            f"This will analyze ~{formula_count} formulas."
        )

        if formula_count > 1000:
            estimated_time = formula_count // 20  # Rough estimate: 20 formulas/second
            self.console.print(
                f"[dim]Estimated time: ~{estimated_time} seconds[/dim]"
            )

        response = prompt("Build graph? (Y/n): ").strip().lower()

        return response in ("", "y", "yes")
