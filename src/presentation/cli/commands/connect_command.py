"""Connect command - Connect to Excel workbook."""

from pathlib import Path
from typing import Optional

from rich.console import Console

from src.application.excel_assistant_service import ExcelAssistantService
from src.infrastructure.excel.workbook_discovery import WorkbookDiscovery, WorkbookInfo
from src.presentation.cli.formatters import ResponseFormatter
from src.presentation.cli.interactive_selector import InteractiveWorkbookSelector
from src.shared.exceptions import ExcelConnectionError


class ConnectCommand:
    """Handle connect command with interactive workbook selection."""

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
        self.selector = InteractiveWorkbookSelector(console)

    def execute(self, full_path: Optional[str] = None) -> bool:
        """
        Execute connect command with interactive selection.

        Args:
            full_path: Optional full path to workbook. If provided, tries to connect directly.
                      If None, shows interactive list.

        Returns:
            True if successful, False otherwise
        """
        try:
            # Disconnect if already connected
            if self.service.is_connected():
                self.console.print("[yellow]Disconnecting from current workbook...[/yellow]")
                self.service.disconnect()

            # Determine which workbook to connect to
            workbook_info = self._select_workbook(full_path)

            if workbook_info is None:
                self.console.print("[yellow]Connection cancelled[/yellow]")
                return False

            # Connect to selected workbook (without building graph yet)
            self.console.print(
                f"\nConnecting to [bold]{workbook_info.workbook_name}[/bold] "
                f"(Excel PID: {workbook_info.excel_pid})..."
            )

            workbook = self.service.connect_to_workbook_info(
                workbook_info, build_graph=False
            )

            # Show connection success
            self.formatter.format_success(
                f"Connected to '{workbook.name}' "
                f"({len(workbook.sheets)} sheets, "
                f"~{workbook.total_formula_count()} formulas)"
            )

            # Show sheet list
            self.console.print("\n[bold]Sheets:[/bold]")
            for sheet in workbook.sheets:
                self.console.print(
                    f"  • {sheet.name} "
                    f"({sheet.formula_count} formulas, "
                    f"{sheet.cell_count} cells)"
                )

            # Ask user if they want to build the dependency graph
            should_build = self._prompt_build_graph(workbook_info, workbook.total_formula_count())

            if should_build:
                self.console.print("\n[dim]Building dependency graph...[/dim]")
                self.service.build_graph()
                self.formatter.format_success("Dependency graph built and cached")

            self.console.print("\n[green]Ready to explore![/green]")

            return True

        except ExcelConnectionError as e:
            self.formatter.format_error(e)

            # Show list of available workbooks on error if configured
            if self.service.config.connection.auto_list_on_error:
                self.console.print("\n[dim]Showing available workbooks...[/dim]\n")
                try:
                    workbooks = WorkbookDiscovery.list_all_workbooks()
                    if workbooks:
                        # Try again with interactive selection
                        return self.execute(full_path=None)
                except Exception:
                    pass

            self.formatter.format_info(
                "Use 'list' command to see available workbooks"
            )
            return False

        except Exception as e:
            self.formatter.format_error(e)
            return False

    def _select_workbook(self, full_path: Optional[str]) -> Optional[WorkbookInfo]:
        """
        Select workbook either by path or interactively.

        Args:
            full_path: Optional full path to workbook

        Returns:
            Selected WorkbookInfo or None if cancelled/not found
        """
        if full_path:
            # User provided a path - try to find it
            return self._select_by_path(full_path)
        else:
            # No path provided - show interactive list
            return self.selector.select_workbook()

    def _select_by_path(self, full_path: str) -> Optional[WorkbookInfo]:
        """
        Select workbook by full path.

        Handles case where same file is open in multiple Excel instances.

        Args:
            full_path: Full path to workbook

        Returns:
            Selected WorkbookInfo or None
        """
        try:
            # Find workbooks matching this path
            matches = WorkbookDiscovery.find_by_path(full_path)

            if not matches:
                # No matches found
                self.console.print(
                    f"[red]Error:[/red] Workbook not found: {full_path}"
                )
                self.console.print(
                    "[dim]Make sure the file is open in Excel[/dim]"
                )
                return None

            elif len(matches) == 1:
                # Single match - use it
                return matches[0]

            else:
                # Multiple matches (same file in different Excel instances)
                self.console.print(
                    f"\n[yellow]Found {len(matches)} instances of this workbook:[/yellow]\n"
                )

                # Display instances
                self.selector._display_workbook_table(matches)

                # Let user select which instance
                selected = self.selector.select_workbook(matches)
                return selected

        except Exception as e:
            self.console.print(f"[red]Error finding workbook:[/red] {e}")
            return None

    def _prompt_build_graph(self, workbook_info: WorkbookInfo, formula_count: int) -> bool:
        """
        Ask user if they want to build dependency graph.

        Respects config setting for auto_build_graph.

        Args:
            workbook_info: Workbook being connected to
            formula_count: Number of formulas in workbook

        Returns:
            True if user wants to build graph
        """
        # Check config setting
        auto_build = self.service.config.connection.auto_build_graph

        if auto_build == "always":
            return True
        elif auto_build == "never":
            return False
        else:  # "prompt"
            return self.selector.prompt_build_graph(workbook_info, formula_count)
