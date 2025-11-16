"""REPL (Read-Eval-Print Loop) for interactive CLI."""

from typing import Optional

from prompt_toolkit import PromptSession
from prompt_toolkit.completion import WordCompleter
from prompt_toolkit.history import FileHistory
from rich.console import Console

from src.application.excel_assistant_service import ExcelAssistantService
from src.presentation.cli.commands import (
    AnnotateCommand,
    AskCommand,
    CacheCommand,
    ConnectCommand,
    ExplainCommand,
    SearchCommand,
    TraceCommand,
)
from src.presentation.cli.formatters import ResponseFormatter, TreeFormatter


class ExcelSidekickREPL:
    """Interactive REPL for Excel Sidekick."""

    def __init__(self, service: ExcelAssistantService):
        """
        Initialize REPL.

        Args:
            service: Application service
        """
        self.service = service
        self.console = Console()

        # Create formatters
        self.formatter = ResponseFormatter(self.console)
        self.tree_formatter = TreeFormatter(self.console)

        # Create command handlers
        self.connect_cmd = ConnectCommand(service, self.console, self.formatter)
        self.ask_cmd = AskCommand(service, self.console, self.formatter)
        self.explain_cmd = ExplainCommand(service, self.console, self.formatter)
        self.trace_cmd = TraceCommand(
            service, self.console, self.formatter, self.tree_formatter
        )
        self.annotate_cmd = AnnotateCommand(service, self.console, self.formatter)
        self.cache_cmd = CacheCommand(service, self.console, self.formatter)
        self.search_cmd = SearchCommand(service, self.console, self.formatter)

        # Set up prompt with auto-completion
        commands = [
            "connect",
            "ask",
            "explain",
            "trace",
            "annotate",
            "search",
            "cache",
            "status",
            "help",
            "exit",
            "quit",
        ]
        self.completer = WordCompleter(commands, ignore_case=True)

        # Create session with history
        self.session: Optional[PromptSession] = None

    def run(self) -> None:
        """Run the REPL."""
        # Print banner
        self._print_banner()

        # Create session with history file
        self.session = PromptSession(
            history=FileHistory(".excel_sidekick_history"),
            completer=self.completer,
        )

        # Main loop
        while True:
            try:
                # Get prompt symbol
                if self.service.is_connected():
                    wb = self.service.get_current_workbook()
                    prompt = f"excel-sidekick ({wb.name if wb else 'connected'})> "
                else:
                    prompt = "excel-sidekick> "

                # Get user input
                user_input = self.session.prompt(prompt)

                # Skip empty input
                if not user_input.strip():
                    continue

                # Parse and execute command
                if not self._execute_command(user_input.strip()):
                    break

            except KeyboardInterrupt:
                self.console.print("\n[yellow]Use 'exit' or 'quit' to exit[/yellow]")
                continue

            except EOFError:
                break

        # Print goodbye message
        self.console.print("\n[dim]Goodbye![/dim]")

    def _print_banner(self) -> None:
        """Print welcome banner."""
        self.console.print(
            "\n[bold cyan]Excel Sidekick[/bold cyan] - AI-powered Excel exploration\n"
        )
        self.console.print("[dim]Type 'help' for available commands[/dim]\n")

    def _execute_command(self, command_line: str) -> bool:
        """
        Execute a command.

        Args:
            command_line: Command line to execute

        Returns:
            True to continue REPL, False to exit
        """
        # Split command and arguments
        parts = command_line.split(maxsplit=1)
        command = parts[0].lower()
        args = parts[1] if len(parts) > 1 else ""

        # Handle commands
        if command in ("exit", "quit"):
            return False

        elif command == "help":
            self._show_help()

        elif command == "connect":
            workbook_name = args if args else None
            self.connect_cmd.execute(workbook_name)

        elif command == "ask":
            if not args:
                self.formatter.format_error(ValueError("Usage: ask <question>"))
            else:
                self.ask_cmd.execute(question=args)

        elif command == "explain":
            self.explain_cmd.execute()

        elif command == "trace":
            if not args:
                self.formatter.format_error(
                    ValueError("Usage: trace <cell_address> [direction] [depth]")
                )
            else:
                # Parse trace arguments
                trace_parts = args.split()
                cell_address = trace_parts[0]
                direction = trace_parts[1] if len(trace_parts) > 1 else "both"
                depth = int(trace_parts[2]) if len(trace_parts) > 2 else 5

                self.trace_cmd.execute(
                    cell_address=cell_address,
                    direction=direction,
                    depth=depth,
                )

        elif command == "annotate":
            if not args:
                # List all annotations
                self.annotate_cmd.list()
            else:
                # Parse annotate arguments
                # Format: annotate <range> <label> [description]
                annotate_parts = args.split(maxsplit=2)
                if len(annotate_parts) < 2:
                    self.formatter.format_error(
                        ValueError("Usage: annotate <range> <label> [description]")
                    )
                else:
                    range_address = annotate_parts[0]
                    label = annotate_parts[1]
                    description = annotate_parts[2] if len(annotate_parts) > 2 else None

                    self.annotate_cmd.add(
                        range_address=range_address,
                        label=label,
                        description=description,
                    )

        elif command == "search":
            if not args:
                self.formatter.format_error(ValueError("Usage: search <query>"))
            else:
                self.search_cmd.execute(query=args)

        elif command == "cache":
            # Parse cache subcommand
            if not args:
                self.cache_cmd.status()
            elif args == "rebuild":
                self.cache_cmd.rebuild()
            elif args == "clear":
                self.cache_cmd.clear()
            elif args == "status":
                self.cache_cmd.status()
            else:
                self.formatter.format_error(
                    ValueError("Usage: cache [status|rebuild|clear]")
                )

        elif command == "status":
            self.cache_cmd.status()

        else:
            self.formatter.format_error(
                ValueError(f"Unknown command: {command}. Type 'help' for available commands.")
            )

        return True

    def _show_help(self) -> None:
        """Show help message."""
        help_text = """
[bold]Available Commands:[/bold]

[cyan]connect [workbook_name][/cyan]
    Connect to Excel workbook (active if no name specified)

[cyan]ask <question>[/cyan]
    Ask a question about the workbook

[cyan]explain[/cyan]
    Explain current Excel selection

[cyan]trace <cell_address> [direction] [depth][/cyan]
    Trace cell dependencies
    - direction: precedents, dependents, both (default: both)
    - depth: maximum depth (default: 5)

[cyan]annotate [range] [label] [description][/cyan]
    Add annotation or list all annotations
    - No arguments: list all annotations
    - With arguments: add new annotation

[cyan]search <query>[/cyan]
    Search annotations

[cyan]cache [status|rebuild|clear][/cyan]
    Manage dependency graph cache
    - status: show cache status (default)
    - rebuild: rebuild cache
    - clear: clear cache

[cyan]status[/cyan]
    Show cache and connection status

[cyan]help[/cyan]
    Show this help message

[cyan]exit, quit[/cyan]
    Exit Excel Sidekick

[bold]Examples:[/bold]

  connect MyWorkbook.xlsx
  ask What does this sheet calculate?
  explain
  trace Sheet1!A1 both 3
  annotate Sheet1!A1:B10 "Revenue Inputs" "Monthly revenue by product"
  search revenue
  cache rebuild
"""
        self.console.print(help_text)
