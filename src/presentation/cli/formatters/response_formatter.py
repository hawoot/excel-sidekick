"""Formatter for assistant responses."""

from rich.console import Console
from rich.markdown import Markdown
from rich.panel import Panel

from src.domain.models.query import AssistantResponse


class ResponseFormatter:
    """Format assistant responses for CLI output."""

    def __init__(self, console: Console):
        """
        Initialize response formatter.

        Args:
            console: Rich console for output
        """
        self.console = console

    def format_response(
        self,
        response: AssistantResponse,
        show_context: bool = False,
        show_metadata: bool = False,
    ) -> None:
        """
        Format and print assistant response.

        Args:
            response: Response to format
            show_context: Whether to show context used
            show_metadata: Whether to show metadata
        """
        # Print question
        self.console.print(f"\n[bold]Question:[/bold] {response.question}\n")

        # Print answer as markdown
        answer_md = Markdown(response.answer)
        self.console.print(Panel(answer_md, title="Answer", border_style="green"))

        # Show context if requested
        if show_context and response.context_used:
            self.console.print("\n[bold]Context Used:[/bold]")
            for key, value in response.context_used.items():
                if isinstance(value, list):
                    self.console.print(f"  {key}: {len(value)} items")
                else:
                    self.console.print(f"  {key}")

        # Show dependencies if traced
        if response.dependencies_traced:
            self.console.print(
                f"\n[dim]Dependencies traced: {response.dependencies_traced.total_nodes} nodes[/dim]"
            )

        # Show annotations if found
        if response.annotations_found:
            self.console.print(
                f"[dim]Annotations found: {len(response.annotations_found)}[/dim]"
            )

        # Show metadata if requested
        if show_metadata and response.metadata:
            self.console.print("\n[bold]Metadata:[/bold]")
            for key, value in response.metadata.items():
                self.console.print(f"  {key}: {value}")

    def format_error(self, error: Exception) -> None:
        """
        Format and print error message.

        Args:
            error: Exception to format
        """
        self.console.print(
            Panel(
                f"[red]{type(error).__name__}:[/red] {str(error)}",
                title="Error",
                border_style="red",
            )
        )

    def format_success(self, message: str) -> None:
        """
        Format and print success message.

        Args:
            message: Success message
        """
        self.console.print(f"[green]✓[/green] {message}")

    def format_info(self, message: str) -> None:
        """
        Format and print info message.

        Args:
            message: Info message
        """
        self.console.print(f"[blue]ℹ[/blue] {message}")

    def format_warning(self, message: str) -> None:
        """
        Format and print warning message.

        Args:
            message: Warning message
        """
        self.console.print(f"[yellow]⚠[/yellow] {message}")
