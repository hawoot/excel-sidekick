"""CLI application entry point."""

import sys
from pathlib import Path

import click
from rich.console import Console

from src.application.excel_assistant_service import ExcelAssistantService
from src.infrastructure.config.config_loader import load_config
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
from src.presentation.cli.repl import ExcelSidekickREPL


@click.group(invoke_without_command=True)
@click.option(
    "--config",
    "-c",
    type=click.Path(exists=True),
    default="config/config.yaml",
    help="Path to config file",
)
@click.pass_context
def cli(ctx: click.Context, config: str) -> None:
    """
    Excel Sidekick - AI-powered Excel exploration.

    If no command is provided, starts interactive REPL mode.
    """
    # Load config
    config_path = Path(config)
    if not config_path.exists():
        console = Console()
        console.print(f"[red]Error:[/red] Config file not found: {config}")
        sys.exit(1)

    cfg = load_config(config_path)

    # Create service
    service = ExcelAssistantService(cfg)

    # Store in context for subcommands
    ctx.ensure_object(dict)
    ctx.obj["service"] = service
    ctx.obj["config"] = cfg

    # If no subcommand, run REPL
    if ctx.invoked_subcommand is None:
        repl = ExcelSidekickREPL(service)
        repl.run()


@cli.command()
@click.argument("workbook_name", required=False)
@click.pass_context
def connect(ctx: click.Context, workbook_name: str) -> None:
    """Connect to Excel workbook."""
    service = ctx.obj["service"]
    console = Console()
    formatter = ResponseFormatter(console)

    cmd = ConnectCommand(service, console, formatter)
    success = cmd.execute(workbook_name)

    sys.exit(0 if success else 1)


@cli.command()
@click.argument("question")
@click.option("--mode", "-m", default="educational", help="Query mode")
@click.option("--show-context", "-c", is_flag=True, help="Show context used")
@click.pass_context
def ask(ctx: click.Context, question: str, mode: str, show_context: bool) -> None:
    """Ask a question about the workbook."""
    service = ctx.obj["service"]
    console = Console()
    formatter = ResponseFormatter(console)

    cmd = AskCommand(service, console, formatter)
    success = cmd.execute(question=question, mode=mode, show_context=show_context)

    sys.exit(0 if success else 1)


@cli.command()
@click.option("--mode", "-m", default="educational", help="Query mode")
@click.option("--show-context", "-c", is_flag=True, help="Show context used")
@click.pass_context
def explain(ctx: click.Context, mode: str, show_context: bool) -> None:
    """Explain current Excel selection."""
    service = ctx.obj["service"]
    console = Console()
    formatter = ResponseFormatter(console)

    cmd = ExplainCommand(service, console, formatter)
    success = cmd.execute(mode=mode, show_context=show_context)

    sys.exit(0 if success else 1)


@cli.command()
@click.argument("cell_address")
@click.option("--direction", "-d", default="both", help="Direction (precedents, dependents, both)")
@click.option("--depth", "-n", default=5, help="Maximum depth")
@click.pass_context
def trace(ctx: click.Context, cell_address: str, direction: str, depth: int) -> None:
    """Trace cell dependencies."""
    service = ctx.obj["service"]
    console = Console()
    formatter = ResponseFormatter(console)
    tree_formatter = TreeFormatter(console)

    cmd = TraceCommand(service, console, formatter, tree_formatter)
    success = cmd.execute(cell_address=cell_address, direction=direction, depth=depth)

    sys.exit(0 if success else 1)


@cli.command()
@click.argument("range_address", required=False)
@click.argument("label", required=False)
@click.argument("description", required=False)
@click.option("--list", "-l", "list_mode", is_flag=True, help="List annotations")
@click.option("--sheet", "-s", help="Sheet name for listing")
@click.pass_context
def annotate(
    ctx: click.Context,
    range_address: str,
    label: str,
    description: str,
    list_mode: bool,
    sheet: str,
) -> None:
    """Add or list annotations."""
    service = ctx.obj["service"]
    console = Console()
    formatter = ResponseFormatter(console)

    cmd = AnnotateCommand(service, console, formatter)

    if list_mode or not range_address:
        success = cmd.list(sheet=sheet)
    else:
        if not label:
            console.print("[red]Error:[/red] Label is required when adding annotation")
            sys.exit(1)

        success = cmd.add(
            range_address=range_address,
            label=label,
            description=description,
        )

    sys.exit(0 if success else 1)


@cli.command()
@click.argument("query")
@click.option("--sheet", "-s", help="Sheet name to search in")
@click.pass_context
def search(ctx: click.Context, query: str, sheet: str) -> None:
    """Search annotations."""
    service = ctx.obj["service"]
    console = Console()
    formatter = ResponseFormatter(console)

    cmd = SearchCommand(service, console, formatter)
    success = cmd.execute(query=query, sheet=sheet)

    sys.exit(0 if success else 1)


@cli.command()
@click.argument("action", required=False, default="status")
@click.pass_context
def cache(ctx: click.Context, action: str) -> None:
    """Manage dependency graph cache (status, rebuild, clear)."""
    service = ctx.obj["service"]
    console = Console()
    formatter = ResponseFormatter(console)

    cmd = CacheCommand(service, console, formatter)

    if action == "status":
        success = cmd.status()
    elif action == "rebuild":
        success = cmd.rebuild()
    elif action == "clear":
        success = cmd.clear()
    else:
        console.print(f"[red]Error:[/red] Unknown cache action: {action}")
        console.print("Available actions: status, rebuild, clear")
        sys.exit(1)

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    cli()
