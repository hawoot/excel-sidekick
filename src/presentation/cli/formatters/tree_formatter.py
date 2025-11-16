"""Formatter for dependency trees."""

from typing import Set

from rich.console import Console
from rich.tree import Tree

from src.domain.models.dependency import DependencyNode, DependencyTree


class TreeFormatter:
    """Format dependency trees for CLI output."""

    def __init__(self, console: Console):
        """
        Initialize tree formatter.

        Args:
            console: Rich console for output
        """
        self.console = console

    def format_tree(self, dep_tree: DependencyTree, max_depth: int = 5) -> None:
        """
        Format and print dependency tree.

        Args:
            dep_tree: Dependency tree to format
            max_depth: Maximum depth to display
        """
        if not dep_tree.root:
            self.console.print("[yellow]No dependencies found[/yellow]")
            return

        # Create rich tree
        rich_tree = Tree(
            f"[bold cyan]{dep_tree.root.cell_address}[/bold cyan] "
            f"= {self._format_formula(dep_tree.root.formula)}"
        )

        # Track visited to avoid cycles
        visited: Set[str] = set()

        # Add precedents
        if dep_tree.root.precedents:
            precedents_branch = rich_tree.add("[bold] Precedents (inputs)[/bold]")
            for precedent in dep_tree.root.precedents:
                self._add_node(
                    precedents_branch,
                    precedent,
                    visited,
                    depth=1,
                    max_depth=max_depth,
                    is_precedent=True,
                )

        # Add dependents
        if dep_tree.root.dependents:
            dependents_branch = rich_tree.add("[bold]’ Dependents (outputs)[/bold]")
            for dependent in dep_tree.root.dependents:
                self._add_node(
                    dependents_branch,
                    dependent,
                    visited,
                    depth=1,
                    max_depth=max_depth,
                    is_precedent=False,
                )

        # Print tree
        self.console.print(rich_tree)

        # Print summary
        self.console.print(
            f"\n[dim]Total nodes: {dep_tree.total_nodes}, "
            f"Max depth: {dep_tree.max_depth}[/dim]"
        )

    def _add_node(
        self,
        parent_branch: Tree,
        node: DependencyNode,
        visited: Set[str],
        depth: int,
        max_depth: int,
        is_precedent: bool,
    ) -> None:
        """
        Recursively add node to tree.

        Args:
            parent_branch: Parent tree branch
            node: Node to add
            visited: Set of visited nodes (cycle detection)
            depth: Current depth
            max_depth: Maximum depth to display
            is_precedent: Whether this is a precedent (vs dependent)
        """
        # Check depth limit
        if depth > max_depth:
            parent_branch.add("[dim]...[/dim]")
            return

        # Check for cycles
        if node.cell_address in visited:
            parent_branch.add(f"[yellow]{node.cell_address} (circular)[/yellow]")
            return

        visited.add(node.cell_address)

        # Format node label
        if node.formula:
            label = f"[cyan]{node.cell_address}[/cyan] = {self._format_formula(node.formula)}"
        elif node.value is not None:
            label = f"[cyan]{node.cell_address}[/cyan] = [green]{node.value}[/green]"
        else:
            label = f"[cyan]{node.cell_address}[/cyan]"

        # Add this node
        branch = parent_branch.add(label)

        # Recurse into children
        children = node.precedents if is_precedent else node.dependents
        if children:
            for child in children:
                self._add_node(
                    branch,
                    child,
                    visited,
                    depth + 1,
                    max_depth,
                    is_precedent,
                )

    def _format_formula(self, formula: str, max_length: int = 80) -> str:
        """
        Format formula for display.

        Args:
            formula: Formula string
            max_length: Maximum length before truncation

        Returns:
            Formatted formula
        """
        if len(formula) > max_length:
            return f"[dim]{formula[:max_length]}...[/dim]"
        return f"[dim]{formula}[/dim]"
