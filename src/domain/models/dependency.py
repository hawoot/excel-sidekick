"""Domain models for dependency graphs and trees."""

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Set

from src.shared.types import CellAddress, SheetName


@dataclass
class DependencyNode:
    """
    Represents a node in the dependency graph.

    Each node is a cell that may have dependencies (predecessors)
    and dependents (successors).
    """

    cell_address: CellAddress
    sheet: SheetName
    formula: Optional[str] = None
    predecessors: Set[CellAddress] = field(default_factory=set)  # Cells this depends on
    successors: Set[CellAddress] = field(default_factory=set)    # Cells that depend on this

    def add_predecessor(self, cell_address: CellAddress) -> None:
        """Add a cell that this node depends on."""
        self.predecessors.add(cell_address)

    def add_successor(self, cell_address: CellAddress) -> None:
        """Add a cell that depends on this node."""
        self.successors.add(cell_address)

    def has_dependencies(self) -> bool:
        """Check if this node has any predecessors."""
        return len(self.predecessors) > 0

    def is_depended_on(self) -> bool:
        """Check if any cells depend on this node."""
        return len(self.successors) > 0

    def __str__(self) -> str:
        """String representation."""
        full_address = f"{self.sheet}!{self.cell_address}"
        deps = f"{len(self.predecessors)} deps" if self.predecessors else "no deps"
        uses = f"{len(self.successors)} uses" if self.successors else "no uses"
        return f"{full_address} ({deps}, {uses})"

    def __repr__(self) -> str:
        """Debug representation."""
        return f"DependencyNode('{self.sheet}!{self.cell_address}')"

    def __hash__(self) -> int:
        """Hash for use in sets."""
        return hash(f"{self.sheet}!{self.cell_address}")


@dataclass
class DependencyGraph:
    """
    Represents the complete dependency graph for a workbook.

    Stores all cells with formulas and their relationships.
    """

    nodes: Dict[CellAddress, DependencyNode] = field(default_factory=dict)
    workbook_name: Optional[str] = None

    def add_node(self, node: DependencyNode) -> None:
        """Add a node to the graph."""
        full_address = f"{node.sheet}!{node.cell_address}"
        self.nodes[full_address] = node

    def get_node(self, cell_address: CellAddress) -> Optional[DependencyNode]:
        """
        Get node by address.

        Args:
            cell_address: Cell address (with or without sheet)

        Returns:
            Dependency node or None if not found
        """
        # Try with and without sheet prefix
        if cell_address in self.nodes:
            return self.nodes[cell_address]

        # Try adding each sheet prefix
        for addr in self.nodes:
            if addr.endswith(f"!{cell_address}"):
                return self.nodes[addr]

        return None

    def get_predecessors(self, cell_address: CellAddress) -> Set[CellAddress]:
        """
        Get all cells that the given cell depends on.

        Args:
            cell_address: Cell to get predecessors for

        Returns:
            Set of predecessor cell addresses
        """
        node = self.get_node(cell_address)
        if node:
            return node.predecessors
        return set()

    def get_successors(self, cell_address: CellAddress) -> Set[CellAddress]:
        """
        Get all cells that depend on the given cell.

        Args:
            cell_address: Cell to get successors for

        Returns:
            Set of successor cell addresses
        """
        node = self.get_node(cell_address)
        if node:
            return node.successors
        return set()

    def node_count(self) -> int:
        """Get total number of nodes in graph."""
        return len(self.nodes)

    def formula_count(self) -> int:
        """Get count of cells with formulas."""
        return sum(1 for node in self.nodes.values() if node.formula)

    def __str__(self) -> str:
        """String representation."""
        name = f" ({self.workbook_name})" if self.workbook_name else ""
        return f"DependencyGraph{name}: {self.node_count()} nodes, {self.formula_count()} formulas"


@dataclass
class DependencyTreeNode:
    """
    Represents a node in a dependency tree (traced from a specific cell).

    Used for visualization and understanding dependency chains.
    """

    cell_address: CellAddress
    sheet: SheetName
    formula: Optional[str] = None
    value: Optional[any] = None
    depth: int = 0
    children: List["DependencyTreeNode"] = field(default_factory=list)

    @property
    def full_address(self) -> CellAddress:
        """Get full address with sheet."""
        return f"{self.sheet}!{self.cell_address}"

    def add_child(self, child: "DependencyTreeNode") -> None:
        """Add a child node."""
        self.children.append(child)

    def is_leaf(self) -> bool:
        """Check if this is a leaf node (no children)."""
        return len(self.children) == 0

    def total_nodes(self) -> int:
        """Count total nodes in subtree (including self)."""
        return 1 + sum(child.total_nodes() for child in self.children)

    def max_depth(self) -> int:
        """Get maximum depth of subtree."""
        if self.is_leaf():
            return self.depth
        return max(child.max_depth() for child in self.children)

    def __str__(self) -> str:
        """String representation."""
        indent = "  " * self.depth
        formula_info = f" = {self.formula}" if self.formula else ""
        value_info = f" [{self.value}]" if self.value is not None else ""
        return f"{indent}{self.full_address}{formula_info}{value_info}"

    def __repr__(self) -> str:
        """Debug representation."""
        return f"DependencyTreeNode('{self.full_address}', depth={self.depth}, children={len(self.children)})"


@dataclass
class DependencyTree:
    """
    Represents a dependency tree traced from a specific cell.

    Shows the full chain of dependencies in a hierarchical structure.
    """

    root: DependencyTreeNode
    direction: str  # "upstream" | "downstream" | "both"
    max_depth: int

    def total_nodes(self) -> int:
        """Get total number of nodes in tree."""
        return self.root.total_nodes()

    def actual_max_depth(self) -> int:
        """Get actual maximum depth reached."""
        return self.root.max_depth()

    def to_lines(self) -> List[str]:
        """
        Convert tree to list of formatted strings (for display).

        Returns:
            List of lines representing the tree
        """
        lines = []
        self._add_node_lines(self.root, lines)
        return lines

    def _add_node_lines(self, node: DependencyTreeNode, lines: List[str]) -> None:
        """Recursively add node and children to lines."""
        lines.append(str(node))
        for child in node.children:
            self._add_node_lines(child, lines)

    def __str__(self) -> str:
        """String representation."""
        lines = self.to_lines()
        header = f"DependencyTree ({self.direction}, max_depth={self.max_depth}, {self.total_nodes()} nodes):"
        return header + "\n" + "\n".join(lines)
