"""Dependency analysis service for building and analyzing dependency graphs."""

from typing import List, Optional, Set

from src.domain.models.dependency import (
    DependencyGraph,
    DependencyNode,
    DependencyTree,
    DependencyTreeNode,
)
from src.domain.models.selection import Range
from src.domain.models.workbook import Cell, Workbook
from src.domain.services.workbook_data_service import WorkbookDataService
from src.infrastructure.config.config_loader import Config
from src.infrastructure.storage.graph_cache import GraphCache
from src.shared.exceptions import DependencyGraphError
from src.shared.logging import get_logger
from src.shared.types import CellAddress, TraceDirection

logger = get_logger(__name__)


class DependencyAnalysisService:
    """
    Service for building and analyzing dependency graphs.

    Builds graphs from Excel workbooks and provides tracing capabilities.
    """

    def __init__(
        self,
        config: Config,
        workbook_data_service: WorkbookDataService,
    ):
        """
        Initialize dependency analysis service.

        Args:
            config: Application configuration
            workbook_data_service: Service for reading workbook data
        """
        self.config = config
        self.workbook_data = workbook_data_service
        self.cache = GraphCache(config.dependencies.cache.location)
        self._current_graph: Optional[DependencyGraph] = None

    def build_graph(
        self,
        workbook: Workbook,
        use_cache: bool = True,
    ) -> DependencyGraph:
        """
        Build dependency graph for workbook.

        Args:
            workbook: Workbook to analyze
            use_cache: Whether to use cached graph if available

        Returns:
            Dependency graph

        Raises:
            DependencyGraphError: If graph building fails
        """
        # Check cache
        if use_cache and self.config.dependencies.cache.enabled:
            cached_graph = self._load_from_cache(workbook.path)
            if cached_graph:
                self._current_graph = cached_graph
                return cached_graph

        # Build graph
        logger.info(f"Building dependency graph for {workbook.name}...")

        try:
            graph = DependencyGraph(workbook_name=workbook.name)

            # Process each sheet
            for sheet in workbook.sheets:
                logger.debug(f"Processing sheet: {sheet.name}")
                self._process_sheet(sheet.name, graph)

            # Cache graph
            if self.config.dependencies.cache.enabled:
                self._save_to_cache(graph, workbook.path)

            self._current_graph = graph
            logger.info(
                f"Graph built: {graph.node_count()} nodes, "
                f"{graph.formula_count()} formulas"
            )

            return graph

        except Exception as e:
            raise DependencyGraphError(f"Failed to build dependency graph: {e}")

    def _process_sheet(self, sheet_name: str, graph: DependencyGraph) -> None:
        """
        Process a single sheet and add its dependencies to graph.

        Args:
            sheet_name: Name of sheet to process
            graph: Graph to add nodes to
        """
        # Get sheet structure to find used range
        workbook_structure = self.workbook_data.get_workbook_structure()
        sheet = workbook_structure.workbook.get_sheet(sheet_name)

        if not sheet or not sheet.used_range:
            return

        # Parse used range
        try:
            range_obj = Range.from_address(f"{sheet_name}!{sheet.used_range}")
        except Exception as e:
            logger.warning(f"Failed to parse used range for {sheet_name}: {e}")
            return

        # Get all cells in used range
        try:
            cells = self.workbook_data.get_range_data(range_obj)
            logger.debug(f"Retrieved {len(cells)} cells from {sheet_name}")
        except Exception as e:
            logger.error(f"Failed to get range data for sheet {sheet_name}: {e}")
            logger.warning(f"Skipping sheet {sheet_name} due to error reading cell data")
            return

        # Process cells with formulas
        formula_cells = 0
        for cell in cells:
            if cell.has_formula():
                self._add_cell_to_graph(cell, graph)
                formula_cells += 1

        logger.info(f"Sheet {sheet_name}: Processed {len(cells)} cells, added {formula_cells} formula cells to graph")

    def _add_cell_to_graph(self, cell: Cell, graph: DependencyGraph) -> None:
        """
        Add a cell and its dependencies to the graph.

        Args:
            cell: Cell to add
            graph: Graph to add to
        """
        full_address = cell.full_address

        # Create or get node for this cell
        node = graph.get_node(full_address)
        if node is None:
            node = DependencyNode(
                cell_address=cell.address,
                sheet=cell.sheet,
                formula=str(cell.formula) if cell.formula else None,
            )
            graph.add_node(node)

        # Add dependencies (predecessors)
        if cell.formula:
            for ref_address in cell.formula.referenced_cells:
                # Normalize address (add sheet if missing)
                if "!" not in ref_address:
                    ref_address = f"{cell.sheet}!{ref_address}"

                # Add predecessor
                node.add_predecessor(ref_address)

                # Create node for referenced cell if it doesn't exist
                ref_node = graph.get_node(ref_address)
                if ref_node is None:
                    # Extract sheet and cell address
                    if "!" in ref_address:
                        ref_sheet, ref_cell = ref_address.split("!", 1)
                    else:
                        ref_sheet = cell.sheet
                        ref_cell = ref_address

                    ref_node = DependencyNode(
                        cell_address=ref_cell,
                        sheet=ref_sheet,
                    )
                    graph.add_node(ref_node)

                # Add this cell as successor to referenced cell
                ref_node.add_successor(full_address)

    def trace_dependencies(
        self,
        cell_address: CellAddress,
        direction: TraceDirection = TraceDirection.BOTH,
        depth: Optional[int] = None,
    ) -> DependencyTree:
        """
        Trace dependencies from a cell.

        Args:
            cell_address: Cell to trace from
            direction: Direction to trace (upstream, downstream, both)
            depth: Maximum depth (None for config default)

        Returns:
            Dependency tree

        Raises:
            DependencyGraphError: If graph not built or tracing fails
        """
        if self._current_graph is None:
            raise DependencyGraphError("No dependency graph available. Build graph first.")

        if depth is None:
            depth = self.config.dependencies.default_depth

        logger.debug(
            f"Tracing dependencies: {cell_address}, "
            f"direction={direction}, depth={depth}"
        )

        # Get starting node
        node = self._current_graph.get_node(cell_address)
        if node is None:
            raise DependencyGraphError(f"Cell not found in graph: {cell_address}")

        # Build tree
        root = DependencyTreeNode(
            cell_address=node.cell_address,
            sheet=node.sheet,
            formula=node.formula,
            depth=0,
        )

        # Trace based on direction
        visited: Set[str] = set()

        if direction in (TraceDirection.UPSTREAM, TraceDirection.BOTH):
            self._trace_upstream(root, depth, visited)

        if direction in (TraceDirection.DOWNSTREAM, TraceDirection.BOTH):
            # Reset visited for downstream trace
            if direction == TraceDirection.BOTH:
                visited = set()
            self._trace_downstream(root, depth, visited)

        tree = DependencyTree(
            root=root,
            direction=direction.value,
            max_depth=depth,
        )

        logger.debug(f"Trace complete: {tree.total_nodes()} nodes")

        return tree

    def _trace_upstream(
        self,
        node: DependencyTreeNode,
        max_depth: int,
        visited: Set[str],
    ) -> None:
        """
        Recursively trace upstream dependencies.

        Args:
            node: Current node
            max_depth: Maximum depth to trace
            visited: Set of visited cells (to prevent cycles)
        """
        if node.depth >= max_depth:
            return

        full_address = node.full_address
        if full_address in visited:
            return

        visited.add(full_address)

        # Get graph node
        graph_node = self._current_graph.get_node(full_address)
        if graph_node is None:
            return

        # Add predecessors as children
        for pred_address in graph_node.predecessors:
            pred_graph_node = self._current_graph.get_node(pred_address)
            if pred_graph_node is None:
                continue

            child = DependencyTreeNode(
                cell_address=pred_graph_node.cell_address,
                sheet=pred_graph_node.sheet,
                formula=pred_graph_node.formula,
                depth=node.depth + 1,
            )

            node.add_child(child)

            # Recurse
            self._trace_upstream(child, max_depth, visited)

    def _trace_downstream(
        self,
        node: DependencyTreeNode,
        max_depth: int,
        visited: Set[str],
    ) -> None:
        """
        Recursively trace downstream dependencies.

        Args:
            node: Current node
            max_depth: Maximum depth to trace
            visited: Set of visited cells (to prevent cycles)
        """
        if node.depth >= max_depth:
            return

        full_address = node.full_address
        if full_address in visited:
            return

        visited.add(full_address)

        # Get graph node
        graph_node = self._current_graph.get_node(full_address)
        if graph_node is None:
            return

        # Add successors as children
        for succ_address in graph_node.successors:
            succ_graph_node = self._current_graph.get_node(succ_address)
            if succ_graph_node is None:
                continue

            child = DependencyTreeNode(
                cell_address=succ_graph_node.cell_address,
                sheet=succ_graph_node.sheet,
                formula=succ_graph_node.formula,
                depth=node.depth + 1,
            )

            node.add_child(child)

            # Recurse
            self._trace_downstream(child, max_depth, visited)

    def get_current_graph(self) -> Optional[DependencyGraph]:
        """Get current dependency graph."""
        return self._current_graph

    def _load_from_cache(self, workbook_path: str) -> Optional[DependencyGraph]:
        """
        Load graph from cache if available and not stale.

        Args:
            workbook_path: Path to workbook

        Returns:
            Cached graph or None
        """
        if self.cache.is_stale(workbook_path):
            logger.debug("Cache is stale, will rebuild")
            return None

        graph = self.cache.load(workbook_path)
        if graph:
            logger.info("Loaded dependency graph from cache")

        return graph

    def _save_to_cache(self, graph: DependencyGraph, workbook_path: str) -> None:
        """
        Save graph to cache.

        Args:
            graph: Graph to save
            workbook_path: Path to workbook
        """
        try:
            self.cache.save(graph, workbook_path)
        except Exception as e:
            logger.warning(f"Failed to cache dependency graph: {e}")

    def clear_cache(self, workbook_path: str) -> None:
        """
        Clear cached graph.

        Args:
            workbook_path: Path to workbook
        """
        self.cache.clear(workbook_path)
        logger.info("Cache cleared")

    def rebuild_graph(self, workbook: Workbook) -> DependencyGraph:
        """
        Force rebuild of dependency graph (bypasses cache).

        Args:
            workbook: Workbook to analyze

        Returns:
            Newly built dependency graph
        """
        return self.build_graph(workbook, use_cache=False)
