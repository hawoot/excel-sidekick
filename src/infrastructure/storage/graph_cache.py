"""Cache storage for dependency graphs."""

import hashlib
import json
from datetime import datetime
from pathlib import Path
from typing import Dict, Optional

from src.domain.models.dependency import DependencyGraph, DependencyNode
from src.shared.exceptions import CacheError
from src.shared.logging import get_logger

logger = get_logger(__name__)


class GraphCache:
    """
    Manages caching of dependency graphs to disk.

    Caches are stored as JSON files with metadata about freshness.
    """

    def __init__(self, cache_dir: str = ".cache"):
        """
        Initialize graph cache.

        Args:
            cache_dir: Directory to store cache files
        """
        self.cache_dir = Path(cache_dir)
        self.cache_dir.mkdir(parents=True, exist_ok=True)

    def get_cache_path(self, workbook_path: str) -> Path:
        """
        Get cache file path for a workbook.

        Args:
            workbook_path: Path to Excel workbook

        Returns:
            Path to cache file
        """
        # Create safe filename from workbook path
        workbook_name = Path(workbook_path).stem
        safe_name = "".join(c if c.isalnum() else "_" for c in workbook_name)
        return self.cache_dir / f"{safe_name}_graph.json"

    def get_metadata_path(self, workbook_path: str) -> Path:
        """
        Get metadata file path for a workbook.

        Args:
            workbook_path: Path to Excel workbook

        Returns:
            Path to metadata file
        """
        cache_path = self.get_cache_path(workbook_path)
        return cache_path.with_suffix(".meta.json")

    def save(
        self,
        graph: DependencyGraph,
        workbook_path: str,
        workbook_hash: Optional[str] = None,
    ) -> None:
        """
        Save dependency graph to cache.

        Args:
            graph: Dependency graph to save
            workbook_path: Path to Excel workbook
            workbook_hash: Optional hash of workbook formulas

        Raises:
            CacheError: If save fails
        """
        try:
            cache_path = self.get_cache_path(workbook_path)
            metadata_path = self.get_metadata_path(workbook_path)

            # Serialize graph
            graph_data = self._serialize_graph(graph)

            # Save graph
            with open(cache_path, "w") as f:
                json.dump(graph_data, f, indent=2)

            # Save metadata
            metadata = {
                "workbook_path": str(workbook_path),
                "workbook_name": graph.workbook_name,
                "cached_at": datetime.now().isoformat(),
                "node_count": graph.node_count(),
                "formula_count": graph.formula_count(),
                "workbook_hash": workbook_hash,
            }

            with open(metadata_path, "w") as f:
                json.dump(metadata, f, indent=2)

            logger.info(
                f"Saved dependency graph to cache: {cache_path} "
                f"({graph.node_count()} nodes, {graph.formula_count()} formulas)"
            )

        except Exception as e:
            raise CacheError(f"Failed to save graph cache: {e}")

    def load(self, workbook_path: str) -> Optional[DependencyGraph]:
        """
        Load dependency graph from cache.

        Args:
            workbook_path: Path to Excel workbook

        Returns:
            Dependency graph or None if not found/invalid

        Raises:
            CacheError: If load fails unexpectedly
        """
        try:
            cache_path = self.get_cache_path(workbook_path)

            if not cache_path.exists():
                logger.debug(f"No cache found for {workbook_path}")
                return None

            # Load graph data
            with open(cache_path, "r") as f:
                graph_data = json.load(f)

            # Deserialize graph
            graph = self._deserialize_graph(graph_data)

            logger.info(
                f"Loaded dependency graph from cache: {cache_path} "
                f"({graph.node_count()} nodes, {graph.formula_count()} formulas)"
            )

            return graph

        except FileNotFoundError:
            return None
        except json.JSONDecodeError as e:
            logger.warning(f"Invalid cache file format: {e}")
            return None
        except Exception as e:
            raise CacheError(f"Failed to load graph cache: {e}")

    def get_metadata(self, workbook_path: str) -> Optional[Dict]:
        """
        Get cache metadata.

        Args:
            workbook_path: Path to Excel workbook

        Returns:
            Metadata dict or None if not found
        """
        try:
            metadata_path = self.get_metadata_path(workbook_path)

            if not metadata_path.exists():
                return None

            with open(metadata_path, "r") as f:
                return json.load(f)

        except Exception as e:
            logger.warning(f"Failed to load cache metadata: {e}")
            return None

    def is_stale(
        self,
        workbook_path: str,
        current_hash: Optional[str] = None,
    ) -> bool:
        """
        Check if cache is stale.

        Args:
            workbook_path: Path to Excel workbook
            current_hash: Current hash of workbook formulas

        Returns:
            True if cache is stale or doesn't exist
        """
        metadata = self.get_metadata(workbook_path)

        if metadata is None:
            return True

        # Check hash if provided
        if current_hash and metadata.get("workbook_hash"):
            return metadata["workbook_hash"] != current_hash

        # Check file modification time
        try:
            wb_path = Path(workbook_path)
            if not wb_path.exists():
                return True

            cached_at = datetime.fromisoformat(metadata["cached_at"])
            modified_at = datetime.fromtimestamp(wb_path.stat().st_mtime)

            return modified_at > cached_at

        except Exception:
            return True

    def clear(self, workbook_path: str) -> None:
        """
        Clear cache for a workbook.

        Args:
            workbook_path: Path to Excel workbook
        """
        try:
            cache_path = self.get_cache_path(workbook_path)
            metadata_path = self.get_metadata_path(workbook_path)

            if cache_path.exists():
                cache_path.unlink()
                logger.info(f"Cleared cache: {cache_path}")

            if metadata_path.exists():
                metadata_path.unlink()

        except Exception as e:
            logger.warning(f"Failed to clear cache: {e}")

    @staticmethod
    def _serialize_graph(graph: DependencyGraph) -> Dict:
        """Serialize dependency graph to dict."""
        nodes_data = {}

        for cell_address, node in graph.nodes.items():
            nodes_data[cell_address] = {
                "cell_address": node.cell_address,
                "sheet": node.sheet,
                "formula": node.formula,
                "predecessors": list(node.predecessors),
                "successors": list(node.successors),
            }

        return {
            "workbook_name": graph.workbook_name,
            "nodes": nodes_data,
        }

    @staticmethod
    def _deserialize_graph(data: Dict) -> DependencyGraph:
        """Deserialize dependency graph from dict."""
        graph = DependencyGraph(workbook_name=data.get("workbook_name"))

        for cell_address, node_data in data["nodes"].items():
            node = DependencyNode(
                cell_address=node_data["cell_address"],
                sheet=node_data["sheet"],
                formula=node_data.get("formula"),
                predecessors=set(node_data.get("predecessors", [])),
                successors=set(node_data.get("successors", [])),
            )
            graph.add_node(node)

        return graph

    @staticmethod
    def compute_workbook_hash(formulas: list[str]) -> str:
        """
        Compute hash of workbook formulas.

        Args:
            formulas: List of formula strings from workbook

        Returns:
            MD5 hash of formulas
        """
        # Sort formulas for consistent hashing
        sorted_formulas = sorted(formulas)
        content = "\n".join(sorted_formulas)
        return hashlib.md5(content.encode()).hexdigest()
