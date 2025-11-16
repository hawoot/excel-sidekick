"""Snapshot generator for creating markdown views of Excel data."""

from typing import List, Optional

from src.domain.models.selection import Range
from src.domain.models.workbook import Cell
from src.infrastructure.config.config_loader import Config
from src.shared.logging import get_logger

logger = get_logger(__name__)


class SnapshotGenerator:
    """
    Generates markdown snapshots of Excel ranges.

    Handles large ranges by collapsing empty rows/columns
    and sampling data intelligently.
    """

    def __init__(self, config: Config):
        """
        Initialize snapshot generator.

        Args:
            config: Application configuration
        """
        self.config = config
        self.snapshot_config = config.snapshot

    def generate(
        self,
        cells: List[Cell],
        range_obj: Range,
        strategy: Optional[str] = None,
    ) -> str:
        """
        Generate markdown snapshot of cell data.

        Args:
            cells: List of cells to include in snapshot
            range_obj: Range being snapshotted
            strategy: Override strategy ("auto", "full", or None for config default)

        Returns:
            Markdown formatted snapshot
        """
        if not cells:
            return self._empty_snapshot(range_obj)

        # Determine strategy
        if strategy is None:
            strategy = "auto"

        cell_count = len(cells)
        max_cells = self.snapshot_config.max_cells_per_snapshot

        # Auto strategy selection
        if strategy == "auto":
            if cell_count > self.snapshot_config.sampling.threshold_cells:
                strategy = "sampled"
            else:
                strategy = "full"

        # Check if we need to warn about size
        if cell_count > max_cells:
            logger.warning(
                f"Range has {cell_count} cells, exceeding max {max_cells}. "
                f"Using sampled strategy."
            )
            strategy = "sampled"

        # Generate based on strategy
        if strategy == "sampled":
            return self._generate_sampled(cells, range_obj)
        else:
            return self._generate_full(cells, range_obj)

    def _generate_full(self, cells: List[Cell], range_obj: Range) -> str:
        """Generate full snapshot with all cells."""
        lines = []

        # Header
        lines.append(f"## Snapshot: {range_obj.to_address()}")
        lines.append(f"")
        lines.append(f"Size: {range_obj.row_count()} rows × {range_obj.col_count()} columns")
        lines.append(f"")

        # Build table
        table = self._build_table(cells, range_obj)
        lines.extend(table)

        return "\n".join(lines)

    def _generate_sampled(self, cells: List[Cell], range_obj: Range) -> str:
        """Generate sampled snapshot for large ranges."""
        lines = []

        # Header
        lines.append(f"## Snapshot: {range_obj.to_address()} (Sampled)")
        lines.append(f"")
        lines.append(f"Size: {range_obj.row_count()} rows × {range_obj.col_count()} columns")
        lines.append(f"*Showing sample due to size*")
        lines.append(f"")

        # Sample cells
        sampled_cells = self._sample_cells(cells, range_obj)

        # Build table
        table = self._build_table(sampled_cells, range_obj, is_sampled=True)
        lines.extend(table)

        return "\n".join(lines)

    def _sample_cells(self, cells: List[Cell], range_obj: Range) -> List[Cell]:
        """
        Sample cells from a large range.

        Always includes:
        - First N rows (headers)
        - Last N rows (totals)
        - Every Nth row in between
        """
        sampling_config = self.snapshot_config.sampling

        rows_per_row = range_obj.col_count()
        total_rows = range_obj.row_count()

        # Calculate which rows to include
        rows_to_include = set()

        # First N rows
        for i in range(min(sampling_config.always_show_first_n, total_rows)):
            rows_to_include.add(i)

        # Last N rows
        for i in range(max(0, total_rows - sampling_config.always_show_last_n), total_rows):
            rows_to_include.add(i)

        # Sampled rows in between
        sample_every = sampling_config.sample_every_n_rows
        for i in range(sampling_config.always_show_first_n, total_rows - sampling_config.always_show_last_n, sample_every):
            rows_to_include.add(i)

        # Filter cells
        sampled = []
        for idx, cell in enumerate(cells):
            row_idx = idx // rows_per_row
            if row_idx in rows_to_include:
                sampled.append(cell)

        return sampled

    def _build_table(
        self,
        cells: List[Cell],
        range_obj: Range,
        is_sampled: bool = False,
    ) -> List[str]:
        """
        Build markdown table from cells.

        Args:
            cells: List of cells
            range_obj: Range info
            is_sampled: Whether this is a sampled view

        Returns:
            List of markdown table lines
        """
        lines = []

        rows_count = range_obj.row_count()
        cols_count = range_obj.col_count()

        if rows_count == 0 or cols_count == 0:
            return ["*Empty range*"]

        # Build grid
        grid = [[None for _ in range(cols_count)] for _ in range(rows_count)]

        for cell in cells:
            # Parse cell address to get row/col
            row_num = int(''.join(filter(str.isdigit, cell.address)))
            col_letter = ''.join(filter(str.isalpha, cell.address))

            # Calculate position in grid
            row_idx = row_num - range_obj.start_row
            col_idx = self._col_letter_to_index(col_letter) - range_obj.start_col

            if 0 <= row_idx < rows_count and 0 <= col_idx < cols_count:
                grid[row_idx][col_idx] = cell

        # Generate column headers (A, B, C, ...)
        col_headers = []
        for col_idx in range(cols_count):
            actual_col = range_obj.start_col + col_idx
            col_headers.append(Range._col_index_to_letter(actual_col))

        # Header row
        header = "| " + " | ".join(col_headers) + " |"
        lines.append(header)

        # Separator
        separator = "|" + "|".join(["---" for _ in range(cols_count)]) + "|"
        lines.append(separator)

        # Data rows
        last_included_row = -1
        for row_idx in range(rows_count):
            row_cells = grid[row_idx]

            # Check if row should be included (if sampled)
            if is_sampled and not any(row_cells):
                # Skip empty rows in sampled view
                continue

            # Check for gap in sampled view
            if is_sampled and row_idx > last_included_row + 1:
                # Show gap indicator
                gap_size = row_idx - last_included_row - 1
                lines.append(f"| *... {gap_size} rows omitted ...* " + "| " * (cols_count - 1) + "|")

            last_included_row = row_idx

            # Format row
            row_values = []
            for cell in row_cells:
                if cell is None:
                    row_values.append("")
                else:
                    # Format cell value
                    value_str = self._format_cell_value(cell)
                    row_values.append(value_str)

            row_line = "| " + " | ".join(row_values) + " |"
            lines.append(row_line)

        return lines

    def _format_cell_value(self, cell: Cell) -> str:
        """
        Format cell value for display.

        Args:
            cell: Cell to format

        Returns:
            Formatted string
        """
        # Show formula if present
        if cell.formula:
            formula_str = str(cell.formula)
            value_str = f"{formula_str}"
            # Add value in parentheses if different
            if cell.value is not None:
                value_str += f" `[{cell.value}]`"
            return value_str

        # Show value
        if cell.value is None or cell.value == "":
            return ""

        # Format based on type
        if isinstance(cell.value, float):
            # Format numbers nicely
            if cell.value == int(cell.value):
                return str(int(cell.value))
            else:
                return f"{cell.value:.2f}"
        else:
            return str(cell.value)

    def _empty_snapshot(self, range_obj: Range) -> str:
        """Generate snapshot for empty range."""
        return f"## Snapshot: {range_obj.to_address()}\n\n*Range is empty*"

    @staticmethod
    def _col_letter_to_index(col_letter: str) -> int:
        """Convert column letter to 0-based index."""
        col_index = 0
        for char in col_letter.upper():
            col_index = col_index * 26 + (ord(char) - ord('A') + 1)
        return col_index - 1
