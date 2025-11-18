"""Domain models for selections and ranges."""

import re
from dataclasses import dataclass
from typing import Optional, Tuple

from src.shared.exceptions import InvalidRangeError


@dataclass
class Range:
    """
    Represents a cell range in Excel.

    A rich domain model with methods for parsing, validation, and manipulation.

    Examples:
        - "A1:B10" (single sheet, implicit)
        - "Sheet1!A1:B10" (with sheet name)
        - "C5" (single cell)
    """

    sheet: Optional[str]
    start_col: int  # 0-based column index
    start_row: int  # 1-based row index
    end_col: int    # 0-based column index
    end_row: int    # 1-based row index

    @classmethod
    def from_address(cls, address: str) -> "Range":
        """
        Parse range from Excel address string.

        Args:
            address: Excel address like "A1:B10" or "Sheet1!A1:B10"

        Returns:
            Range object

        Raises:
            InvalidRangeError: If address format is invalid
        """
        # Split sheet name if present
        sheet = None
        if "!" in address:
            sheet, address = address.split("!", 1)
            sheet = sheet.strip("'")  # Remove quotes if present

        # Check if single cell or range
        if ":" in address:
            start_addr, end_addr = address.split(":", 1)
        else:
            # Single cell - make it a 1x1 range
            start_addr = end_addr = address

        # Parse start and end
        start_col, start_row = cls._parse_cell_address(start_addr)
        end_col, end_row = cls._parse_cell_address(end_addr)

        # Validate
        if start_row > end_row or start_col > end_col:
            raise InvalidRangeError(
                f"Invalid range: start ({start_addr}) must be before end ({end_addr})"
            )

        return cls(
            sheet=sheet,
            start_col=start_col,
            start_row=start_row,
            end_col=end_col,
            end_row=end_row,
        )

    @staticmethod
    def _parse_cell_address(cell: str) -> Tuple[int, int]:
        """
        Parse a single cell address like "A1" into (col, row).

        Args:
            cell: Cell address like "A1" or "AB123" or "$A$1" (absolute references)

        Returns:
            Tuple of (column_index, row_number) where column is 0-based

        Raises:
            InvalidRangeError: If cell address is invalid
        """
        # Remove dollar signs (absolute references like $A$1) before parsing
        cell_clean = cell.replace("$", "").upper()
        match = re.match(r"^([A-Z]+)(\d+)$", cell_clean)
        if not match:
            raise InvalidRangeError(f"Invalid cell address: {cell}")

        col_letters, row_str = match.groups()

        # Convert column letters to index (A=0, B=1, ..., Z=25, AA=26, etc.)
        col_index = 0
        for char in col_letters:
            col_index = col_index * 26 + (ord(char) - ord('A') + 1)
        col_index -= 1  # Make it 0-based

        row_number = int(row_str)

        return col_index, row_number

    @staticmethod
    def _col_index_to_letter(col_index: int) -> str:
        """
        Convert 0-based column index to Excel column letters.

        Args:
            col_index: 0-based column index (0=A, 1=B, 25=Z, 26=AA, etc.)

        Returns:
            Column letters (e.g., "A", "AB", "ZZ")
        """
        letters = ""
        col_index += 1  # Make it 1-based for calculation

        while col_index > 0:
            col_index -= 1
            letters = chr(col_index % 26 + ord('A')) + letters
            col_index //= 26

        return letters

    def to_address(self, include_sheet: bool = True) -> str:
        """
        Convert range to Excel address string.

        Args:
            include_sheet: Whether to include sheet name

        Returns:
            Excel address string
        """
        start_col_letter = self._col_index_to_letter(self.start_col)
        end_col_letter = self._col_index_to_letter(self.end_col)

        start = f"{start_col_letter}{self.start_row}"
        end = f"{end_col_letter}{self.end_row}"

        if start == end:
            # Single cell
            range_part = start
        else:
            range_part = f"{start}:{end}"

        if include_sheet and self.sheet:
            return f"{self.sheet}!{range_part}"
        else:
            return range_part

    def cell_count(self) -> int:
        """Calculate total number of cells in range."""
        rows = self.end_row - self.start_row + 1
        cols = self.end_col - self.start_col + 1
        return rows * cols

    def row_count(self) -> int:
        """Calculate number of rows in range."""
        return self.end_row - self.start_row + 1

    def col_count(self) -> int:
        """Calculate number of columns in range."""
        return self.end_col - self.start_col + 1

    def is_single_cell(self) -> bool:
        """Check if range represents a single cell."""
        return self.start_row == self.end_row and self.start_col == self.end_col

    def contains(self, other: "Range") -> bool:
        """
        Check if this range contains another range.

        Args:
            other: Another range to check

        Returns:
            True if this range fully contains the other range
        """
        if self.sheet and other.sheet and self.sheet != other.sheet:
            return False

        return (
            self.start_row <= other.start_row
            and self.end_row >= other.end_row
            and self.start_col <= other.start_col
            and self.end_col >= other.end_col
        )

    def overlaps(self, other: "Range") -> bool:
        """
        Check if this range overlaps with another range.

        Args:
            other: Another range to check

        Returns:
            True if ranges overlap
        """
        if self.sheet and other.sheet and self.sheet != other.sheet:
            return False

        # Ranges overlap if they intersect in both dimensions
        rows_overlap = not (self.end_row < other.start_row or self.start_row > other.end_row)
        cols_overlap = not (self.end_col < other.start_col or self.start_col > other.end_col)

        return rows_overlap and cols_overlap

    def expand(self, rows: int = 0, cols: int = 0) -> "Range":
        """
        Expand range by specified number of rows and columns.

        Args:
            rows: Number of rows to add (up and down)
            cols: Number of columns to add (left and right)

        Returns:
            New expanded range
        """
        return Range(
            sheet=self.sheet,
            start_col=max(0, self.start_col - cols),
            start_row=max(1, self.start_row - rows),
            end_col=self.end_col + cols,
            end_row=self.end_row + rows,
        )

    def __str__(self) -> str:
        """String representation."""
        return self.to_address(include_sheet=True)

    def __repr__(self) -> str:
        """Debug representation."""
        return f"Range('{self.to_address()}')"


@dataclass
class Selection:
    """
    Represents a user's selection in Excel.

    Includes the range and additional context about the selection.
    """

    range: Range
    sheet_name: str
    has_formulas: bool = False

    @classmethod
    def from_address(cls, address: str, sheet_name: Optional[str] = None) -> "Selection":
        """
        Create selection from address string.

        Args:
            address: Excel address (with or without sheet name)
            sheet_name: Sheet name if not in address

        Returns:
            Selection object
        """
        range_obj = Range.from_address(address)

        # Use sheet from address or parameter
        final_sheet = range_obj.sheet or sheet_name
        if not final_sheet:
            raise InvalidRangeError("Sheet name must be specified")

        # Update range sheet if needed
        if not range_obj.sheet:
            range_obj.sheet = final_sheet

        return cls(
            range=range_obj,
            sheet_name=final_sheet,
        )

    def to_address(self) -> str:
        """Get full address with sheet name."""
        return self.range.to_address(include_sheet=True)

    def __str__(self) -> str:
        """String representation."""
        return self.to_address()
