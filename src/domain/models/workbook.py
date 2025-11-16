"""Domain models for Excel workbooks, sheets, cells, and formulas."""

import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, List, Optional, Set

from src.shared.types import CellAddress, FormulaString, SheetName


@dataclass
class Formula:
    """
    Represents an Excel formula.

    A rich model that can parse and understand formula structure.
    """

    formula_text: FormulaString
    referenced_cells: List[CellAddress] = field(default_factory=list)

    def __post_init__(self) -> None:
        """Parse formula to extract referenced cells."""
        if not self.referenced_cells:
            self.referenced_cells = self._extract_cell_references()

    def _extract_cell_references(self) -> List[CellAddress]:
        """
        Extract all cell references from formula.

        Returns:
            List of cell addresses referenced in the formula

        Examples:
            "=A1+B2" -> ["A1", "B2"]
            "=SUM(Sheet1!A1:A10)" -> ["Sheet1!A1:A10"]
            "=VLOOKUP(A1,Data!B:C,2,FALSE)" -> ["A1", "Data!B:C"]
        """
        references = []

        # Pattern for cell references (with optional sheet name)
        # Matches: A1, $A$1, Sheet1!A1, 'Sheet Name'!A1, A1:B10, etc.
        patterns = [
            # Sheet reference with range (e.g., Sheet1!A1:B10, 'Sheet Name'!A1:B10)
            r"(?:'([^']+)'|(\w+))!(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)",
            # Sheet reference with column/row (e.g., Sheet1!A:A, Sheet1!1:1)
            r"(?:'([^']+)'|(\w+))!(\$?[A-Z]+:\$?[A-Z]+|\$?\d+:\$?\d+)",
            # Simple range (e.g., A1:B10)
            r"(\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+)",
            # Column or row reference (e.g., A:A, 1:1)
            r"(\$?[A-Z]+:\$?[A-Z]+|\$?\d+:\$?\d+)",
            # Single cell (e.g., A1, $A$1)
            r"(\$?[A-Z]+\$?\d+)",
        ]

        formula_upper = self.formula_text.upper()

        for pattern in patterns:
            for match in re.finditer(pattern, formula_upper):
                if match.lastindex == 3:
                    # Sheet reference
                    sheet = match.group(1) or match.group(2)
                    cell_part = match.group(3)
                    ref = f"{sheet}!{cell_part}"
                else:
                    # No sheet reference
                    ref = match.group(0)

                # Clean up $ signs for internal representation
                ref = ref.replace("$", "")

                if ref not in references:
                    references.append(ref)

        return references

    def has_cross_sheet_references(self) -> bool:
        """Check if formula references other sheets."""
        return any("!" in ref for ref in self.referenced_cells)

    def get_referenced_sheets(self) -> Set[SheetName]:
        """
        Get all sheets referenced in formula.

        Returns:
            Set of sheet names referenced
        """
        sheets = set()
        for ref in self.referenced_cells:
            if "!" in ref:
                sheet = ref.split("!")[0]
                sheets.add(sheet)
        return sheets

    def __str__(self) -> str:
        """String representation."""
        return self.formula_text

    def __repr__(self) -> str:
        """Debug representation."""
        return f"Formula('{self.formula_text}')"


@dataclass
class Cell:
    """
    Represents an Excel cell.

    A rich model with methods for understanding cell content and relationships.
    """

    address: CellAddress
    sheet: SheetName
    value: Any
    formula: Optional[Formula] = None
    data_type: Optional[str] = None  # number, text, boolean, error, empty

    @property
    def full_address(self) -> CellAddress:
        """Get full address with sheet name."""
        return f"{self.sheet}!{self.address}"

    def has_formula(self) -> bool:
        """Check if cell contains a formula."""
        return self.formula is not None

    def is_empty(self) -> bool:
        """Check if cell is empty."""
        return self.value is None or (isinstance(self.value, str) and self.value == "")

    def is_calculation(self) -> bool:
        """Check if cell performs a calculation."""
        return self.has_formula()

    def get_direct_dependencies(self) -> List[CellAddress]:
        """
        Get cells that this cell directly depends on.

        Returns:
            List of cell addresses this cell references
        """
        if not self.formula:
            return []
        return self.formula.referenced_cells

    def references_cell(self, cell_address: CellAddress) -> bool:
        """
        Check if this cell references another cell.

        Args:
            cell_address: Address to check

        Returns:
            True if this cell references the given address
        """
        if not self.formula:
            return False
        return cell_address in self.formula.referenced_cells

    def references_sheet(self, sheet_name: SheetName) -> bool:
        """
        Check if this cell references another sheet.

        Args:
            sheet_name: Sheet name to check

        Returns:
            True if this cell references the given sheet
        """
        if not self.formula:
            return False
        return sheet_name in self.formula.get_referenced_sheets()

    def __str__(self) -> str:
        """String representation."""
        if self.formula:
            return f"{self.full_address}: {self.formula} = {self.value}"
        else:
            return f"{self.full_address}: {self.value}"

    def __repr__(self) -> str:
        """Debug representation."""
        return f"Cell('{self.full_address}', value={self.value}, formula={self.formula})"


@dataclass
class Sheet:
    """
    Represents an Excel worksheet.

    Contains metadata about the sheet structure.
    """

    name: SheetName
    used_range: Optional[str] = None  # E.g., "A1:Z500"
    row_count: int = 0
    col_count: int = 0
    formula_count: int = 0

    def __str__(self) -> str:
        """String representation."""
        return f"Sheet('{self.name}', {self.row_count}x{self.col_count}, {self.formula_count} formulas)"

    def __repr__(self) -> str:
        """Debug representation."""
        return self.__str__()


@dataclass
class Workbook:
    """
    Represents an Excel workbook.

    Contains metadata and provides access to sheets.
    """

    name: str
    path: str
    sheets: List[Sheet] = field(default_factory=list)
    active_sheet: Optional[SheetName] = None
    last_modified: Optional[datetime] = None

    def get_sheet(self, sheet_name: SheetName) -> Optional[Sheet]:
        """
        Get sheet by name.

        Args:
            sheet_name: Name of sheet to retrieve

        Returns:
            Sheet object or None if not found
        """
        for sheet in self.sheets:
            if sheet.name == sheet_name:
                return sheet
        return None

    def has_sheet(self, sheet_name: SheetName) -> bool:
        """
        Check if workbook has a sheet with given name.

        Args:
            sheet_name: Name to check

        Returns:
            True if sheet exists
        """
        return any(sheet.name == sheet_name for sheet in self.sheets)

    def get_sheet_names(self) -> List[SheetName]:
        """Get list of all sheet names."""
        return [sheet.name for sheet in self.sheets]

    def total_formula_count(self) -> int:
        """Get total number of formulas across all sheets."""
        return sum(sheet.formula_count for sheet in self.sheets)

    def __str__(self) -> str:
        """String representation."""
        return f"Workbook('{self.name}', {len(self.sheets)} sheets)"

    def __repr__(self) -> str:
        """Debug representation."""
        return self.__str__()


@dataclass
class WorkbookStructure:
    """
    High-level structure of a workbook.

    Used for quick overview without loading all data.
    """

    workbook: Workbook
    sheet_summaries: List[Sheet]

    def __str__(self) -> str:
        """String representation."""
        sheets_info = "\n".join(f"  - {sheet}" for sheet in self.sheet_summaries)
        return f"{self.workbook}\nSheets:\n{sheets_info}"
