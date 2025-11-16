"""Workbook data service for Excel data access."""

from typing import List, Optional

from src.domain.models.selection import Range, Selection
from src.domain.models.workbook import Cell, Workbook, WorkbookStructure
from src.infrastructure.config.config_loader import Config
from src.infrastructure.excel.snapshot_generator import SnapshotGenerator
from src.infrastructure.excel.workbook_discovery import WorkbookInfo
from src.infrastructure.excel.xlwings_connector import XlwingsConnector
from src.shared.exceptions import ExcelConnectionError
from src.shared.logging import get_logger
from src.shared.types import CellAddress, SheetName

logger = get_logger(__name__)


class WorkbookDataService:
    """
    Service for accessing Excel workbook data.

    Wraps the Excel connector and provides high-level methods
    for reading workbook structure, cells, ranges, and snapshots.
    """

    def __init__(self, config: Config):
        """
        Initialize workbook data service.

        Args:
            config: Application configuration
        """
        self.config = config
        self.connector = XlwingsConnector()
        self.snapshot_generator = SnapshotGenerator(config)
        self._workbook: Optional[Workbook] = None

    def connect(self, workbook_name: Optional[str] = None) -> Workbook:
        """
        Connect to Excel workbook.

        DEPRECATED: Use connect_to_workbook_info() for more reliable connection.

        Args:
            workbook_name: Name of workbook (None for active workbook)

        Returns:
            Workbook model

        Raises:
            ExcelConnectionError: If connection fails
        """
        logger.info(
            f"Connecting to workbook: {workbook_name or 'active workbook'}"
        )

        self._workbook = self.connector.connect(workbook_name)
        logger.info(
            f"Connected to '{self._workbook.name}' "
            f"({len(self._workbook.sheets)} sheets)"
        )

        return self._workbook

    def connect_to_workbook_info(self, workbook_info: WorkbookInfo) -> Workbook:
        """
        Connect to Excel workbook using WorkbookInfo.

        This is the preferred connection method.

        Args:
            workbook_info: WorkbookInfo from discovery

        Returns:
            Workbook model

        Raises:
            ExcelConnectionError: If connection fails
        """
        logger.info(
            f"Connecting to workbook: {workbook_info.workbook_name} "
            f"(Excel PID: {workbook_info.excel_pid})"
        )

        self._workbook = self.connector.connect_to_workbook_info(workbook_info)
        logger.info(
            f"Connected to '{self._workbook.name}' "
            f"({len(self._workbook.sheets)} sheets)"
        )

        return self._workbook

    def disconnect(self) -> None:
        """Disconnect from Excel."""
        self.connector.disconnect()
        self._workbook = None
        logger.info("Disconnected from Excel")

    def is_connected(self) -> bool:
        """Check if currently connected to a workbook."""
        return self.connector.is_connected()

    def get_workbook_structure(self) -> WorkbookStructure:
        """
        Get high-level workbook structure.

        Returns:
            WorkbookStructure with overview

        Raises:
            ExcelConnectionError: If not connected
        """
        if not self.is_connected():
            raise ExcelConnectionError("Not connected to workbook")

        return self.connector.get_workbook_structure()

    def get_current_selection(self) -> Optional[Selection]:
        """
        Get user's current selection in Excel.

        Returns:
            Selection or None if no selection

        Raises:
            ExcelConnectionError: If not connected
        """
        if not self.is_connected():
            raise ExcelConnectionError("Not connected to workbook")

        return self.connector.get_current_selection()

    def get_active_sheet(self) -> SheetName:
        """
        Get name of active sheet.

        Returns:
            Active sheet name

        Raises:
            ExcelConnectionError: If not connected
        """
        if not self.is_connected():
            raise ExcelConnectionError("Not connected to workbook")

        return self.connector.get_active_sheet()

    def get_cell(
        self,
        address: CellAddress,
        sheet: Optional[SheetName] = None,
    ) -> Cell:
        """
        Get single cell data.

        Args:
            address: Cell address (e.g., "A1")
            sheet: Sheet name (None for active sheet)

        Returns:
            Cell model

        Raises:
            ExcelConnectionError: If not connected
        """
        if not self.is_connected():
            raise ExcelConnectionError("Not connected to workbook")

        return self.connector.get_cell(address, sheet)

    def get_range_data(self, range_obj: Range) -> List[Cell]:
        """
        Get data for a range of cells.

        Args:
            range_obj: Range to read

        Returns:
            List of cells

        Raises:
            ExcelConnectionError: If not connected
        """
        if not self.is_connected():
            raise ExcelConnectionError("Not connected to workbook")

        return self.connector.get_range_data(range_obj)

    def get_snapshot(
        self,
        sheet: SheetName,
        range_address: str,
        strategy: Optional[str] = None,
    ) -> str:
        """
        Get markdown snapshot of a range.

        Args:
            sheet: Sheet name
            range_address: Range address (e.g., "A1:B10")
            strategy: Override strategy ("auto", "full", "sampled")

        Returns:
            Markdown formatted snapshot

        Raises:
            ExcelConnectionError: If not connected
        """
        if not self.is_connected():
            raise ExcelConnectionError("Not connected to workbook")

        # Parse range
        range_obj = Range.from_address(f"{sheet}!{range_address}")

        # Get cell data
        cells = self.get_range_data(range_obj)

        # Generate snapshot
        snapshot = self.snapshot_generator.generate(cells, range_obj, strategy)

        return snapshot

    def expand_selection_context(
        self,
        selection: Selection,
        rows: Optional[int] = None,
        cols: Optional[int] = None,
    ) -> Range:
        """
        Expand selection to include surrounding context.

        Args:
            selection: Current selection
            rows: Rows to expand (uses config default if None)
            cols: Columns to expand (uses config default if None)

        Returns:
            Expanded range
        """
        if rows is None:
            rows = self.config.selection.expand_rows
        if cols is None:
            cols = self.config.selection.expand_cols

        return selection.range.expand(rows=rows, cols=cols)

    def search_cells(
        self,
        value_contains: Optional[str] = None,
        sheet: Optional[SheetName] = None,
    ) -> List[CellAddress]:
        """
        Search for cells matching criteria.

        Args:
            value_contains: Search for cells containing this text
            sheet: Limit search to specific sheet (None for all sheets)

        Returns:
            List of matching cell addresses

        Note:
            This is a basic implementation. For production, would need
            more sophisticated search with formula matching, etc.
        """
        # TODO: Implement full search functionality
        # For now, this is a placeholder that would need to iterate
        # through ranges and check values
        raise NotImplementedError("Cell search not yet implemented")

    def get_workbook_info(self) -> Workbook:
        """
        Get current workbook info.

        Returns:
            Workbook model

        Raises:
            ExcelConnectionError: If not connected
        """
        if not self.is_connected():
            raise ExcelConnectionError("Not connected to workbook")

        if self._workbook is None:
            raise ExcelConnectionError("No workbook loaded")

        return self._workbook

    def __enter__(self) -> "WorkbookDataService":
        """Context manager entry."""
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        """Context manager exit."""
        self.disconnect()
