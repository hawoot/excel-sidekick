"""Excel connector using xlwings for Windows/Mac integration."""

from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

try:
    import xlwings as xw
except ImportError:
    xw = None  # Handle case where xlwings is not available

from src.domain.models.selection import Range, Selection
from src.domain.models.workbook import Cell, Formula, Sheet, Workbook, WorkbookStructure
from src.infrastructure.excel.workbook_discovery import WorkbookInfo
from src.shared.exceptions import (
    ExcelConnectionError,
    InvalidRangeError,
    SheetNotFoundError,
    WorkbookNotFoundError,
)
from src.shared.logging import get_logger
from src.shared.types import CellAddress, SheetName

logger = get_logger(__name__)


class XlwingsConnector:
    """
    Excel connector using xlwings library.

    Provides connection to Excel and methods to read workbook data.
    """

    def __init__(self) -> None:
        """Initialize connector."""
        if xw is None:
            raise ExcelConnectionError(
                "xlwings is not installed. Please install it with: pip install xlwings"
            )

        self._app: Optional[xw.App] = None
        self._workbook: Optional[xw.Book] = None
        self._connected = False

    def connect(self, workbook_name: Optional[str] = None) -> Workbook:
        """
        Connect to Excel workbook.

        DEPRECATED: Use connect_to_workbook_info() for more reliable connection.

        Args:
            workbook_name: Name of workbook to connect to.
                          If None, connects to active workbook.

        Returns:
            Workbook domain model

        Raises:
            ExcelConnectionError: If connection fails
            WorkbookNotFoundError: If workbook is not found
        """
        try:
            # Get or create Excel application instance
            if workbook_name:
                # Connect to specific workbook
                try:
                    self._workbook = xw.Book(workbook_name)
                    self._app = self._workbook.app
                except Exception as e:
                    raise WorkbookNotFoundError(
                        f"Workbook '{workbook_name}' not found. "
                        f"Make sure it's open in Excel. Error: {e}"
                    )
            else:
                # Connect to active workbook
                try:
                    self._app = xw.apps.active
                    if self._app is None:
                        raise ExcelConnectionError("No active Excel application found")

                    self._workbook = self._app.books.active
                    if self._workbook is None:
                        raise WorkbookNotFoundError("No active workbook found")
                except Exception as e:
                    raise ExcelConnectionError(f"Failed to connect to active Excel: {e}")

            self._connected = True
            logger.info(f"Connected to workbook: {self._workbook.name}")

            # Build workbook model
            return self._build_workbook_model()

        except (ExcelConnectionError, WorkbookNotFoundError):
            raise
        except Exception as e:
            raise ExcelConnectionError(f"Unexpected error connecting to Excel: {e}")

    def connect_to_workbook_info(self, workbook_info: WorkbookInfo) -> Workbook:
        """
        Connect to Excel workbook using WorkbookInfo.

        This is the preferred connection method as it uniquely identifies
        workbooks by both path and Excel instance PID.

        Args:
            workbook_info: WorkbookInfo object from workbook discovery

        Returns:
            Workbook domain model

        Raises:
            ExcelConnectionError: If connection fails
        """
        try:
            # Use the pre-discovered app and workbook instances
            self._app = workbook_info.app_instance
            self._workbook = workbook_info.workbook_instance

            if self._app is None or self._workbook is None:
                raise ExcelConnectionError(
                    f"Invalid WorkbookInfo: app or workbook instance is None"
                )

            self._connected = True
            logger.info(
                f"Connected to workbook: {self._workbook.name} "
                f"(Excel PID: {workbook_info.excel_pid})"
            )

            # Build workbook model
            return self._build_workbook_model()

        except Exception as e:
            raise ExcelConnectionError(
                f"Failed to connect to workbook {workbook_info.workbook_name}: {e}"
            )

    def disconnect(self) -> None:
        """Disconnect from Excel (does not close Excel)."""
        self._workbook = None
        self._app = None
        self._connected = False
        logger.info("Disconnected from Excel")

    def is_connected(self) -> bool:
        """Check if currently connected."""
        return self._connected and self._workbook is not None

    def _ensure_connected(self) -> None:
        """Ensure we're connected to a workbook."""
        if not self.is_connected():
            raise ExcelConnectionError("Not connected to Excel. Call connect() first.")

    def _build_workbook_model(self) -> Workbook:
        """Build Workbook domain model from xlwings workbook."""
        self._ensure_connected()

        # Get workbook info
        wb_name = self._workbook.name
        wb_path = self._workbook.fullname

        # Get modification time
        try:
            last_modified = datetime.fromtimestamp(Path(wb_path).stat().st_mtime)
        except Exception:
            last_modified = None

        # Get sheets
        sheets = []
        active_sheet_name = None

        try:
            active_sheet_name = self._workbook.sheets.active.name
        except Exception:
            pass

        for xw_sheet in self._workbook.sheets:
            sheet = self._build_sheet_model(xw_sheet)
            sheets.append(sheet)

        return Workbook(
            name=wb_name,
            path=wb_path,
            sheets=sheets,
            active_sheet=active_sheet_name,
            last_modified=last_modified,
        )

    def _build_sheet_model(self, xw_sheet: xw.Sheet) -> Sheet:
        """Build Sheet domain model from xlwings sheet."""
        # Get used range
        used_range = None
        try:
            used_range = xw_sheet.used_range
            if used_range:
                used_range_address = used_range.address
                rows = used_range.shape[0] if used_range.shape else 0
                cols = used_range.shape[1] if used_range.shape else 0
            else:
                used_range_address = None
                rows = 0
                cols = 0
        except Exception as e:
            logger.warning(f"Could not access used range for sheet '{xw_sheet.name}': {e}")
            used_range = None
            used_range_address = None
            rows = 0
            cols = 0

        # Count formulas (approximate - counts cells with formulas in used range)
        formula_count = 0
        logger.debug(f"Sheet '{xw_sheet.name}': Starting formula count - used_range: {used_range_address}, rows: {rows}, cols: {cols}")

        try:
            # Get fresh reference to used_range if needed
            if used_range is None and rows > 0:
                logger.debug(f"Sheet '{xw_sheet.name}': used_range is None, attempting to re-access")
                try:
                    used_range = xw_sheet.used_range
                    logger.debug(f"Sheet '{xw_sheet.name}': Successfully re-accessed used_range")
                except Exception as e:
                    logger.warning(f"Could not re-access used range for formula counting on sheet '{xw_sheet.name}': {e}")

            if used_range:
                # Get all formulas from used range
                logger.debug(f"Sheet '{xw_sheet.name}': Reading formulas from used_range")
                formulas = used_range.formula
                logger.debug(f"Sheet '{xw_sheet.name}': Formula data type: {type(formulas)}, is_list: {isinstance(formulas, list)}, is_none: {formulas is None}")

                if formulas is None:
                    logger.warning(f"Sheet '{xw_sheet.name}': xlwings returned None for used_range.formula")
                elif isinstance(formulas, list):
                    formula_count = sum(
                        1
                        for row in formulas
                        for cell in (row if isinstance(row, list) else [row])
                        if cell and isinstance(cell, str) and cell.startswith("=")
                    )
                    logger.info(f"Sheet '{xw_sheet.name}': Found {formula_count} formulas in used range ({rows}x{cols} cells)")
                elif isinstance(formulas, str) and formulas.startswith("="):
                    formula_count = 1
                    logger.info(f"Sheet '{xw_sheet.name}': Found 1 formula in single cell")
                else:
                    logger.info(f"Sheet '{xw_sheet.name}': No formulas found (data type: {type(formulas)})")
            else:
                logger.debug(f"Sheet '{xw_sheet.name}': used_range is None, skipping formula count")
        except Exception as e:
            logger.warning(
                f"Could not count formulas for sheet '{xw_sheet.name}': {e}. "
                f"Formula count will be 0 (formulas will still be available during graph building)"
            )

        return Sheet(
            name=xw_sheet.name,
            used_range=used_range_address,
            row_count=rows,
            col_count=cols,
            formula_count=formula_count,
        )

    def get_workbook_structure(self) -> WorkbookStructure:
        """
        Get high-level workbook structure.

        Returns:
            WorkbookStructure with workbook and sheet info
        """
        self._ensure_connected()

        workbook = self._build_workbook_model()
        return WorkbookStructure(workbook=workbook, sheet_summaries=workbook.sheets)

    def get_current_selection(self) -> Optional[Selection]:
        """
        Get user's current selection in Excel.

        Returns:
            Selection object or None if no selection
        """
        self._ensure_connected()

        try:
            selection = self._app.selection
            if selection is None:
                return None

            # Get selection address and sheet
            address = selection.address
            sheet_name = selection.sheet.name

            # Check if any cells have formulas
            has_formulas = False
            try:
                formulas = selection.formula
                if isinstance(formulas, str):
                    has_formulas = formulas.startswith("=")
                elif isinstance(formulas, list):
                    has_formulas = any(
                        isinstance(cell, str) and cell.startswith("=")
                        for row in formulas
                        for cell in (row if isinstance(row, list) else [row])
                    )
            except Exception:
                pass

            # Parse address (remove sheet prefix if present)
            if "!" in address:
                address = address.split("!")[1]

            selection_obj = Selection.from_address(address, sheet_name)
            selection_obj.has_formulas = has_formulas

            return selection_obj

        except Exception as e:
            logger.warning(f"Failed to get current selection: {e}")
            return None

    def get_active_sheet(self) -> SheetName:
        """
        Get name of active sheet.

        Returns:
            Active sheet name

        Raises:
            ExcelConnectionError: If not connected
        """
        self._ensure_connected()

        try:
            return self._workbook.sheets.active.name
        except Exception as e:
            raise ExcelConnectionError(f"Failed to get active sheet: {e}")

    def get_cell(self, address: CellAddress, sheet: Optional[SheetName] = None) -> Cell:
        """
        Get single cell data.

        Args:
            address: Cell address (e.g., "A1")
            sheet: Sheet name (if None, uses active sheet)

        Returns:
            Cell domain model

        Raises:
            SheetNotFoundError: If sheet doesn't exist
            InvalidRangeError: If address is invalid
        """
        self._ensure_connected()

        # Get sheet
        if sheet is None:
            sheet = self.get_active_sheet()

        try:
            xw_sheet = self._workbook.sheets[sheet]
        except KeyError:
            raise SheetNotFoundError(f"Sheet '{sheet}' not found")

        try:
            # Get cell
            xw_cell = xw_sheet.range(address)

            # Get value
            value = xw_cell.value

            # Get formula
            formula_text = xw_cell.formula
            formula = Formula(formula_text) if formula_text and formula_text.startswith("=") else None

            # Determine data type
            data_type = self._get_data_type(value)

            return Cell(
                address=address,
                sheet=sheet,
                value=value,
                formula=formula,
                data_type=data_type,
            )

        except Exception as e:
            raise InvalidRangeError(f"Failed to get cell {address}: {e}")

    def get_range_data(
        self, range_obj: Range
    ) -> List[Cell]:
        """
        Get data for a range of cells.

        Args:
            range_obj: Range domain model

        Returns:
            List of Cell domain models

        Raises:
            SheetNotFoundError: If sheet doesn't exist
            InvalidRangeError: If range is invalid
        """
        self._ensure_connected()

        sheet_name = range_obj.sheet
        if sheet_name is None:
            sheet_name = self.get_active_sheet()

        try:
            xw_sheet = self._workbook.sheets[sheet_name]
        except KeyError:
            raise SheetNotFoundError(f"Sheet '{sheet_name}' not found")

        # Get range address without sheet
        range_address = range_obj.to_address(include_sheet=False)
        logger.debug(f"Reading range {sheet_name}!{range_address}")

        # Get xlwings range object
        try:
            xw_range = xw_sheet.range(range_address)
        except Exception as e:
            raise InvalidRangeError(f"Failed to access range {range_address}: {e}")

        # Get values - this should always work
        try:
            values = xw_range.value
            logger.debug(f"Read values from {sheet_name}!{range_address} - type: {type(values)}, is_list: {isinstance(values, list)}")
        except Exception as e:
            raise InvalidRangeError(f"Failed to read values from range {range_address}: {e}")

        # Get formulas - this might fail, so handle separately
        formulas = None
        formula_read_failed = False
        try:
            formulas = xw_range.formula
            logger.debug(f"Read formulas from {sheet_name}!{range_address} - type: {type(formulas)}, is_list: {isinstance(formulas, list)}, is_none: {formulas is None}")
            if formulas is None:
                logger.warning(f"xlwings returned None for formulas in range {sheet_name}!{range_address}")
                formula_read_failed = True
        except Exception as e:
            logger.warning(f"Failed to read formulas from range {sheet_name}!{range_address}: {e}. Cells will be created without formula info.")
            formula_read_failed = True

        # Convert to list of cells
        cells = []

        # Normalize values to 2D array
        if not isinstance(values, list):
            values = [[values]]
        elif values and not isinstance(values[0], list):
            values = [values]

        # Normalize formulas to 2D array (only if we got formulas)
        if not formula_read_failed and formulas is not None:
            if not isinstance(formulas, list):
                formulas = [[formulas]]
            elif formulas and not isinstance(formulas[0], list):
                formulas = [formulas]
        else:
            # Create empty formulas array matching values shape
            formulas = [[None] * len(row) for row in values]
            logger.debug(f"Created empty formulas array for {sheet_name}!{range_address}")

        # Iterate through cells
        formula_count = 0
        for row_idx, row_values in enumerate(values):
            for col_idx, value in enumerate(row_values):
                # Calculate cell address
                row_num = range_obj.start_row + row_idx
                col_num = range_obj.start_col + col_idx
                col_letter = Range._col_index_to_letter(col_num)
                cell_address = f"{col_letter}{row_num}"

                # Get formula for this cell
                formula_text = None
                try:
                    formula_text = formulas[row_idx][col_idx]
                except (IndexError, TypeError) as e:
                    logger.debug(f"Could not access formula at row {row_idx}, col {col_idx}: {e}")

                formula = (
                    Formula(formula_text)
                    if formula_text and isinstance(formula_text, str) and formula_text.startswith("=")
                    else None
                )

                if formula:
                    formula_count += 1
                    logger.debug(f"Found formula in {cell_address}: {formula_text[:50]}..." if len(formula_text) > 50 else f"Found formula in {cell_address}: {formula_text}")

                # Create cell
                cell = Cell(
                    address=cell_address,
                    sheet=sheet_name,
                    value=value,
                    formula=formula,
                    data_type=self._get_data_type(value),
                )
                cells.append(cell)

        logger.info(f"✓ Read {len(cells)} cells from {sheet_name}!{range_address} - {formula_count} cells with formulas")

        if formula_count == 0 and not formula_read_failed:
            logger.info(f"Note: No formulas found in {sheet_name}!{range_address} (range may contain only values)")

        return cells

    @staticmethod
    def _get_data_type(value: Any) -> str:
        """Determine data type of cell value."""
        if value is None:
            return "empty"
        elif isinstance(value, bool):
            return "boolean"
        elif isinstance(value, (int, float)):
            return "number"
        elif isinstance(value, str):
            if value.startswith("#"):
                return "error"
            return "text"
        else:
            return "unknown"

    def __enter__(self) -> "XlwingsConnector":
        """Context manager entry."""
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        """Context manager exit."""
        self.disconnect()
