"""Workbook discovery service for listing open Excel workbooks."""

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import List, Optional

import xlwings as xw

from src.shared.exceptions import ExcelConnectionError


@dataclass
class WorkbookInfo:
    """Information about an open Excel workbook."""

    excel_pid: int  # Process ID of Excel instance
    workbook_name: str  # Filename (e.g., "Model.xlsx")
    full_path: str  # Full path to file
    sheet_count: int  # Number of sheets
    is_saved: bool  # Whether workbook has unsaved changes
    modified_time: Optional[datetime] = None  # Last modification time
    app_instance: Optional[any] = None  # xlwings app instance (not serialised)
    workbook_instance: Optional[any] = None  # xlwings workbook instance (not serialised)

    def __str__(self) -> str:
        """String representation."""
        saved_indicator = "" if self.is_saved else " *"
        return f"[PID {self.excel_pid}] {self.workbook_name}{saved_indicator} - {self.full_path}"

    def display_name(self) -> str:
        """Short display name for UI."""
        return f"{self.workbook_name} (PID: {self.excel_pid})"


class WorkbookDiscovery:
    """Discover and list open Excel workbooks across all instances."""

    @staticmethod
    def list_all_workbooks() -> List[WorkbookInfo]:
        """
        List all open Excel workbooks across all Excel instances.

        Returns:
            List of WorkbookInfo objects

        Raises:
            ExcelConnectionError: If no Excel instances are running
        """
        workbooks: List[WorkbookInfo] = []

        try:
            # Get all Excel application instances
            apps = xw.apps

            if not apps:
                raise ExcelConnectionError("No Excel instances found. Please open Excel first.")

            # Iterate through all app instances
            for app in apps:
                try:
                    pid = app.pid

                    # Iterate through all workbooks in this app
                    for book in app.books:
                        try:
                            # Get workbook details
                            info = WorkbookInfo(
                                excel_pid=pid,
                                workbook_name=book.name,
                                full_path=book.fullname if book.fullname else book.name,
                                sheet_count=len(book.sheets),
                                is_saved=book.saved,
                                app_instance=app,
                                workbook_instance=book,
                            )

                            # Try to get modification time (file must be saved)
                            if book.fullname and Path(book.fullname).exists():
                                info.modified_time = datetime.fromtimestamp(
                                    Path(book.fullname).stat().st_mtime
                                )

                            workbooks.append(info)

                        except Exception as e:
                            # Skip this workbook if we can't read it
                            continue

                except Exception as e:
                    # Skip this app instance if we can't read it
                    continue

            if not workbooks:
                raise ExcelConnectionError(
                    "No open workbooks found. Please open an Excel workbook first."
                )

            return workbooks

        except Exception as e:
            if isinstance(e, ExcelConnectionError):
                raise
            raise ExcelConnectionError(f"Failed to discover workbooks: {e}")

    @staticmethod
    def find_by_path(full_path: str) -> List[WorkbookInfo]:
        """
        Find workbooks by full path.

        Note: May return multiple WorkbookInfo objects if the same file
        is open in multiple Excel instances.

        Args:
            full_path: Full path to workbook file

        Returns:
            List of matching WorkbookInfo objects (may be empty)
        """
        all_workbooks = WorkbookDiscovery.list_all_workbooks()

        # Normalize paths for comparison
        target_path = Path(full_path).resolve()

        matches = []
        for wb_info in all_workbooks:
            try:
                wb_path = Path(wb_info.full_path).resolve()
                if wb_path == target_path:
                    matches.append(wb_info)
            except Exception:
                # Skip if path can't be resolved
                continue

        return matches

    @staticmethod
    def find_by_name(workbook_name: str) -> List[WorkbookInfo]:
        """
        Find workbooks by filename (not full path).

        May return multiple matches if files with same name exist in different locations.

        Args:
            workbook_name: Workbook filename (e.g., "Model.xlsx")

        Returns:
            List of matching WorkbookInfo objects (may be empty)
        """
        all_workbooks = WorkbookDiscovery.list_all_workbooks()

        matches = [
            wb_info
            for wb_info in all_workbooks
            if wb_info.workbook_name.lower() == workbook_name.lower()
        ]

        return matches

    @staticmethod
    def group_duplicates(
        workbooks: List[WorkbookInfo],
    ) -> dict[str, List[WorkbookInfo]]:
        """
        Group workbooks by full path to identify duplicates across Excel instances.

        Args:
            workbooks: List of WorkbookInfo objects

        Returns:
            Dictionary mapping full_path -> list of WorkbookInfo objects
        """
        groups: dict[str, List[WorkbookInfo]] = {}

        for wb_info in workbooks:
            path = wb_info.full_path
            if path not in groups:
                groups[path] = []
            groups[path].append(wb_info)

        return groups

    @staticmethod
    def has_duplicates(workbooks: List[WorkbookInfo]) -> bool:
        """
        Check if any workbook appears in multiple Excel instances.

        Args:
            workbooks: List of WorkbookInfo objects

        Returns:
            True if duplicates exist
        """
        groups = WorkbookDiscovery.group_duplicates(workbooks)
        return any(len(group) > 1 for group in groups.values())

    @staticmethod
    def get_duplicate_paths(workbooks: List[WorkbookInfo]) -> List[str]:
        """
        Get list of file paths that are open in multiple Excel instances.

        Args:
            workbooks: List of WorkbookInfo objects

        Returns:
            List of file paths that appear multiple times
        """
        groups = WorkbookDiscovery.group_duplicates(workbooks)
        return [path for path, group in groups.items() if len(group) > 1]
