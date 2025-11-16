"""Annotation management service."""

from typing import List, Optional

from src.domain.models.annotation import Annotation
from src.domain.models.selection import Range
from src.infrastructure.config.config_loader import Config
from src.infrastructure.storage.annotation_storage import AnnotationStorage
from src.shared.logging import get_logger
from src.shared.types import SheetName

logger = get_logger(__name__)


class AnnotationManagementService:
    """
    Service for managing semantic annotations on Excel ranges.

    Provides high-level methods for adding, retrieving, and searching annotations.
    """

    def __init__(self, config: Config):
        """
        Initialize annotation management service.

        Args:
            config: Application configuration
        """
        self.config = config
        self.storage = AnnotationStorage(config.annotations.file_location)
        self._current_workbook_path: Optional[str] = None

    def set_workbook(self, workbook_path: str) -> None:
        """
        Set the current workbook path for annotation operations.

        Args:
            workbook_path: Path to Excel workbook
        """
        self._current_workbook_path = workbook_path
        logger.debug(f"Set current workbook for annotations: {workbook_path}")

    def add_annotation(
        self,
        range_address: str,
        label: str,
        description: Optional[str] = None,
    ) -> Annotation:
        """
        Add a new annotation.

        Args:
            range_address: Range address (e.g., "Sheet1!A1:B10")
            label: Short label for the range
            description: Optional detailed description

        Returns:
            Created annotation

        Raises:
            AnnotationError: If operation fails
        """
        self._ensure_workbook_set()

        annotation = Annotation.from_address(
            address=range_address,
            label=label,
            description=description,
        )

        self.storage.add(annotation, self._current_workbook_path)

        logger.info(f"Added annotation: '{label}' for {range_address}")

        return annotation

    def get_annotations(
        self,
        sheet: Optional[SheetName] = None,
    ) -> List[Annotation]:
        """
        Get annotations, optionally filtered by sheet.

        Args:
            sheet: Sheet name to filter by (None for all sheets)

        Returns:
            List of annotations
        """
        self._ensure_workbook_set()

        if sheet:
            annotations = self.storage.get_for_sheet(sheet, self._current_workbook_path)
            logger.debug(f"Retrieved {len(annotations)} annotations for sheet '{sheet}'")
        else:
            annotations = self.storage.load(self._current_workbook_path)
            logger.debug(f"Retrieved {len(annotations)} annotations")

        return annotations

    def search_annotations(self, query: str) -> List[Annotation]:
        """
        Search annotations by label or description.

        Args:
            query: Search query

        Returns:
            List of matching annotations
        """
        self._ensure_workbook_set()

        annotations = self.storage.search(query, self._current_workbook_path)

        logger.debug(f"Found {len(annotations)} annotations matching '{query}'")

        return annotations

    def remove_annotation(self, range_address: str) -> bool:
        """
        Remove annotation for a specific range.

        Args:
            range_address: Range address (e.g., "Sheet1!A1:B10")

        Returns:
            True if annotation was removed
        """
        self._ensure_workbook_path()

        range_obj = Range.from_address(range_address)
        removed = self.storage.remove(range_obj, self._current_workbook_path)

        if removed:
            logger.info(f"Removed annotation for {range_address}")
        else:
            logger.debug(f"No annotation found for {range_address}")

        return removed

    def get_annotations_for_range(
        self,
        range_obj: Range,
        exact_match: bool = False,
    ) -> List[Annotation]:
        """
        Get annotations that match or overlap with a range.

        Args:
            range_obj: Range to check
            exact_match: If True, only exact matches returned

        Returns:
            List of matching annotations
        """
        self._ensure_workbook_set()

        # Get all annotations for the sheet
        all_annotations = self.storage.get_for_sheet(
            range_obj.sheet,
            self._current_workbook_path,
        )

        # Filter by range
        matching = [
            ann
            for ann in all_annotations
            if ann.matches_range(range_obj, strict=exact_match)
        ]

        return matching

    def clear_all(self) -> None:
        """Clear all annotations for current workbook."""
        self._ensure_workbook_set()

        self.storage.clear(self._current_workbook_path)
        logger.info("Cleared all annotations")

    def has_annotations(self) -> bool:
        """
        Check if workbook has any annotations.

        Returns:
            True if annotations exist
        """
        self._ensure_workbook_set()

        return self.storage.exists(self._current_workbook_path)

    def _ensure_workbook_set(self) -> None:
        """Ensure workbook path is set."""
        if self._current_workbook_path is None:
            raise ValueError("Workbook path not set. Call set_workbook() first.")
