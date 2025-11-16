"""Storage for annotations."""

import json
from pathlib import Path
from typing import List, Optional

from src.domain.models.annotation import Annotation
from src.domain.models.selection import Range
from src.shared.exceptions import AnnotationError
from src.shared.logging import get_logger
from src.shared.types import SheetName

logger = get_logger(__name__)


class AnnotationStorage:
    """
    Manages storage of annotations to disk.

    Annotations are stored separately from dependency graph
    to ensure they survive graph rebuilds.
    """

    def __init__(self, storage_dir: str = ".cache"):
        """
        Initialize annotation storage.

        Args:
            storage_dir: Directory to store annotation files
        """
        self.storage_dir = Path(storage_dir)
        self.storage_dir.mkdir(parents=True, exist_ok=True)

    def get_storage_path(self, workbook_path: str) -> Path:
        """
        Get storage file path for a workbook's annotations.

        Args:
            workbook_path: Path to Excel workbook

        Returns:
            Path to annotations file
        """
        # Create safe filename from workbook path
        workbook_name = Path(workbook_path).stem
        safe_name = "".join(c if c.isalnum() else "_" for c in workbook_name)
        return self.storage_dir / f"{safe_name}_annotations.json"

    def save(self, annotations: List[Annotation], workbook_path: str) -> None:
        """
        Save annotations to disk.

        Args:
            annotations: List of annotations to save
            workbook_path: Path to Excel workbook

        Raises:
            AnnotationError: If save fails
        """
        try:
            storage_path = self.get_storage_path(workbook_path)

            # Serialize annotations
            data = {
                "workbook_path": str(workbook_path),
                "annotations": [ann.to_dict() for ann in annotations],
            }

            # Save to file
            with open(storage_path, "w") as f:
                json.dump(data, f, indent=2)

            logger.info(
                f"Saved {len(annotations)} annotations to: {storage_path}"
            )

        except Exception as e:
            raise AnnotationError(f"Failed to save annotations: {e}")

    def load(self, workbook_path: str) -> List[Annotation]:
        """
        Load annotations from disk.

        Args:
            workbook_path: Path to Excel workbook

        Returns:
            List of annotations (empty list if none found)

        Raises:
            AnnotationError: If load fails unexpectedly
        """
        try:
            storage_path = self.get_storage_path(workbook_path)

            if not storage_path.exists():
                logger.debug(f"No annotations found for {workbook_path}")
                return []

            # Load data
            with open(storage_path, "r") as f:
                data = json.load(f)

            # Deserialize annotations
            annotations = [
                Annotation.from_dict(ann_data)
                for ann_data in data.get("annotations", [])
            ]

            logger.info(
                f"Loaded {len(annotations)} annotations from: {storage_path}"
            )

            return annotations

        except FileNotFoundError:
            return []
        except json.JSONDecodeError as e:
            logger.warning(f"Invalid annotations file format: {e}")
            return []
        except Exception as e:
            raise AnnotationError(f"Failed to load annotations: {e}")

    def add(
        self,
        annotation: Annotation,
        workbook_path: str,
    ) -> None:
        """
        Add a new annotation.

        Args:
            annotation: Annotation to add
            workbook_path: Path to Excel workbook

        Raises:
            AnnotationError: If operation fails
        """
        # Load existing
        annotations = self.load(workbook_path)

        # Add new annotation
        annotations.append(annotation)

        # Save back
        self.save(annotations, workbook_path)

    def remove(
        self,
        range_obj: Range,
        workbook_path: str,
    ) -> bool:
        """
        Remove annotation for a specific range.

        Args:
            range_obj: Range to remove annotation from
            workbook_path: Path to Excel workbook

        Returns:
            True if annotation was removed, False if not found

        Raises:
            AnnotationError: If operation fails
        """
        # Load existing
        annotations = self.load(workbook_path)

        # Find and remove matching annotation
        initial_count = len(annotations)
        annotations = [
            ann
            for ann in annotations
            if not ann.matches_range(range_obj, strict=True)
        ]

        if len(annotations) == initial_count:
            return False

        # Save back
        self.save(annotations, workbook_path)
        return True

    def get_for_sheet(
        self,
        sheet_name: SheetName,
        workbook_path: str,
    ) -> List[Annotation]:
        """
        Get all annotations for a specific sheet.

        Args:
            sheet_name: Name of sheet
            workbook_path: Path to Excel workbook

        Returns:
            List of annotations for the sheet
        """
        all_annotations = self.load(workbook_path)

        return [
            ann
            for ann in all_annotations
            if ann.sheet == sheet_name
        ]

    def search(
        self,
        query: str,
        workbook_path: str,
    ) -> List[Annotation]:
        """
        Search annotations by label or description.

        Args:
            query: Search query
            workbook_path: Path to Excel workbook

        Returns:
            List of matching annotations
        """
        all_annotations = self.load(workbook_path)
        query_lower = query.lower()

        return [
            ann
            for ann in all_annotations
            if query_lower in ann.label.lower()
            or (ann.description and query_lower in ann.description.lower())
        ]

    def clear(self, workbook_path: str) -> None:
        """
        Clear all annotations for a workbook.

        Args:
            workbook_path: Path to Excel workbook
        """
        try:
            storage_path = self.get_storage_path(workbook_path)

            if storage_path.exists():
                storage_path.unlink()
                logger.info(f"Cleared annotations: {storage_path}")

        except Exception as e:
            logger.warning(f"Failed to clear annotations: {e}")

    def exists(self, workbook_path: str) -> bool:
        """
        Check if annotations exist for a workbook.

        Args:
            workbook_path: Path to Excel workbook

        Returns:
            True if annotations file exists
        """
        storage_path = self.get_storage_path(workbook_path)
        return storage_path.exists()
