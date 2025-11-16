"""Domain models for annotations."""

from dataclasses import dataclass
from datetime import datetime
from typing import Optional

from src.domain.models.selection import Range
from src.shared.types import RangeAddress, SheetName


@dataclass
class Annotation:
    """
    Represents a semantic annotation on a range of cells.

    Annotations add business context to Excel ranges,
    helping the LLM provide better explanations.

    Examples:
        - Range C5:C100 labeled as "Position Deltas"
        - Cell B5 labeled as "VaR 95% Confidence"
    """

    range: Range
    label: str
    description: Optional[str] = None
    created_at: Optional[datetime] = None
    metadata: Optional[dict] = None

    def __post_init__(self) -> None:
        """Initialize created timestamp if not provided."""
        if self.created_at is None:
            self.created_at = datetime.now()

    @classmethod
    def from_address(
        cls,
        address: RangeAddress,
        label: str,
        description: Optional[str] = None,
        metadata: Optional[dict] = None,
    ) -> "Annotation":
        """
        Create annotation from address string.

        Args:
            address: Excel address (e.g., "Sheet1!A1:B10")
            label: Short label for the range
            description: Optional detailed description
            metadata: Optional additional metadata

        Returns:
            Annotation object
        """
        range_obj = Range.from_address(address)
        return cls(
            range=range_obj,
            label=label,
            description=description,
            metadata=metadata,
        )

    @property
    def sheet(self) -> Optional[SheetName]:
        """Get sheet name from range."""
        return self.range.sheet

    @property
    def address(self) -> RangeAddress:
        """Get range address as string."""
        return self.range.to_address()

    def matches_range(self, other_range: Range, strict: bool = False) -> bool:
        """
        Check if this annotation matches a given range.

        Args:
            other_range: Range to check against
            strict: If True, requires exact match. If False, allows overlap.

        Returns:
            True if ranges match (exact or overlapping)
        """
        if strict:
            # Exact match
            return (
                self.range.sheet == other_range.sheet
                and self.range.start_row == other_range.start_row
                and self.range.end_row == other_range.end_row
                and self.range.start_col == other_range.start_col
                and self.range.end_col == other_range.end_col
            )
        else:
            # Allow overlap
            return self.range.overlaps(other_range)

    def contains_range(self, other_range: Range) -> bool:
        """
        Check if this annotation's range contains another range.

        Args:
            other_range: Range to check

        Returns:
            True if annotation range contains the other range
        """
        return self.range.contains(other_range)

    def to_dict(self) -> dict:
        """
        Convert annotation to dictionary for serialization.

        Returns:
            Dictionary representation
        """
        return {
            "range": self.address,
            "label": self.label,
            "description": self.description,
            "created_at": self.created_at.isoformat() if self.created_at else None,
            "metadata": self.metadata,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "Annotation":
        """
        Create annotation from dictionary.

        Args:
            data: Dictionary with annotation data

        Returns:
            Annotation object
        """
        created_at = None
        if data.get("created_at"):
            created_at = datetime.fromisoformat(data["created_at"])

        return cls.from_address(
            address=data["range"],
            label=data["label"],
            description=data.get("description"),
            metadata=data.get("metadata"),
        )

    def __str__(self) -> str:
        """String representation."""
        desc = f" - {self.description}" if self.description else ""
        return f"'{self.label}' @ {self.address}{desc}"

    def __repr__(self) -> str:
        """Debug representation."""
        return f"Annotation(range='{self.address}', label='{self.label}')"
