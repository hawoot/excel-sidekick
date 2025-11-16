"""Domain models for queries and responses."""

from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Dict, List, Optional

from src.domain.models.annotation import Annotation
from src.domain.models.dependency import DependencyTree
from src.domain.models.selection import Selection


@dataclass
class QuestionContext:
    """
    Context for a user question.

    Provides all relevant context to help the agent answer the question.
    """

    question: str
    selection: Optional[Selection] = None
    active_sheet: Optional[str] = None
    mode: str = "educational"  # educational | technical | concise
    include_dependencies: bool = True
    include_annotations: bool = True
    max_depth: int = 3

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary."""
        return {
            "question": self.question,
            "selection": str(self.selection) if self.selection else None,
            "active_sheet": self.active_sheet,
            "mode": self.mode,
            "include_dependencies": self.include_dependencies,
            "include_annotations": self.include_annotations,
            "max_depth": self.max_depth,
        }


@dataclass
class ExplorationStep:
    """
    Represents a single step in the agent's exploration.

    Useful for debugging and showing the agent's thinking process.
    """

    step_number: int
    tool_name: str
    tool_input: Dict[str, Any]
    tool_output: Any
    timestamp: datetime = field(default_factory=datetime.now)

    def __str__(self) -> str:
        """String representation."""
        return f"Step {self.step_number}: {self.tool_name}({self.tool_input})"


@dataclass
class AssistantResponse:
    """
    Response from the Excel assistant.

    Contains the answer plus context about how it was derived.
    """

    question: str
    answer: str
    context_used: Dict[str, Any] = field(default_factory=dict)
    dependencies_traced: Optional[DependencyTree] = None
    annotations_found: List[Annotation] = field(default_factory=list)
    exploration_steps: List[ExplorationStep] = field(default_factory=list)
    timestamp: datetime = field(default_factory=datetime.now)
    metadata: Dict[str, Any] = field(default_factory=dict)

    def add_exploration_step(
        self,
        tool_name: str,
        tool_input: Dict[str, Any],
        tool_output: Any,
    ) -> None:
        """Add an exploration step."""
        step = ExplorationStep(
            step_number=len(self.exploration_steps) + 1,
            tool_name=tool_name,
            tool_input=tool_input,
            tool_output=tool_output,
        )
        self.exploration_steps.append(step)

    def add_context(self, key: str, value: Any) -> None:
        """Add context information."""
        self.context_used[key] = value

    def summary(self) -> str:
        """Get a summary of the response."""
        lines = [
            f"Question: {self.question}",
            f"Answer: {self.answer}",
            f"",
            f"Context:",
            f"  - Exploration steps: {len(self.exploration_steps)}",
            f"  - Dependencies traced: {self.dependencies_traced is not None}",
            f"  - Annotations found: {len(self.annotations_found)}",
        ]
        return "\n".join(lines)

    def __str__(self) -> str:
        """String representation."""
        return self.answer


@dataclass
class LLMContext:
    """
    Context assembled for LLM query.

    Contains all information needed for the LLM to answer the question.
    """

    question: str
    selection_info: Optional[str] = None
    formulas: List[str] = field(default_factory=list)
    dependencies: Optional[str] = None  # Formatted dependency tree
    annotations: List[str] = field(default_factory=list)
    spatial_context: Optional[str] = None  # Snapshot or region info
    additional_context: Dict[str, Any] = field(default_factory=dict)

    def to_prompt(self, system_prompt: Optional[str] = None) -> str:
        """
        Convert context to a formatted prompt for the LLM.

        Args:
            system_prompt: Optional system prompt to include

        Returns:
            Formatted prompt string
        """
        parts = []

        if system_prompt:
            parts.append(f"# System\n{system_prompt}\n")

        parts.append("# Context from Excel\n")

        if self.selection_info:
            parts.append(f"## User Selection\n{self.selection_info}\n")

        if self.spatial_context:
            parts.append(f"## Spatial Context\n{self.spatial_context}\n")

        if self.formulas:
            parts.append("## Formulas\n")
            for formula in self.formulas:
                parts.append(f"- {formula}")
            parts.append("")

        if self.dependencies:
            parts.append(f"## Dependencies\n{self.dependencies}\n")

        if self.annotations:
            parts.append("## Annotations\n")
            for annotation in self.annotations:
                parts.append(f"- {annotation}")
            parts.append("")

        if self.additional_context:
            parts.append("## Additional Context\n")
            for key, value in self.additional_context.items():
                parts.append(f"**{key}**: {value}")
            parts.append("")

        parts.append(f"# Question\n{self.question}\n")
        parts.append("# Answer\nPlease explain this in terms of financial concepts and business logic:")

        return "\n".join(parts)

    def token_estimate(self) -> int:
        """
        Rough estimate of token count.

        Returns:
            Estimated token count (assuming ~4 chars per token)
        """
        prompt = self.to_prompt()
        return len(prompt) // 4


@dataclass
class LLMResponse:
    """
    Response from LLM provider.

    Wraps the raw LLM response with metadata.
    """

    content: str
    provider: str
    model: Optional[str] = None
    tokens_used: Optional[int] = None
    timestamp: datetime = field(default_factory=datetime.now)
    metadata: Dict[str, Any] = field(default_factory=dict)

    def __str__(self) -> str:
        """String representation."""
        return self.content
