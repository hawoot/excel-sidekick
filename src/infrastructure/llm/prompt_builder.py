"""Prompt builder for constructing LLM prompts."""

from typing import List, Optional

from src.domain.models.annotation import Annotation
from src.domain.models.dependency import DependencyTree
from src.domain.models.query import LLMContext
from src.domain.models.selection import Selection


class PromptBuilder:
    """
    Builds prompts for LLM queries.

    Assembles context from various sources into a structured prompt.
    """

    @staticmethod
    def build_context(
        question: str,
        selection: Optional[Selection] = None,
        formulas: Optional[List[str]] = None,
        dependency_tree: Optional[DependencyTree] = None,
        annotations: Optional[List[Annotation]] = None,
        spatial_context: Optional[str] = None,
        mode: str = "educational",
    ) -> LLMContext:
        """
        Build LLM context from components.

        Args:
            question: User's question
            selection: User's selection (if any)
            formulas: List of formula strings
            dependency_tree: Dependency tree
            annotations: Relevant annotations
            spatial_context: Snapshot or region info
            mode: Query mode (educational, technical, concise)

        Returns:
            LLM context ready for prompting
        """
        context = LLMContext(question=question)

        # Add selection info
        if selection:
            context.selection_info = (
                f"User selected: {selection.to_address()} "
                f"({selection.range.cell_count()} cells)"
            )

        # Add formulas
        if formulas:
            context.formulas = formulas

        # Add dependency tree
        if dependency_tree:
            context.dependencies = PromptBuilder._format_dependency_tree(dependency_tree)

        # Add annotations
        if annotations:
            context.annotations = [str(ann) for ann in annotations]

        # Add spatial context
        if spatial_context:
            context.spatial_context = spatial_context

        # Add mode context
        context.additional_context["mode"] = mode

        return context

    @staticmethod
    def _format_dependency_tree(tree: DependencyTree) -> str:
        """
        Format dependency tree for prompt.

        Args:
            tree: Dependency tree to format

        Returns:
            Formatted string representation
        """
        lines = [
            f"Dependency Tree ({tree.direction}, depth={tree.max_depth}):",
            "",
        ]

        # Add tree lines
        lines.extend(tree.to_lines())

        return "\n".join(lines)

    @staticmethod
    def get_system_prompt(mode: str = "educational") -> str:
        """
        Get system prompt based on mode.

        Args:
            mode: Query mode

        Returns:
            System prompt string
        """
        prompts = {
            "educational": """You are an expert financial analyst and Excel specialist helping someone understand complex risk management spreadsheets.

Your role:
- Explain calculations in terms of financial concepts and business logic
- Focus on the "why" and "what" rather than just the "how"
- Use clear, educational language
- Connect Excel formulas to their business meaning
- Explain financial concepts when relevant (VaR, Greeks, P&L, etc.)

Guidelines:
- Start with the high-level purpose before diving into details
- Explain dependencies and how data flows through calculations
- Highlight key assumptions or important aspects
- Use examples when helpful
- Be thorough but accessible""",
            "technical": """You are an Excel and financial engineering expert providing technical analysis.

Your role:
- Provide precise technical explanations of formulas and calculations
- Explain mathematical relationships and dependencies
- Highlight formula structure and logic
- Note any technical issues or edge cases

Guidelines:
- Be precise and technical
- Focus on accuracy and completeness
- Explain formula syntax and functions used
- Highlight dependencies and data flow""",
            "concise": """You are an expert providing quick, direct answers about Excel calculations.

Your role:
- Provide brief, accurate answers
- Focus on the essentials
- Be direct and clear

Guidelines:
- Keep responses short
- Answer the specific question asked
- Avoid unnecessary detail""",
        }

        return prompts.get(mode, prompts["educational"])
