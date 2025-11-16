"""Agent state definition for LangGraph."""

from typing import Any, Dict, List, Optional, TypedDict

from src.domain.models.selection import Selection


class AgentState(TypedDict):
    """
    State for the exploration agent.

    LangGraph uses this state to track the agent's progress.
    """

    # Input
    question: str
    selection: Optional[Selection]
    active_sheet: Optional[str]
    mode: str  # educational | technical | concise

    # Agent execution
    messages: List[Dict[str, Any]]  # LangGraph messages
    iteration: int  # Current iteration count

    # Gathered context
    snapshot: Optional[str]  # Snapshot of selected/explored region
    formulas: List[str]  # Formulas found
    dependencies: Optional[str]  # Dependency tree (formatted)
    annotations: List[str]  # Annotations found

    # Output
    answer: Optional[str]  # Final answer from LLM
    exploration_steps: List[Dict[str, Any]]  # Tool calls made
