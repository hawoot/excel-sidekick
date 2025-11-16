"""Excel Assistant Service - Main application service."""

from typing import Optional

from src.domain.models.query import AssistantResponse, QuestionContext
from src.domain.models.selection import Selection
from src.domain.models.workbook import Workbook
from src.domain.services.annotation_management_service import AnnotationManagementService
from src.domain.services.dependency_analysis_service import DependencyAnalysisService
from src.domain.services.exploration_agent import ExplorationAgent
from src.domain.services.llm_interaction_service import LLMInteractionService
from src.domain.services.workbook_data_service import WorkbookDataService
from src.infrastructure.config.config_loader import Config
from src.shared.logging import get_logger, setup_logging

logger = get_logger(__name__)


class ExcelAssistantService:
    """
    Main application service for Excel Sidekick.

    Thin facade that coordinates all domain services.
    Used by both CLI and future Web API.
    """

    def __init__(self, config: Config):
        """
        Initialize Excel Assistant Service.

        Args:
            config: Application configuration
        """
        self.config = config

        # Set up logging
        setup_logging(
            level=config.logging.level.value,
            log_file=config.logging.file,
            console=config.logging.console,
            log_format=config.logging.format,
        )

        logger.info("Initializing Excel Assistant Service")

        # Initialize domain services
        self.workbook_data = WorkbookDataService(config)
        self.dependency_analysis = DependencyAnalysisService(
            config, self.workbook_data
        )
        self.annotation_management = AnnotationManagementService(config)
        self.llm_interaction = LLMInteractionService(config)

        # Initialize exploration agent
        self.agent = ExplorationAgent(
            config=config,
            workbook_data=self.workbook_data,
            dependency_analysis=self.dependency_analysis,
            annotation_management=self.annotation_management,
            llm_interaction=self.llm_interaction,
        )

        self._current_workbook: Optional[Workbook] = None

        logger.info("Excel Assistant Service initialized")

    def connect(self, workbook_name: Optional[str] = None) -> Workbook:
        """
        Connect to Excel workbook.

        Args:
            workbook_name: Name of workbook (None for active)

        Returns:
            Workbook model
        """
        logger.info(f"Connecting to workbook: {workbook_name or 'active'}")

        # Connect to workbook
        workbook = self.workbook_data.connect(workbook_name)
        self._current_workbook = workbook

        # Set workbook for annotation management
        self.annotation_management.set_workbook(workbook.path)

        # Build dependency graph
        logger.info("Building dependency graph...")
        self.dependency_analysis.build_graph(workbook)

        logger.info(
            f"Connected to '{workbook.name}' "
            f"({len(workbook.sheets)} sheets, "
            f"{workbook.total_formula_count()} formulas)"
        )

        return workbook

    def disconnect(self) -> None:
        """Disconnect from current workbook."""
        self.workbook_data.disconnect()
        self._current_workbook = None
        logger.info("Disconnected from workbook")

    def ask_question(
        self,
        question: str,
        selection: Optional[Selection] = None,
        mode: str = "educational",
    ) -> AssistantResponse:
        """
        Ask a question about the workbook.

        Args:
            question: User's question
            selection: Optional specific selection (None for current selection)
            mode: Query mode (educational, technical, concise)

        Returns:
            Assistant response with answer and context
        """
        logger.info(f"Question: {question}")

        context = QuestionContext(
            question=question,
            selection=selection,
            active_sheet=self.workbook_data.get_active_sheet() if self.workbook_data.is_connected() else None,
            mode=mode,
        )

        response = self.agent.explore_and_answer(context)

        return response

    def explain_selection(
        self,
        selection: Optional[Selection] = None,
        mode: str = "educational",
    ) -> AssistantResponse:
        """
        Explain the current or specified selection.

        Args:
            selection: Selection to explain (None for current)
            mode: Query mode

        Returns:
            Assistant response
        """
        if selection is None:
            selection = self.workbook_data.get_current_selection()

        if selection is None:
            # No selection - return helpful message
            return AssistantResponse(
                question="Explain selection",
                answer="No selection found. Please select a range in Excel and try again.",
            )

        question = f"Explain what {selection.to_address()} calculates and why."

        return self.ask_question(question, selection, mode)

    def add_annotation(
        self,
        range_address: str,
        label: str,
        description: Optional[str] = None,
    ) -> None:
        """
        Add semantic annotation to a range.

        Args:
            range_address: Range address (e.g., "Sheet1!A1:B10")
            label: Short label
            description: Optional detailed description
        """
        self.annotation_management.add_annotation(
            range_address=range_address,
            label=label,
            description=description,
        )

        logger.info(f"Added annotation: '{label}' for {range_address}")

    def get_annotations(self, sheet: Optional[str] = None) -> list:
        """
        Get annotations for a sheet or all sheets.

        Args:
            sheet: Sheet name (None for all)

        Returns:
            List of annotations
        """
        return self.annotation_management.get_annotations(sheet)

    def rebuild_cache(self) -> None:
        """Rebuild dependency graph cache."""
        if self._current_workbook is None:
            raise ValueError("No workbook connected")

        logger.info("Rebuilding dependency graph...")
        self.dependency_analysis.rebuild_graph(self._current_workbook)
        logger.info("Graph rebuilt")

    def clear_cache(self) -> None:
        """Clear dependency graph cache."""
        if self._current_workbook is None:
            raise ValueError("No workbook connected")

        self.dependency_analysis.clear_cache(self._current_workbook.path)
        logger.info("Cache cleared")

    def get_cache_status(self) -> dict:
        """
        Get cache status information.

        Returns:
            Dict with cache status
        """
        if self._current_workbook is None:
            return {"connected": False}

        graph = self.dependency_analysis.get_current_graph()

        return {
            "connected": True,
            "workbook": self._current_workbook.name,
            "graph_cached": graph is not None,
            "node_count": graph.node_count() if graph else 0,
            "formula_count": graph.formula_count() if graph else 0,
            "has_annotations": self.annotation_management.has_annotations(),
        }

    def is_connected(self) -> bool:
        """Check if connected to a workbook."""
        return self.workbook_data.is_connected()

    def get_current_workbook(self) -> Optional[Workbook]:
        """Get current workbook."""
        return self._current_workbook
