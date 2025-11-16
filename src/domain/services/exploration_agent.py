"""Exploration agent for intelligently analyzing Excel workbooks."""

from typing import List, Optional

from src.domain.models.annotation import Annotation
from src.domain.models.query import AssistantResponse, QuestionContext
from src.domain.models.selection import Range, Selection
from src.domain.services.annotation_management_service import AnnotationManagementService
from src.domain.services.dependency_analysis_service import DependencyAnalysisService
from src.domain.services.llm_interaction_service import LLMInteractionService
from src.domain.services.workbook_data_service import WorkbookDataService
from src.infrastructure.config.config_loader import Config
from src.shared.logging import get_logger
from src.shared.types import TraceDirection

logger = get_logger(__name__)


class ExplorationAgent:
    """
    Agent that intelligently explores Excel workbooks to answer questions.

    Uses domain services as tools to gather context, then queries LLM.
    """

    def __init__(
        self,
        config: Config,
        workbook_data: WorkbookDataService,
        dependency_analysis: DependencyAnalysisService,
        annotation_management: AnnotationManagementService,
        llm_interaction: LLMInteractionService,
    ):
        """
        Initialize exploration agent.

        Args:
            config: Application configuration
            workbook_data: Service for reading Excel data
            dependency_analysis: Service for dependency analysis
            annotation_management: Service for annotations
            llm_interaction: Service for LLM queries
        """
        self.config = config
        self.workbook_data = workbook_data
        self.dependency_analysis = dependency_analysis
        self.annotation_management = annotation_management
        self.llm_interaction = llm_interaction

    def explore_and_answer(self, context: QuestionContext) -> AssistantResponse:
        """
        Explore workbook and answer question.

        Args:
            context: Question context with question and optional selection

        Returns:
            Assistant response with answer and context
        """
        logger.info(f"Starting exploration: {context.question}")

        response = AssistantResponse(
            question=context.question,
            answer="",  # Will be filled by LLM
        )

        # Step 1: Get starting point (selection or active sheet)
        selection = context.selection or self.workbook_data.get_current_selection()

        if selection:
            logger.info(f"Working from selection: {selection}")
            response.add_context("selection", str(selection))
            self._explore_from_selection(selection, context, response)
        else:
            logger.info("No selection, using active sheet")
            active_sheet = self.workbook_data.get_active_sheet()
            response.add_context("active_sheet", active_sheet)
            self._explore_active_sheet(active_sheet, context, response)

        # Step 2: Query LLM with gathered context
        self._query_llm(context, response)

        logger.info("Exploration complete")

        return response

    def _explore_from_selection(
        self,
        selection: Selection,
        context: QuestionContext,
        response: AssistantResponse,
    ) -> None:
        """
        Explore starting from user's selection.

        Args:
            selection: User's selection
            context: Question context
            response: Response to populate
        """
        # Get snapshot of selection + surrounding context
        logger.debug("Getting snapshot of selection")
        expanded_range = self.workbook_data.expand_selection_context(selection)

        snapshot = self.workbook_data.get_snapshot(
            sheet=selection.sheet_name,
            range_address=expanded_range.to_address(include_sheet=False),
        )

        response.add_context("snapshot", snapshot)

        # Get formulas in selection
        if selection.has_formulas:
            logger.debug("Getting formulas from selection")
            cells = self.workbook_data.get_range_data(selection.range)
            formulas = [str(cell.formula) for cell in cells if cell.has_formula()]
            response.add_context("formulas", formulas)

            # Trace dependencies if requested
            if context.include_dependencies and formulas:
                logger.debug("Tracing dependencies")
                # Pick first formula cell to trace from
                formula_cells = [cell for cell in cells if cell.has_formula()]
                if formula_cells:
                    first_cell = formula_cells[0]
                    dep_tree = self.dependency_analysis.trace_dependencies(
                        cell_address=first_cell.full_address,
                        direction=TraceDirection.BOTH,
                        depth=context.max_depth,
                    )
                    response.dependencies_traced = dep_tree

        # Get annotations for sheet
        if context.include_annotations:
            logger.debug("Getting annotations")
            annotations = self.annotation_management.get_annotations(
                sheet=selection.sheet_name
            )
            response.annotations_found = annotations

    def _explore_active_sheet(
        self,
        sheet_name: str,
        context: QuestionContext,
        response: AssistantResponse,
    ) -> None:
        """
        Explore active sheet (when no selection).

        Args:
            sheet_name: Active sheet name
            context: Question context
            response: Response to populate
        """
        # Get annotations (might give clues about what to look at)
        if context.include_annotations:
            logger.debug("Getting annotations for active sheet")
            annotations = self.annotation_management.get_annotations(sheet=sheet_name)
            response.annotations_found = annotations

            # If we have annotations, might want to explore annotated regions
            # For now, just note them

        # Get workbook structure
        structure = self.workbook_data.get_workbook_structure()
        response.add_context("workbook_structure", str(structure))

    def _query_llm(
        self,
        context: QuestionContext,
        response: AssistantResponse,
    ) -> None:
        """
        Query LLM with gathered context.

        Args:
            context: Question context
            response: Response with gathered context
        """
        logger.debug("Querying LLM")

        # Build formulas list
        formulas = response.context_used.get("formulas", [])

        # Build spatial context
        spatial_context = response.context_used.get("snapshot")

        # Query LLM
        llm_response = self.llm_interaction.query(
            question=context.question,
            selection=context.selection,
            formulas=formulas,
            dependency_tree=response.dependencies_traced,
            annotations=response.annotations_found,
            spatial_context=spatial_context,
            mode=context.mode,
        )

        response.answer = llm_response.content
        response.metadata["llm_provider"] = llm_response.provider
        response.metadata["llm_model"] = llm_response.model
