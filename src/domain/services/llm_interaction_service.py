"""LLM interaction service."""

from typing import Dict, List, Optional

from src.domain.models.annotation import Annotation
from src.domain.models.dependency import DependencyTree
from src.domain.models.query import LLMContext, LLMResponse
from src.domain.models.selection import Selection
from src.infrastructure.config.config_loader import Config
from src.infrastructure.llm.prompt_builder import PromptBuilder
from src.infrastructure.llm.providers.base_provider import BaseLLMProvider
from src.infrastructure.llm.providers.manual_provider import ManualLLMProvider
from src.infrastructure.llm.providers.mock_provider import MockLLMProvider
from src.shared.exceptions import LLMProviderError
from src.shared.logging import get_logger

logger = get_logger(__name__)


class LLMInteractionService:
    """
    Service for interacting with LLM providers.

    Manages LLM providers and handles prompt building and querying.
    """

    def __init__(self, config: Config):
        """
        Initialize LLM interaction service.

        Args:
            config: Application configuration
        """
        self.config = config
        self.providers: Dict[str, BaseLLMProvider] = {}
        self._initialize_providers()

    def _initialize_providers(self) -> None:
        """Initialize configured LLM providers."""
        providers_config = self.config.llm.providers

        # Manual provider
        if providers_config.manual.enabled:
            self.providers["manual"] = ManualLLMProvider(
                input_file=providers_config.manual.input_file,
                output_file=providers_config.manual.output_file,
            )
            logger.debug("Manual LLM provider initialized")

        # Mock provider (always available for testing)
        if providers_config.mock.enabled:
            self.providers["mock"] = MockLLMProvider()
            logger.debug("Mock LLM provider initialized")

        # TODO: Add other providers (internal_api, openai, claude) in Phase 2

    def query(
        self,
        question: str,
        selection: Optional[Selection] = None,
        formulas: Optional[List[str]] = None,
        dependency_tree: Optional[DependencyTree] = None,
        annotations: Optional[List[Annotation]] = None,
        spatial_context: Optional[str] = None,
        mode: str = "educational",
        provider_name: Optional[str] = None,
    ) -> LLMResponse:
        """
        Query LLM with context.

        Args:
            question: User's question
            selection: User's selection
            formulas: Relevant formulas
            dependency_tree: Dependency tree
            annotations: Relevant annotations
            spatial_context: Snapshot or spatial info
            mode: Query mode (educational, technical, concise)
            provider_name: Specific provider to use (None for default)

        Returns:
            LLM response

        Raises:
            LLMProviderError: If query fails
        """
        # Build context
        context = PromptBuilder.build_context(
            question=question,
            selection=selection,
            formulas=formulas,
            dependency_tree=dependency_tree,
            annotations=annotations,
            spatial_context=spatial_context,
            mode=mode,
        )

        # Get system prompt
        system_prompt = PromptBuilder.get_system_prompt(mode)

        # Get provider
        if provider_name is None:
            provider_name = self.config.llm.default_provider

        provider = self.get_provider(provider_name)

        # Query
        logger.info(f"Querying LLM provider: {provider_name}")
        logger.debug(f"Question: {question}")
        logger.debug(f"Context token estimate: {context.token_estimate()}")

        response = provider.query(context, system_prompt)

        logger.info(f"Received response from {provider_name}")

        return response

    def get_provider(self, name: str) -> BaseLLMProvider:
        """
        Get LLM provider by name.

        Args:
            name: Provider name

        Returns:
            Provider instance

        Raises:
            LLMProviderError: If provider not found or unavailable
        """
        if name not in self.providers:
            available = ", ".join(self.providers.keys())
            raise LLMProviderError(
                f"Provider '{name}' not found. Available providers: {available}"
            )

        provider = self.providers[name]

        if not provider.is_available():
            raise LLMProviderError(f"Provider '{name}' is not available")

        return provider

    def list_providers(self) -> List[str]:
        """
        List available provider names.

        Returns:
            List of provider names
        """
        return list(self.providers.keys())

    def is_provider_available(self, name: str) -> bool:
        """
        Check if provider is available.

        Args:
            name: Provider name

        Returns:
            True if provider exists and is available
        """
        if name not in self.providers:
            return False

        return self.providers[name].is_available()
