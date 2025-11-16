"""Base LLM provider interface."""

from abc import ABC, abstractmethod

from src.domain.models.query import LLMContext, LLMResponse


class BaseLLMProvider(ABC):
    """
    Abstract base class for LLM providers.

    All LLM providers must implement this interface.
    """

    @abstractmethod
    def query(self, context: LLMContext, system_prompt: str = "") -> LLMResponse:
        """
        Query the LLM with given context.

        Args:
            context: LLM context with question and supporting information
            system_prompt: Optional system prompt

        Returns:
            LLM response

        Raises:
            LLMProviderError: If query fails
        """
        pass

    @abstractmethod
    def is_available(self) -> bool:
        """
        Check if provider is available/configured.

        Returns:
            True if provider can be used
        """
        pass

    @abstractmethod
    def get_name(self) -> str:
        """
        Get provider name.

        Returns:
            Provider name (e.g., "manual", "internal_api", "openai")
        """
        pass
