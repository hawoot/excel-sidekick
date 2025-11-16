"""Mock LLM provider for testing."""

from src.domain.models.query import LLMContext, LLMResponse
from src.infrastructure.llm.providers.base_provider import BaseLLMProvider


class MockLLMProvider(BaseLLMProvider):
    """
    Mock LLM provider for testing.

    Returns canned responses for testing without actual LLM calls.
    """

    def __init__(self, canned_response: str = "This is a mock response from the test LLM."):
        """
        Initialize mock provider.

        Args:
            canned_response: Response to return for all queries
        """
        self.canned_response = canned_response
        self.last_context: LLMContext | None = None
        self.last_system_prompt: str = ""
        self.call_count = 0

    def query(self, context: LLMContext, system_prompt: str = "") -> LLMResponse:
        """
        Return canned response.

        Args:
            context: LLM context (stored for inspection)
            system_prompt: System prompt (stored for inspection)

        Returns:
            Mock LLM response
        """
        # Store for testing/inspection
        self.last_context = context
        self.last_system_prompt = system_prompt
        self.call_count += 1

        # Return canned response
        return LLMResponse(
            content=self.canned_response,
            provider="mock",
            model="mock-model-v1",
        )

    def is_available(self) -> bool:
        """Mock provider is always available."""
        return True

    def get_name(self) -> str:
        """Get provider name."""
        return "mock"

    def reset(self) -> None:
        """Reset call tracking (useful for tests)."""
        self.last_context = None
        self.last_system_prompt = ""
        self.call_count = 0
