"""Manual LLM provider using file-based interaction."""

from pathlib import Path

from src.domain.models.query import LLMContext, LLMResponse
from src.infrastructure.llm.providers.base_provider import BaseLLMProvider
from src.shared.exceptions import LLMProviderError
from src.shared.logging import get_logger

logger = get_logger(__name__)


class ManualLLMProvider(BaseLLMProvider):
    """
    Manual LLM provider for proof of concept.

    Writes prompt to a file, waits for user to copy-paste to/from
    Copilot or internal LLMSuite, then reads response from another file.
    """

    def __init__(self, input_file: str = "llm_input.txt", output_file: str = "llm_output.txt"):
        """
        Initialize manual provider.

        Args:
            input_file: File to write prompts to
            output_file: File to read responses from
        """
        self.input_file = Path(input_file)
        self.output_file = Path(output_file)

    def query(self, context: LLMContext, system_prompt: str = "") -> LLMResponse:
        """
        Query LLM via manual file-based interaction.

        Args:
            context: LLM context
            system_prompt: Optional system prompt

        Returns:
            LLM response

        Raises:
            LLMProviderError: If file operations fail
        """
        try:
            # Build full prompt
            prompt = context.to_prompt(system_prompt)

            # Write to input file
            self.input_file.write_text(prompt, encoding="utf-8")

            logger.info(f" Prompt written to: {self.input_file}")
            logger.info(
                f"\n{'='*80}\n"
                f"NEXT STEPS:\n"
                f"1. Open {self.input_file}\n"
                f"2. Copy the contents\n"
                f"3. Paste into Copilot or LLMSuite\n"
                f"4. Copy the response\n"
                f"5. Paste into {self.output_file}\n"
                f"{'='*80}\n"
            )

            # Wait for user input
            input(f"\nPress Enter when you've pasted the response to {self.output_file}...")

            # Read response
            if not self.output_file.exists():
                raise LLMProviderError(
                    f"Response file not found: {self.output_file}. "
                    f"Please create it and paste the LLM response."
                )

            response_text = self.output_file.read_text(encoding="utf-8")

            if not response_text.strip():
                raise LLMProviderError(
                    f"Response file is empty: {self.output_file}. "
                    f"Please paste the LLM response."
                )

            logger.info(" Response loaded successfully")

            # Clean up output file for next use
            try:
                self.output_file.write_text("", encoding="utf-8")
            except Exception:
                pass

            return LLMResponse(
                content=response_text.strip(),
                provider="manual",
                model=None,
            )

        except LLMProviderError:
            raise
        except Exception as e:
            raise LLMProviderError(f"Manual provider failed: {e}")

    def is_available(self) -> bool:
        """Check if manual provider is available (always true)."""
        return True

    def get_name(self) -> str:
        """Get provider name."""
        return "manual"
