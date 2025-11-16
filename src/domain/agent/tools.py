"""Agent tools (for future LangGraph integration).

Note: For Phase 1, we're using a simplified exploration agent that doesn't
require explicit LangGraph tools since we're using a manual file-based LLM provider.

The agent in exploration_agent.py directly calls domain services to gather context,
then makes a single LLM call via the manual provider.

This file is a placeholder for Phase 2 when we integrate with API-based LLM providers
and can use full LangGraph capabilities with tools/function calling.

Phase 2 will define tools like:
- get_workbook_overview_tool
- get_snapshot_tool
- trace_dependencies_tool
- etc.

Each tool will be a LangGraph tool decorator wrapping our domain services.
"""

# TODO: Implement LangGraph tools for Phase 2 when using API-based providers
