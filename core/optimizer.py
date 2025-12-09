"""
Tool Optimizer for Google Workspace MCP Server

This module provides optional tool optimization through semantic search,
allowing LLMs to discover and use tools on-demand rather than loading
all tool schemas at once.

Requires: numpy, sentence-transformers
"""

import logging
from typing import Optional, Any
import json

try:
    import numpy as np
    from sentence_transformers import SentenceTransformer

    OPTIMIZER_AVAILABLE = True
except ImportError:
    OPTIMIZER_AVAILABLE = False
    np = None
    SentenceTransformer = None

logger = logging.getLogger(__name__)


class ToolOptimizer:
    """
    Optimizer that wraps a collection of tools for semantic search and on-demand retrieval.
    """

    def __init__(self):
        if not OPTIMIZER_AVAILABLE:
            raise RuntimeError(
                "Optimizer mode requires numpy and sentence-transformers. "
                "Install with: pip install numpy sentence-transformers"
            )

        self.tools: dict[str, Any] = {}  # name -> tool definition
        self.tool_functions: dict[str, Any] = {}  # name -> actual tool function
        self.embeddings: Optional[np.ndarray] = None
        self.tool_names: list[str] = []
        self.embedding_model: Optional[SentenceTransformer] = None
        self._initialized = False

    def initialize(self, tools: dict[str, Any], tool_functions: dict[str, Any]) -> None:
        """
        Initialize the optimizer with a collection of tools.

        Args:
            tools: Dictionary mapping tool names to tool definitions
                   Each tool should have 'name', 'description', 'inputSchema', and 'service'
            tool_functions: Dictionary mapping tool names to their actual callable functions
        """
        if self._initialized:
            logger.warning("Optimizer already initialized, skipping")
            return

        self.tools = tools
        self.tool_functions = tool_functions
        self.tool_names = list(tools.keys())

        logger.info(f"Initializing optimizer with {len(self.tool_names)} tools")

        # Initialize embedding model (using a small, fast model)
        logger.info("Loading embedding model (all-MiniLM-L6-v2)...")
        self.embedding_model = SentenceTransformer("all-MiniLM-L6-v2")

        # Compute embeddings for all tools
        if self.tool_names:
            texts = []
            for name in self.tool_names:
                tool = self.tools[name]
                # Combine name and description for better semantic matching
                text = f"{name}: {tool.get('description', '')}"
                texts.append(text)

            logger.info("Computing embeddings for all tools...")
            self.embeddings = self.embedding_model.encode(texts, convert_to_numpy=True)
            logger.info(f"Embeddings computed with shape {self.embeddings.shape}")

        self._initialized = True
        logger.info("Optimizer initialization complete")

    def find_similar_tools(self, query: str, top_k: int = 10) -> list[dict]:
        """
        Find tools similar to the query using semantic search.

        Args:
            query: Natural language description of what you're looking for
            top_k: Number of top matches to return

        Returns:
            List of dictionaries with 'name', 'excerpt', and 'score' keys
        """
        if not self._initialized:
            raise RuntimeError("Optimizer not initialized")

        if self.embedding_model is None or self.embeddings is None:
            return []

        # Embed the query
        query_embedding = self.embedding_model.encode([query], convert_to_numpy=True)[0]

        # Compute cosine similarity
        # Normalize vectors
        query_norm = query_embedding / np.linalg.norm(query_embedding)
        embeddings_norm = self.embeddings / np.linalg.norm(
            self.embeddings, axis=1, keepdims=True
        )

        # Dot product gives cosine similarity for normalized vectors
        similarities = np.dot(embeddings_norm, query_norm)

        # Get top-k indices
        top_indices = np.argsort(similarities)[::-1][:top_k]

        results = []
        for idx in top_indices:
            name = self.tool_names[idx]
            tool = self.tools[name]
            desc = tool.get("description", "")
            # Truncate description for excerpt
            excerpt = desc[:150] + "..." if len(desc) > 150 else desc
            results.append(
                {"name": name, "excerpt": excerpt, "score": float(similarities[idx])}
            )

        return results

    def get_tool_definition(self, name: str) -> dict:
        """
        Get the full definition of a specific tool.

        Args:
            name: Name of the tool

        Returns:
            Dictionary with 'name', 'description', and 'inputSchema'

        Raises:
            ValueError: If tool not found
        """
        if not self._initialized:
            raise RuntimeError("Optimizer not initialized")

        if name not in self.tools:
            raise ValueError(f"Tool '{name}' not found")

        tool = self.tools[name]
        return {
            "name": tool.get("name", name),
            "description": tool.get("description", ""),
            "inputSchema": tool.get("inputSchema", {}),
        }

    def list_all_tools(self, service: Optional[str] = None) -> list[str]:
        """
        Get a list of all available tool names, optionally filtered by service.

        Args:
            service: Optional service name to filter by (e.g., 'gmail', 'docs', 'sheets')

        Returns:
            List of tool names
        """
        if not self._initialized:
            raise RuntimeError("Optimizer not initialized")

        if service is None:
            return self.tool_names.copy()
        
        # Filter tools by service
        filtered_tools = []
        for name in self.tool_names:
            tool = self.tools[name]
            tool_service = tool.get("service", "")
            if tool_service.lower() == service.lower():
                filtered_tools.append(name)
        
        return filtered_tools

    async def call_tool(self, name: str, arguments: dict) -> Any:
        """
        Execute a tool by name with the given arguments.

        Args:
            name: Name of the tool to call
            arguments: Dictionary of arguments to pass to the tool

        Returns:
            Result from the tool execution

        Raises:
            ValueError: If tool not found
            RuntimeError: If optimizer not initialized
        """
        if not self._initialized:
            raise RuntimeError("Optimizer not initialized")

        if name not in self.tool_functions:
            raise ValueError(f"Tool '{name}' not found")

        tool_func = self.tool_functions[name]
        
        # Call the tool function with unpacked arguments
        # Handle both sync and async functions
        import inspect
        if inspect.iscoroutinefunction(tool_func):
            result = await tool_func(**arguments)
        else:
            result = tool_func(**arguments)
        
        return result


# Global optimizer instance
_optimizer_instance: Optional[ToolOptimizer] = None


def get_optimizer() -> Optional[ToolOptimizer]:
    """Get the global optimizer instance."""
    return _optimizer_instance


def initialize_optimizer(tools: dict[str, Any], tool_functions: dict[str, Any]) -> ToolOptimizer:
    """
    Initialize the global optimizer instance.

    Args:
        tools: Dictionary mapping tool names to tool definitions
        tool_functions: Dictionary mapping tool names to their callable functions

    Returns:
        The initialized optimizer instance
    """
    global _optimizer_instance

    if not OPTIMIZER_AVAILABLE:
        raise RuntimeError(
            "Optimizer mode requires numpy and sentence-transformers. "
            "Install with: pip install numpy sentence-transformers"
        )

    _optimizer_instance = ToolOptimizer()
    _optimizer_instance.initialize(tools, tool_functions)
    return _optimizer_instance


def is_optimizer_available() -> bool:
    """Check if optimizer dependencies are installed."""
    return OPTIMIZER_AVAILABLE

