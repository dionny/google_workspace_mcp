# ADR 003: Optional Optimizer Mode for Large Tool Sets

## Status
Accepted

## Context
The Google Workspace MCP server exposes a large number of tools across multiple services (Gmail, Drive, Calendar, Docs, Sheets, etc.). With the complete tool set, there can be hundreds of individual tool schemas that need to be sent to LLM clients.

This creates several challenges:
1. **Context Window Usage**: Large tool schemas consume significant portions of the LLM's context window
2. **Performance**: Parsing and processing hundreds of tools can slow down initial connections
3. **Cost**: More tokens used for tool schemas means higher API costs
4. **Cognitive Load**: LLMs may struggle to select the right tool from hundreds of options

## Decision
We have implemented an optional **Optimizer Mode** that uses semantic search to enable on-demand tool discovery and execution. When enabled via the `--optimizer` flag, the server:

1. **Loads all enabled tools normally** (respecting `--tools` and `--tool-tier` flags)
2. **Computes semantic embeddings** for all tool descriptions using sentence-transformers
3. **Stores tool functions** for later execution
4. **Replaces the tool registry** with just 4 meta-tools:
   - `google_workspace_find_tool` - Semantic search for tools by natural language description
   - `google_workspace_describe_tool` - Get full schema for a specific tool by name
   - `google_workspace_list_tools` - List all available tools (with optional service filter)
   - `google_workspace_call_tool` - Execute a discovered tool with provided arguments

5. **LLM workflow**:
   - Instead of seeing all tools upfront, LLM starts with 4 meta-tools
   - To accomplish a task, LLM searches for relevant tools using natural language
   - Once found, LLM retrieves the full schema to understand parameters
   - LLM then calls the tool via `google_workspace_call_tool` with proper arguments
   - The actual tool execution happens through the optimizer's stored function references

### Implementation Details

**Core Components**:
- `core/optimizer.py` - Semantic search engine using sentence-transformers (all-MiniLM-L6-v2 model)
  - Stores both tool definitions (metadata) and tool functions (callables)
  - Supports filtering by service (gmail, docs, sheets, etc.)
  - Handles both sync and async tool functions
- `core/server.py` - Four optimizer meta-tools with `google_workspace_` prefix
- `main.py` - CLI flag and initialization logic, service detection from tool modules

**Dependencies** (optional):
- `numpy>=1.24.0`
- `sentence-transformers>=2.2.0`

Specified as optional dependencies in `pyproject.toml`:
```toml
[project.optional-dependencies]
optimizer = [
    "numpy>=1.24.0",
    "sentence-transformers>=2.2.0",
]
```

**Installation**:
```bash
pip install numpy sentence-transformers
# or with uv
uv pip install numpy sentence-transformers
```

### Usage

```bash
# Enable optimizer mode
uv run main.py --optimizer

# Works with tool selection
uv run main.py --optimizer --tools gmail drive calendar

# Works with tool tiers
uv run main.py --optimizer --tool-tier core
```

### Design Principles

1. **Fully Optional**: Optimizer mode is opt-in via CLI flag, doesn't affect normal operation
2. **Transparent to Tools**: Underlying tools are unchanged, only discovery mechanism differs
3. **Respects Configuration**: Works with existing `--tools` and `--tool-tier` flags
4. **Graceful Degradation**: Clear error if dependencies not installed
5. **Namespace Isolation**: All meta-tools prefixed with `google_workspace_` to avoid conflicts

## Consequences

### Positive
- **Reduced Context Usage**: Only 4 tool schemas sent initially instead of hundreds
- **Improved Performance**: Faster initial connections and tool list processing
- **Lower Costs**: Fewer tokens used for tool schemas
- **Semantic Discovery**: LLMs can find tools using natural language, not exact names
- **Scalable**: Can handle arbitrarily large tool sets without overwhelming LLMs
- **Service Filtering**: Can list tools by specific service (gmail, docs, etc.)
- **Direct Execution**: Tools can be called directly through the optimizer

### Negative
- **Additional Dependencies**: Requires numpy and sentence-transformers (~500MB download)
- **Initialization Time**: First startup downloads the embedding model and computes embeddings (30-60 seconds)
- **Memory Usage**: Embedding model and vectors consume additional memory (~200MB)
- **Three-Step Process**: Requires find → describe → call workflow vs. direct tool calling
- **LLM Adaptation**: Not all LLMs may effectively use the meta-tool pattern
- **Function Storage**: Must maintain references to all tool functions in memory

### Neutral
- **Model Download**: First run downloads all-MiniLM-L6-v2 (~90MB) from HuggingFace
- **Embedding Computation**: Done at startup, takes ~1-5 seconds for typical tool sets
- **Search Quality**: Depends on quality of tool descriptions and query phrasing

## Alternatives Considered

### 1. Static Tool Groups
Group tools by category and expose groups as separate resources.
- **Rejected**: Still requires sending many tool schemas, just organized differently

### 2. Lazy Loading via Resources
Use MCP resources to provide tool schemas on-demand.
- **Rejected**: Resources are for data, not tool schemas; doesn't reduce initial tool count

### 3. Server-Side Tool Selection
Have LLM describe intent, server selects and calls tools automatically.
- **Rejected**: Reduces LLM control and transparency; harder to debug

### 4. Manual Tool Filtering
Require users to specify which tools they want via config.
- **Rejected**: Puts burden on users to know which tools exist; defeats purpose of comprehensive server

## Implementation Notes

### Semantic Search
We use `sentence-transformers` with the `all-MiniLM-L6-v2` model because:
- Small and fast (~80MB model, ~10ms inference)
- Good semantic understanding for English descriptions
- Widely used and well-tested
- No API keys or external services required

### Tool Metadata
Each tool provides metadata for semantic search:
```python
{
    "name": "send_gmail_message",
    "description": "Send an email message via Gmail...",
    "inputSchema": {...}
}
```

The description is used for embedding and matching.

### Cosine Similarity
We use cosine similarity for semantic matching:
1. Embed tool descriptions at startup
2. Embed user query at search time
3. Compute cosine similarity between query and all tools
4. Return top-k matches sorted by similarity score

## Future Considerations

1. **Caching**: Cache embeddings across server restarts
2. **Custom Models**: Support for other embedding models
3. **Usage Analytics**: Track which tools are most commonly searched for
4. **Tool Recommendations**: Suggest related tools based on usage patterns
5. **Hybrid Search**: Combine semantic search with keyword matching
6. **Tool Clustering**: Group similar tools for better organization

## References

- Inspiration: [MCP Tool Optimizer Example](https://github.com/example/optimize-mcp)
- Sentence Transformers: https://www.sbert.net/
- MCP Protocol: https://modelcontextprotocol.io/

## Date
December 9, 2024

## Authors
- mbradshaw (with AI assistance)

