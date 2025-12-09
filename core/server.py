import logging
import json
from typing import List, Optional
from importlib import metadata

from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
from starlette.applications import Starlette
from starlette.requests import Request
from starlette.middleware import Middleware

from fastmcp import FastMCP
from fastmcp.server.auth.providers.google import GoogleProvider

from auth.oauth21_session_store import get_oauth21_session_store, set_auth_provider
from auth.google_auth import handle_auth_callback, start_auth_flow, check_client_secrets
from auth.mcp_session_middleware import MCPSessionMiddleware
from auth.oauth_responses import (
    create_error_response,
    create_success_response,
    create_server_error_response,
)
from auth.auth_info_middleware import AuthInfoMiddleware
from auth.scopes import SCOPES, get_current_scopes  # noqa
from core.config import (
    USER_GOOGLE_EMAIL,
    get_transport_mode,
    set_transport_mode as _set_transport_mode,
    get_oauth_redirect_uri as get_oauth_redirect_uri_for_current_mode,
)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

_auth_provider: Optional[GoogleProvider] = None
_legacy_callback_registered = False

session_middleware = Middleware(MCPSessionMiddleware)


# Custom FastMCP that adds secure middleware stack for OAuth 2.1
class SecureFastMCP(FastMCP):
    def streamable_http_app(self) -> "Starlette":
        """Override to add secure middleware stack for OAuth 2.1."""
        app = super().streamable_http_app()

        # Add middleware in order (first added = outermost layer)
        # Session Management - extracts session info for MCP context
        app.user_middleware.insert(0, session_middleware)

        # Rebuild middleware stack
        app.middleware_stack = app.build_middleware_stack()
        logger.info("Added middleware stack: Session Management")
        return app


server = SecureFastMCP(
    name="google_workspace",
    auth=None,
)

# Add the AuthInfo middleware to inject authentication into FastMCP context
auth_info_middleware = AuthInfoMiddleware()
server.add_middleware(auth_info_middleware)


def set_transport_mode(mode: str):
    """Sets the transport mode for the server."""
    _set_transport_mode(mode)
    logger.info(f"Transport: {mode}")


def _ensure_legacy_callback_route() -> None:
    global _legacy_callback_registered
    if _legacy_callback_registered:
        return
    server.custom_route("/oauth2callback", methods=["GET"])(legacy_oauth2_callback)
    _legacy_callback_registered = True


def configure_server_for_http():
    """
    Configures the authentication provider for HTTP transport.
    This must be called BEFORE server.run().
    """
    global _auth_provider

    transport_mode = get_transport_mode()

    if transport_mode != "streamable-http":
        return

    # Use centralized OAuth configuration
    from auth.oauth_config import get_oauth_config

    config = get_oauth_config()

    # Check if OAuth 2.1 is enabled via centralized config
    oauth21_enabled = config.is_oauth21_enabled()

    if oauth21_enabled:
        if not config.is_configured():
            logger.warning("OAuth 2.1 enabled but OAuth credentials not configured")
            return

        try:
            required_scopes: List[str] = sorted(get_current_scopes())

            # Check if external OAuth provider is configured
            if config.is_external_oauth21_provider():
                # External OAuth mode: use custom provider that handles ya29.* access tokens
                from auth.external_oauth_provider import ExternalOAuthProvider

                provider = ExternalOAuthProvider(
                    client_id=config.client_id,
                    client_secret=config.client_secret,
                    base_url=config.get_oauth_base_url(),
                    redirect_path=config.redirect_path,
                    required_scopes=required_scopes,
                )
                # Disable protocol-level auth, expect bearer tokens in tool calls
                server.auth = None
                logger.info(
                    "OAuth 2.1 enabled with EXTERNAL provider mode - protocol-level auth disabled"
                )
                logger.info(
                    "Expecting Authorization bearer tokens in tool call headers"
                )
            else:
                # Standard OAuth 2.1 mode: use FastMCP's GoogleProvider
                provider = GoogleProvider(
                    client_id=config.client_id,
                    client_secret=config.client_secret,
                    base_url=config.get_oauth_base_url(),
                    redirect_path=config.redirect_path,
                    required_scopes=required_scopes,
                )
                # Enable protocol-level auth
                server.auth = provider
                logger.info(
                    "OAuth 2.1 enabled using FastMCP GoogleProvider with protocol-level auth"
                )

            # Always set auth provider for token validation in middleware
            set_auth_provider(provider)
            _auth_provider = provider
        except Exception as exc:
            logger.error(
                "Failed to initialize FastMCP GoogleProvider: %s", exc, exc_info=True
            )
            raise
    else:
        logger.info("OAuth 2.0 mode - Server will use legacy authentication.")
        server.auth = None
        _auth_provider = None
        set_auth_provider(None)
        _ensure_legacy_callback_route()


def get_auth_provider() -> Optional[GoogleProvider]:
    """Gets the global authentication provider instance."""
    return _auth_provider


@server.custom_route("/health", methods=["GET"])
async def health_check(request: Request):
    try:
        version = metadata.version("workspace-mcp")
    except metadata.PackageNotFoundError:
        version = "dev"
    return JSONResponse(
        {
            "status": "healthy",
            "service": "workspace-mcp",
            "version": version,
            "transport": get_transport_mode(),
        }
    )


@server.custom_route("/attachments/{file_id}", methods=["GET"])
async def serve_attachment(file_id: str, request: Request):
    """Serve a stored attachment file."""
    from core.attachment_storage import get_attachment_storage

    storage = get_attachment_storage()
    metadata = storage.get_attachment_metadata(file_id)

    if not metadata:
        return JSONResponse(
            {"error": "Attachment not found or expired"}, status_code=404
        )

    file_path = storage.get_attachment_path(file_id)
    if not file_path:
        return JSONResponse({"error": "Attachment file not found"}, status_code=404)

    return FileResponse(
        path=str(file_path),
        filename=metadata["filename"],
        media_type=metadata["mime_type"],
    )


async def legacy_oauth2_callback(request: Request) -> HTMLResponse:
    state = request.query_params.get("state")
    code = request.query_params.get("code")
    error = request.query_params.get("error")

    if error:
        msg = (
            f"Authentication failed: Google returned an error: {error}. State: {state}."
        )
        logger.error(msg)
        return create_error_response(msg)

    if not code:
        msg = "Authentication failed: No authorization code received from Google."
        logger.error(msg)
        return create_error_response(msg)

    try:
        error_message = check_client_secrets()
        if error_message:
            return create_server_error_response(error_message)

        logger.info(f"OAuth callback: Received code (state: {state}).")

        mcp_session_id = None
        if hasattr(request, "state") and hasattr(request.state, "session_id"):
            mcp_session_id = request.state.session_id

        verified_user_id, credentials = handle_auth_callback(
            scopes=get_current_scopes(),
            authorization_response=str(request.url),
            redirect_uri=get_oauth_redirect_uri_for_current_mode(),
            session_id=mcp_session_id,
        )

        logger.info(
            f"OAuth callback: Successfully authenticated user: {verified_user_id}."
        )

        try:
            store = get_oauth21_session_store()

            store.store_session(
                user_email=verified_user_id,
                access_token=credentials.token,
                refresh_token=credentials.refresh_token,
                token_uri=credentials.token_uri,
                client_id=credentials.client_id,
                client_secret=credentials.client_secret,
                scopes=credentials.scopes,
                expiry=credentials.expiry,
                session_id=f"google-{state}",
                mcp_session_id=mcp_session_id,
            )
            logger.info(
                f"Stored Google credentials in OAuth 2.1 session store for {verified_user_id}"
            )
        except Exception as e:
            logger.error(f"Failed to store credentials in OAuth 2.1 store: {e}")

        return create_success_response(verified_user_id)
    except Exception as e:
        logger.error(f"Error processing OAuth callback: {str(e)}", exc_info=True)
        return create_server_error_response(str(e))


@server.tool()
async def start_google_auth(
    service_name: str, user_google_email: str = USER_GOOGLE_EMAIL
) -> str:
    """
    Manually initiate Google OAuth authentication flow.

    NOTE: This tool should typically NOT be called directly. The authentication system
    automatically handles credential checks and prompts for authentication when needed.
    Only use this tool if:
    1. You need to re-authenticate with different credentials
    2. You want to proactively authenticate before using other tools
    3. The automatic authentication flow failed and you need to retry

    In most cases, simply try calling the Google Workspace tool you need - it will
    automatically handle authentication if required.
    """
    if not user_google_email:
        raise ValueError("user_google_email must be provided.")

    error_message = check_client_secrets()
    if error_message:
        return f"**Authentication Error:** {error_message}"

    try:
        auth_message = await start_auth_flow(
            user_google_email=user_google_email,
            service_name=service_name,
            redirect_uri=get_oauth_redirect_uri_for_current_mode(),
        )
        return auth_message
    except Exception as e:
        logger.error(f"Failed to start Google authentication flow: {e}", exc_info=True)
        return f"**Error:** An unexpected error occurred: {e}"


# Optimizer mode tools (conditionally registered)
def register_optimizer_tools():
    """Register optimizer tools for on-demand tool discovery."""
    from core.optimizer import get_optimizer

    @server.tool()
    async def google_workspace_find_tool(query: str, top_k: int = 10) -> str:
        """
        Semantic search for Google Workspace tools by description.

        This is the FIRST STEP in using Google Workspace tools in optimizer mode.
        Use this tool to discover which tools are available for your task.

        WORKFLOW:
        1. Use this tool (google_workspace_find_tool) to search for relevant tools
        2. Use google_workspace_describe_tool to get the full schema of a tool
        3. Use google_workspace_call_tool to execute the tool with proper arguments

        Args:
            query: Natural language description of what you're trying to do
                   Examples:
                   - "send an email"
                   - "create a document"
                   - "list calendar events"
                   - "search for files in drive"
            top_k: Number of top matching tools to return (default: 10)

        Returns:
            JSON array of matching tools with:
            - name: The tool name to use in describe_tool and call_tool
            - excerpt: Brief description of what the tool does
            - score: Similarity score (higher is better match)

        Example:
            Query: "send an email"
            Returns: [{"name": "send_gmail_message", "excerpt": "Send an email...", "score": 0.85}]
        """
        optimizer = get_optimizer()
        if optimizer is None:
            return json.dumps(
                {"error": "Optimizer not initialized"}, indent=2
            )

        try:
            results = optimizer.find_similar_tools(query, top_k)
            return json.dumps(results, indent=2)
        except Exception as e:
            logger.error(f"Error in google_workspace_find_tool: {e}", exc_info=True)
            return json.dumps({"error": str(e)}, indent=2)

    @server.tool()
    async def google_workspace_describe_tool(name: str) -> str:
        """
        Get the full definition and parameter schema of a specific tool.

        This is the SECOND STEP after finding a tool with google_workspace_find_tool.
        Use this to understand what parameters the tool requires before calling it.

        WORKFLOW:
        1. google_workspace_find_tool - Search for tools
        2. Use THIS tool to get the full schema
        3. google_workspace_call_tool - Execute with proper arguments

        Args:
            name: The exact tool name from google_workspace_find_tool results
                  (e.g., "send_gmail_message", "create_doc", "list_calendar_events")

        Returns:
            JSON object with:
            - name: Tool name
            - description: Full description of what the tool does
            - inputSchema: Complete JSON schema of all parameters (required and optional)

        Example:
            Input: name="send_gmail_message"
            Returns: {
              "name": "send_gmail_message",
              "description": "Send an email message via Gmail...",
              "inputSchema": {
                "type": "object",
                "properties": {
                  "to": {"type": "string", "description": "Recipient email"},
                  "subject": {"type": "string", "description": "Email subject"},
                  ...
                },
                "required": ["to", "subject", "body"]
              }
            }
        """
        optimizer = get_optimizer()
        if optimizer is None:
            return json.dumps(
                {"error": "Optimizer not initialized"}, indent=2
            )

        try:
            definition = optimizer.get_tool_definition(name)
            return json.dumps(definition, indent=2)
        except ValueError as e:
            return json.dumps({"error": str(e)}, indent=2)
        except Exception as e:
            logger.error(f"Error in google_workspace_describe_tool: {e}", exc_info=True)
            return json.dumps({"error": str(e)}, indent=2)

    @server.tool()
    async def google_workspace_list_tools(service: str = None) -> str:
        """
        List all available Google Workspace tool names, optionally filtered by service.

        Use this as an alternative to google_workspace_find_tool when you want to
        browse all available tools or see all tools for a specific service.

        IMPORTANT: This tool should ONLY be used when you know the specific service you need
        (e.g., 'gmail', 'docs', 'sheets'). For general task-based discovery, use
        google_workspace_find_tool instead with a natural language query.

        RECOMMENDED USAGE:
        - ✅ GOOD: list_tools(service='gmail') - when you know you need Gmail tools
        - ✅ GOOD: After find_tool narrows down to a service, list all tools in that service
        - ❌ AVOID: list_tools() with no filter - this dumps all tools and defeats the optimizer's purpose
        - ❌ AVOID: Using this as your first step - use find_tool for semantic search instead

        Args:
            service: Optional service name to filter by. Valid values:
                    - 'gmail' - Email tools
                    - 'docs' - Google Docs tools
                    - 'sheets' - Google Sheets tools
                    - 'drive' - Google Drive tools
                    - 'calendar' - Google Calendar tools
                    - 'slides' - Google Slides tools
                    - 'forms' - Google Forms tools
                    - 'tasks' - Google Tasks tools
                    - 'chat' - Google Chat tools
                    - 'search' - Custom Search tools
                    - None (default) - All tools (use sparingly!)

        Returns:
            JSON array of tool names that you can then describe or call

        Example Workflow:
            1. find_tool('send an email') → identifies this is a Gmail task
            2. list_tools(service='gmail') → see all available Gmail tools
            3. describe_tool('send_gmail_message') → get full schema
            4. call_tool('send_gmail_message', args) → execute
        """
        optimizer = get_optimizer()
        if optimizer is None:
            return json.dumps(
                {"error": "Optimizer not initialized"}, indent=2
            )

        try:
            tools = optimizer.list_all_tools(service=service)
            
            # Warn if returning too many tools without filtering
            if service is None and len(tools) > 20:
                return json.dumps({
                    "warning": f"Returning {len(tools)} tools without service filter. Consider using google_workspace_find_tool with a natural language query, or filter by service.",
                    "tools": tools,
                    "suggestion": "Use find_tool('your task description') for semantic search, or specify a service parameter to filter results."
                }, indent=2)
            
            return json.dumps(tools, indent=2)
        except Exception as e:
            logger.error(f"Error in google_workspace_list_tools: {e}", exc_info=True)
            return json.dumps({"error": str(e)}, indent=2)

    @server.tool()
    async def google_workspace_call_tool(name: str, arguments: dict) -> str:
        """
        Execute a Google Workspace tool with the given arguments.

        This is the FINAL STEP after finding and understanding a tool.
        Use this to actually execute the tool and perform the action.

        COMPLETE WORKFLOW:
        1. google_workspace_find_tool - Search: "send an email"
           → Returns: [{"name": "send_gmail_message", ...}]
        
        2. google_workspace_describe_tool - Get schema: name="send_gmail_message"
           → Returns: Full parameter schema showing required fields: to, subject, body
        
        3. Use THIS tool to execute: 
           name="send_gmail_message"
           arguments={"to": "user@example.com", "subject": "Hello", "body": "Message"}
           → Sends the actual email

        Args:
            name: The exact tool name from find_tool or describe_tool
                  (e.g., "send_gmail_message", "create_doc", "list_calendar_events")
            
            arguments: Dictionary of arguments matching the tool's inputSchema
                      - Must include all required parameters
                      - Can include optional parameters
                      - Parameter names and types must match the schema exactly

        Returns:
            The result from the tool execution (format varies by tool)
            - Success: Tool output (could be confirmation, data, resource URL, etc.)
            - Error: JSON with error details

        Example:
            Input:
              name="send_gmail_message"
              arguments={
                "to": "colleague@example.com",
                "subject": "Meeting Notes",
                "body": "Here are the notes from our meeting...",
                "user_google_email": "myemail@gmail.com"
              }
            Returns: "Email sent successfully. Message ID: abc123..."
        """
        optimizer = get_optimizer()
        if optimizer is None:
            return json.dumps(
                {"error": "Optimizer not initialized"}, indent=2
            )

        try:
            # Call the underlying tool function
            result = await optimizer.call_tool(name, arguments)
            
            # If result is already a string, return it directly
            if isinstance(result, str):
                return result
            
            # Otherwise, convert to JSON
            return json.dumps({"result": result}, indent=2)
        except ValueError as e:
            return json.dumps({"error": str(e)}, indent=2)
        except Exception as e:
            logger.error(f"Error in google_workspace_call_tool: {e}", exc_info=True)
            return json.dumps({"error": str(e)}, indent=2)

    logger.info("Registered 4 optimizer tools: find_tool, describe_tool, list_tools, call_tool")


