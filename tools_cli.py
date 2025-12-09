#!/usr/bin/env python3
"""
Tools CLI for Google Workspace MCP Server

This script allows you to test MCP tools directly without the protocol overhead.
You can call tools as regular Python async functions with real or mocked credentials.

Usage:
    python tools_cli.py --tool get_doc_content --document_id "doc_id_here" --user_google_email "user@example.com"
    python tools_cli.py --interactive  # Interactive REPL mode
"""
import argparse
import asyncio
import logging
import os
import sys
from typing import Any, Dict
from dotenv import load_dotenv

# Load environment variables
dotenv_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env')
load_dotenv(dotenv_path=dotenv_path)

# Suppress googleapiclient discovery cache warning
logging.getLogger('googleapiclient.discovery_cache').setLevel(logging.ERROR)

# Configure logging - use WARNING to reduce noise
logging.basicConfig(
    level=os.getenv('LOG_LEVEL', 'WARNING'),
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def init_server():
    """Initialize the server and import all tools."""
    try:
        # Import server and initialize
        from auth.oauth_config import reload_oauth_config
        from core.server import server, set_transport_mode
        from core.tool_registry import set_enabled_tools as set_enabled_tool_names, wrap_server_tool_method
        from auth.scopes import set_enabled_tools

        # Reload OAuth configuration
        reload_oauth_config()

        # Configure for testing
        set_transport_mode('stdio')  # Use stdio mode by default for testing

        # Import all tool modules to register them (side-effect imports)
        import gmail.gmail_tools  # noqa: F401
        import gdrive.drive_tools  # noqa: F401
        import gcalendar.calendar_tools  # noqa: F401
        import gdocs.docs_tools  # noqa: F401
        import gsheets.sheets_tools  # noqa: F401
        import gchat.chat_tools  # noqa: F401
        import gforms.forms_tools  # noqa: F401
        import gslides.slides_tools  # noqa: F401
        import gtasks.tasks_tools  # noqa: F401
        import gsearch.search_tools  # noqa: F401

        # Enable all tools
        all_services = ['gmail', 'drive', 'calendar', 'docs', 'sheets', 'chat', 'forms', 'slides', 'tasks', 'search']
        set_enabled_tools(all_services)
        set_enabled_tool_names(None)

        # Wrap and filter tools
        wrap_server_tool_method(server)
        from core.tool_registry import filter_server_tools
        filter_server_tools(server)
        
        return server
    except Exception as e:
        logger.error(f"Failed to initialize server: {e}")
        raise


class ToolTester:
    """Helper class to test MCP tools directly."""
    
    def __init__(self, server_instance):
        self.server = server_instance
        self.tools = {}
        
    async def _collect_tools(self) -> Dict[str, Any]:
        """Collect all registered tools from the server."""
        tools = {}
        if hasattr(self.server, '_tool_manager'):
            tool_manager = self.server._tool_manager
            if hasattr(tool_manager, '_tools'):
                # FastMCP stores tools in _tools dict directly, no list_tools() method
                tools = dict(tool_manager._tools)
        return tools
    
    async def init_tools(self):
        """Initialize tools collection (must be called after __init__)."""
        self.tools = await self._collect_tools()
    
    def list_tools(self) -> None:
        """Print all available tools."""
        print("\n📋 Available Tools:")
        print("=" * 60)
        for name, tool in sorted(self.tools.items()):
            desc = tool.description.split('\n')[0] if tool.description else "No description"
            print(f"  • {name}")
            print(f"    {desc}")
            print()
    
    def get_tool_info(self, tool_name: str) -> None:
        """Print detailed information about a specific tool."""
        if tool_name not in self.tools:
            print(f"❌ Tool '{tool_name}' not found.")
            return
        
        tool = self.tools[tool_name]
        print(f"\n🔧 Tool: {tool_name}")
        print("=" * 60)
        print(f"Description: {tool.description}")
        print("\nParameters:")
        
        if hasattr(tool, 'fn'):
            import inspect
            sig = inspect.signature(tool.fn)
            for param_name, param in sig.parameters.items():
                if param_name in ['service', 'drive_service', 'docs_service', 'sheets_service']:
                    continue  # Skip injected service parameters
                annotation = param.annotation if param.annotation != inspect.Parameter.empty else "Any"
                default = f" = {param.default}" if param.default != inspect.Parameter.empty else ""
                print(f"  • {param_name}: {annotation}{default}")
        print()
    
    async def call_tool(self, tool_name: str, **kwargs) -> Any:
        """Call a tool with the given parameters."""
        if tool_name not in self.tools:
            raise ValueError(f"Tool '{tool_name}' not found. Use list_tools() to see available tools.")
        
        tool = self.tools[tool_name]
        
        # Remove None values from kwargs
        kwargs = {k: v for k, v in kwargs.items() if v is not None}
        
        print(f"\n🚀 Calling tool: {tool_name}")
        print(f"   Parameters: {kwargs}")
        print("=" * 60)
        
        try:
            # Call the tool function directly
            if hasattr(tool, 'fn'):
                result = await tool.fn(**kwargs)
            else:
                result = await tool(**kwargs)
            
            print("\n✅ Result:")
            print("-" * 60)
            print(result)
            print("-" * 60)
            return result
        except Exception as e:
            print(f"\n❌ Error: {e}")
            logger.exception("Tool execution failed")
            raise


def interactive_mode(tester: ToolTester):
    """Run an interactive REPL for testing tools."""
    print("\n🎯 Interactive Test Mode")
    print("=" * 60)
    print("Commands:")
    print("  list              - List all available tools")
    print("  info <tool_name>  - Get detailed info about a tool")
    print("  call <tool_name>  - Call a tool (will prompt for parameters)")
    print("  quit              - Exit interactive mode")
    print("=" * 60)
    
    while True:
        try:
            cmd = input("\n> ").strip()
            
            if not cmd:
                continue
            
            if cmd == "quit":
                print("👋 Goodbye!")
                break
            
            if cmd == "list":
                tester.list_tools()
                continue
            
            if cmd.startswith("info "):
                tool_name = cmd[5:].strip()
                tester.get_tool_info(tool_name)
                continue
            
            if cmd.startswith("call "):
                tool_name = cmd[5:].strip()
                if tool_name not in tester.tools:
                    print(f"❌ Tool '{tool_name}' not found.")
                    continue
                
                # Get tool parameters
                tester.get_tool_info(tool_name)
                
                # Prompt for parameters
                print("\nEnter parameters (press Enter to skip optional parameters):")
                kwargs = {}
                
                tool = tester.tools[tool_name]
                if hasattr(tool, 'fn'):
                    import inspect
                    sig = inspect.signature(tool.fn)
                    for param_name, param in sig.parameters.items():
                        if param_name in ['service', 'drive_service', 'docs_service', 'sheets_service']:
                            continue
                        
                        if param.default == inspect.Parameter.empty:
                            # Required parameter
                            value = input(f"  {param_name} (required): ")
                            kwargs[param_name] = value
                        else:
                            # Optional parameter
                            default_str = f" [default: {param.default}]" if param.default is not None else ""
                            value = input(f"  {param_name} (optional){default_str}: ")
                            if value:
                                kwargs[param_name] = value
                
                # Execute the tool
                asyncio.run(tester.call_tool(tool_name, **kwargs))
            else:
                print("❌ Unknown command. Try 'list', 'info <tool>', 'call <tool>', or 'quit'")
                
        except KeyboardInterrupt:
            print("\n👋 Goodbye!")
            break
        except Exception as e:
            print(f"❌ Error: {e}")


def main():
    parser = argparse.ArgumentParser(
        description="CLI for Google Workspace MCP tools",
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    
    parser.add_argument('--interactive', '-i', action='store_true',
                        help='Run in interactive REPL mode')
    parser.add_argument('--list', '-l', action='store_true',
                        help='List all available tools')
    parser.add_argument('--tool', '-t', type=str,
                        help='Tool name to call')
    parser.add_argument('--info', type=str,
                        help='Show detailed info about a tool')
    parser.add_argument('--user_google_email', type=str,
                        default=os.getenv('USER_GOOGLE_EMAIL'),
                        help='User email for authentication (can also be set via USER_GOOGLE_EMAIL env var)')
    parser.add_argument('--verbose', '-v', action='store_true',
                        help='Enable verbose logging')
    
    # Allow arbitrary additional arguments for tool parameters
    args, unknown = parser.parse_known_args()
    
    # Set verbose logging if requested
    if args.verbose:
        logging.getLogger().setLevel(logging.INFO)

    # Parse unknown args as tool parameters (raw strings initially)
    import json
    import inspect
    raw_kwargs = {}
    i = 0
    while i < len(unknown):
        arg = unknown[i]
        if arg.startswith('--'):
            param_name = arg[2:]
            if i + 1 < len(unknown) and not unknown[i + 1].startswith('--'):
                raw_kwargs[param_name] = unknown[i + 1]
                i += 2
            else:
                raw_kwargs[param_name] = True
                i += 1
        else:
            i += 1

    def convert_value_by_type(value: str, expected_type) -> Any:
        """Convert a string value based on expected type annotation."""
        if value is True:  # Boolean flag (--flag without value)
            return value

        # Handle Optional types (e.g., Optional[int] = Union[int, None])
        origin = getattr(expected_type, '__origin__', None)
        if origin is type(None):
            return None

        # Extract the actual type from Optional/Union
        type_args = getattr(expected_type, '__args__', ())
        if type_args:
            # Filter out NoneType from Union types
            non_none_types = [t for t in type_args if t is not type(None)]
            if non_none_types:
                expected_type = non_none_types[0]

        # Convert based on expected type
        if expected_type is bool:
            return value.lower() == 'true'
        elif expected_type is int:
            return int(value)
        elif expected_type is float:
            return float(value)
        elif expected_type in (list, dict) or origin in (list, dict):
            try:
                return json.loads(value)
            except json.JSONDecodeError:
                return value
        else:
            # Default: keep as string
            return value

    def convert_value_fallback(value: str) -> Any:
        """Fallback conversion when no type info is available."""
        if value is True:
            return value
        if value.lower() in ('true', 'false'):
            return value.lower() == 'true'
        elif value.startswith('[') or value.startswith('{'):
            try:
                return json.loads(value)
            except json.JSONDecodeError:
                return value
        # Don't auto-convert numeric strings - keep as string by default
        return value

    def parse_tool_kwargs(tool, raw_kwargs: dict) -> dict:
        """Parse raw kwargs using tool's type annotations."""
        tool_kwargs = {}
        type_hints = {}

        # Get type hints from tool function
        if hasattr(tool, 'fn'):
            try:
                type_hints = inspect.signature(tool.fn).parameters
            except (ValueError, TypeError):
                pass

        for param_name, value in raw_kwargs.items():
            if param_name in type_hints:
                param = type_hints[param_name]
                if param.annotation != inspect.Parameter.empty:
                    tool_kwargs[param_name] = convert_value_by_type(value, param.annotation)
                else:
                    tool_kwargs[param_name] = convert_value_fallback(value)
            else:
                tool_kwargs[param_name] = convert_value_fallback(value)

        return tool_kwargs

    # Initialize server
    print("🔧 Initializing server...")
    try:
        server = init_server()
        tester = ToolTester(server)
        # Initialize tools asynchronously
        asyncio.run(tester.init_tools())
        print(f"✅ Server initialized ({len(tester.tools)} tools loaded)")
    except Exception as e:
        print(f"❌ Failed to initialize: {e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)
    
    if args.list:
        tester.list_tools()
        return
    
    if args.info:
        tester.get_tool_info(args.info)
        return
    
    if args.interactive:
        interactive_mode(tester)
        return
    
    if args.tool:
        # Get the tool to access type hints for conversion
        if args.tool not in tester.tools:
            print(f"❌ Tool '{args.tool}' not found. Use --list to see available tools.")
            sys.exit(1)

        tool = tester.tools[args.tool]
        # Parse kwargs using tool's type annotations
        tool_kwargs = parse_tool_kwargs(tool, raw_kwargs)

        # Add user_google_email to kwargs if provided
        if args.user_google_email:
            tool_kwargs['user_google_email'] = args.user_google_email

        asyncio.run(tester.call_tool(args.tool, **tool_kwargs))
        return
    
    # No action specified, show help
    parser.print_help()
    print("\n💡 Quick start:")
    print("  python tools_cli.py --list                           # List all tools")
    print("  python tools_cli.py --info get_doc_content           # Get tool info")
    print("  python tools_cli.py --interactive                    # Interactive mode")
    print("  python tools_cli.py --tool search_docs --query 'test' --user_google_email user@example.com")


if __name__ == "__main__":
    main()

