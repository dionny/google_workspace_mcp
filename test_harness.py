#!/usr/bin/env python3
"""
Test Harness for Google Workspace MCP Server

This script allows you to test MCP tools directly without the protocol overhead.
You can call tools as regular Python async functions with real or mocked credentials.

Usage:
    python test_harness.py --tool get_doc_content --document_id "doc_id_here" --user_google_email "user@example.com"
    python test_harness.py --interactive  # Interactive REPL mode
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

        # Import all tool modules to register them
        import gmail.gmail_tools
        import gdrive.drive_tools
        import gcalendar.calendar_tools
        import gdocs.docs_tools
        import gsheets.sheets_tools
        import gchat.chat_tools
        import gforms.forms_tools
        import gslides.slides_tools
        import gtasks.tasks_tools
        import gsearch.search_tools

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
            tool_list = await self.server._tool_manager.list_tools()
            for tool in tool_list:
                tools[tool.name] = tool
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
        description="Test harness for Google Workspace MCP tools",
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
    
    # Parse unknown args as tool parameters
    tool_kwargs = {}
    i = 0
    while i < len(unknown):
        arg = unknown[i]
        if arg.startswith('--'):
            param_name = arg[2:]
            if i + 1 < len(unknown) and not unknown[i + 1].startswith('--'):
                value = unknown[i + 1]
                # Convert common types
                if value.lower() in ('true', 'false'):
                    tool_kwargs[param_name] = value.lower() == 'true'
                elif value.isdigit() or (value.startswith('-') and value[1:].isdigit()):
                    tool_kwargs[param_name] = int(value)
                else:
                    tool_kwargs[param_name] = value
                i += 2
            else:
                tool_kwargs[param_name] = True
                i += 1
        else:
            i += 1
    
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
        # Add user_google_email to kwargs if provided
        if args.user_google_email:
            tool_kwargs['user_google_email'] = args.user_google_email
        
        asyncio.run(tester.call_tool(args.tool, **tool_kwargs))
        return
    
    # No action specified, show help
    parser.print_help()
    print("\n💡 Quick start:")
    print("  python test_harness.py --list                           # List all tools")
    print("  python test_harness.py --info get_doc_content           # Get tool info")
    print("  python test_harness.py --interactive                    # Interactive mode")
    print("  python test_harness.py --tool search_docs --query 'test' --user_google_email user@example.com")


if __name__ == "__main__":
    main()

