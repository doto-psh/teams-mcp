#!/usr/bin/env python3
"""
Microsoft Teams MCP Server

A Model Context Protocol server for Microsoft Teams integration.
Adapted    # Import tool modules to register them with the MCP server via decorators
    tool_imports = {
        'teams': lambda: __import__('teams'),
    }

    tool_icons = {
        'teams': '🫸',
    }ogle Workspace MCP for Microsoft Graph API.
"""

import argparse
import logging
import os
import sys
from importlib import metadata
from dotenv import load_dotenv

from auth.oauth_config import reload_oauth_config
from auth.scopes import set_enabled_tools
from core.server import server, configure_server_for_http
from core.config import set_transport_mode

# Load environment variables
dotenv_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env')
if not os.path.exists(dotenv_path):
    dotenv_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env.oauth21')
load_dotenv(dotenv_path=dotenv_path)

# Suppress httpx debug logs
logging.getLogger('httpx').setLevel(logging.WARNING)

reload_oauth_config()

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

try:
    root_logger = logging.getLogger()
    log_file_dir = os.path.dirname(os.path.abspath(__file__))
    log_file_path = os.path.join(log_file_dir, 'teams_mcp_server_debug.log')

    file_handler = logging.FileHandler(log_file_path, mode='a')
    file_handler.setLevel(logging.DEBUG)

    file_formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(process)d - %(threadName)s '
        '[%(module)s.%(funcName)s:%(lineno)d] - %(message)s'
    )
    file_handler.setFormatter(file_formatter)
    root_logger.addHandler(file_handler)

    logger.debug(f"Detailed file logging configured to: {log_file_path}")
except Exception as e:
    sys.stderr.write(f"CRITICAL: Failed to set up file logging to '{log_file_path}': {e}\n")

def safe_print(text):
    """Print safely when running as MCP server."""
    if not sys.stderr.isatty():
        logger.debug(f"[MCP Server] {text}")
        return

    try:
        print(text, file=sys.stderr)
    except UnicodeEncodeError:
        print(text.encode('ascii', errors='replace').decode(), file=sys.stderr)

def main():
    """
    Main entry point for the Microsoft Teams MCP server.
    Uses FastMCP's native streamable-http transport.
    """
    # Parse command line arguments
    parser = argparse.ArgumentParser(description='Microsoft Teams MCP Server')
    parser.add_argument('--single-user', action='store_true',
                        help='Run in single-user mode - bypass session mapping and use any credentials from the credentials directory')
    parser.add_argument('--tools', nargs='*',
                        choices=['teams'],
                        help='Specify which tools to register. If not provided, all tools are registered.')
    parser.add_argument('--transport', choices=['stdio', 'streamable-http'], default='stdio',
                        help='Transport mode: stdio (default) or streamable-http')
    parser.add_argument('--port', type=int, default=None,
                        help='Port to run the server on (for streamable-http transport)')
    args = parser.parse_args()

    # Set port and base URI once for reuse throughout the function
    port = args.port or int(os.getenv("PORT", os.getenv("TEAMS_MCP_PORT", 8000)))
    base_uri = os.getenv("TEAMS_MCP_BASE_URI", "http://localhost")

    safe_print("🔧 Microsoft Teams MCP Server")
    safe_print("=" * 35)
    safe_print("📋 Server Information:")
    try:
        version = metadata.version("teams-mcp-server")
    except metadata.PackageNotFoundError:
        version = "dev"
    safe_print(f"   📦 Version: {version}")
    safe_print(f"   🌐 Transport: {args.transport}")
    if args.transport == 'streamable-http':
        safe_print(f"   🔗 URL: {base_uri}:{port}")
        safe_print(f"   🔐 OAuth Callback: {base_uri}:{port}/callback")
    safe_print(f"   👤 Mode: {'Single-user' if args.single_user else 'Multi-user'}")
    safe_print(f"   🐍 Python: {sys.version.split()[0]}")
    safe_print("")

    # Active Configuration
    safe_print("⚙️ Active Configuration:")


    # Redact client secret for security
    client_secret = os.getenv('MICROSOFT_OAUTH_CLIENT_SECRET', 'Not Set')
    redacted_secret = f"{client_secret[:4]}...{client_secret[-4:]}" if len(client_secret) > 8 else "Invalid or too short"

    config_vars = {
        "MICROSOFT_OAUTH_CLIENT_ID": os.getenv('MICROSOFT_OAUTH_CLIENT_ID', 'Not Set'),
        "MICROSOFT_OAUTH_CLIENT_SECRET": redacted_secret,
        "MICROSOFT_TENANT_ID": os.getenv('MICROSOFT_TENANT_ID', 'Not Set'),
        "MCP_SINGLE_USER_MODE": os.getenv('MCP_SINGLE_USER_MODE', 'false'),
        "MCP_ENABLE_OAUTH21": os.getenv('MCP_ENABLE_OAUTH21', 'false'),
        "OAUTHLIB_INSECURE_TRANSPORT": os.getenv('OAUTHLIB_INSECURE_TRANSPORT', 'false'),
    }

    for key, value in config_vars.items():
        safe_print(f"   - {key}: {value}")
    safe_print("")


    # Import tool modules to register them with the MCP server via decorators
    tool_imports = {
        'teams': lambda: __import__('teams.teams_tools'),
    }

    tool_icons = {
        'teams': '�',
    }

    # Import specified tools or all tools if none specified
    tools_to_import = args.tools if args.tools is not None else tool_imports.keys()

    # Set enabled tools for scope management
    from auth.scopes import set_enabled_tools
    set_enabled_tools(list(tools_to_import))

    safe_print(f"🛠️  Loading {len(tools_to_import)} tool module{'s' if len(tools_to_import) != 1 else ''}:")
    for tool in tools_to_import:
        tool_imports[tool]()
        safe_print(f"   {tool_icons[tool]} {tool.title()} - Microsoft {tool.title()} API integration")
        safe_print(f"     ✅ Teams Tools: Basic team operations (list, channels, messages)")
        safe_print(f"     ✅ Chat Tools: Direct chat management (1:1, group chats)")
        safe_print(f"     ✅ Search Tools: Advanced search across Teams (messages, mentions)")
        safe_print(f"     ✅ Users Tools: User management and directory search")
    safe_print("")

    safe_print("📊 Configuration Summary:")
    safe_print(f"   🔧 Tools Enabled: {len(tools_to_import)}/{len(tool_imports)}")
    safe_print(f"   📝 Log Level: {logging.getLogger().getEffectiveLevel()}")
    safe_print("")

    # Set global single-user mode flag
    if args.single_user:
        os.environ['MCP_SINGLE_USER_MODE'] = '1'
        safe_print("🔐 Single-user mode enabled")
        safe_print("")

    try:
        # Set transport mode for OAuth callback handling
        set_transport_mode(args.transport)

        # Configure auth initialization for FastMCP lifecycle events
        if args.transport == 'streamable-http':
            configure_server_for_http()
            safe_print("")
            safe_print(f"🚀 Starting HTTP server on {base_uri}:{port}")
        else:
            safe_print("")
            safe_print("🚀 Starting STDIO server")
            # Start minimal OAuth callback server for stdio mode
            from auth.oauth_callback_server import ensure_oauth_callback_available
            success, error_msg = ensure_oauth_callback_available('stdio', port, base_uri)
            if success:
                safe_print(f"   OAuth callback server started on {base_uri}:{port}/callback")
            else:
                warning_msg = "   ⚠️  Warning: Failed to start OAuth callback server"
                if error_msg:
                    warning_msg += f": {error_msg}"
                safe_print(warning_msg)

        safe_print("✅ Ready for MCP connections")
        safe_print("")

        if args.transport == 'streamable-http':
            # The server has CORS middleware built-in via CORSEnabledFastMCP
            server.run(transport="streamable-http", host="0.0.0.0", port=port)
        else:
            server.run()
    except KeyboardInterrupt:
        safe_print("\n👋 Server shutdown requested")
        # Clean up OAuth callback server if running
        from auth.oauth_callback_server import cleanup_oauth_callback_server
        cleanup_oauth_callback_server()
        sys.exit(0)
    except Exception as e:
        safe_print(f"\n❌ Server error: {e}")
        logger.error(f"Unexpected error running server: {e}", exc_info=True)
        # Clean up OAuth callback server if running
        from auth.oauth_callback_server import cleanup_oauth_callback_server
        cleanup_oauth_callback_server()
        sys.exit(1)

if __name__ == "__main__":
    main()
