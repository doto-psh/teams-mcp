import logging
import os
from typing import Optional
from importlib import metadata

from fastapi.responses import HTMLResponse, JSONResponse
from starlette.applications import Starlette
from starlette.requests import Request
from starlette.middleware import Middleware

from fastmcp import FastMCP

from auth.oauth21_session_store import get_oauth21_session_store
from auth.teams_auth import handle_auth_callback, start_auth_flow, check_client_secrets
from auth.mcp_session_middleware import MCPSessionMiddleware
from auth.oauth_responses import create_error_response, create_success_response, create_server_error_response
from auth.auth_info_middleware import AuthInfoMiddleware
from auth.scopes import SCOPES
from core.config import (
    USER_MICROSOFT_EMAIL,
    get_transport_mode,
    set_transport_mode as _set_transport_mode,
    get_oauth_redirect_uri as get_oauth_redirect_uri_for_current_mode,
)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# --- Middleware Definitions ---
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

# --- Server Instance ---
server = SecureFastMCP(
    name="microsoft_teams",
    auth=None,
)

# Add the AuthInfo middleware to inject authentication into FastMCP context
auth_info_middleware = AuthInfoMiddleware()
server.add_middleware(auth_info_middleware)

def set_transport_mode(mode: str):
    """Set transport mode for OAuth callback handling."""
    _set_transport_mode(mode)
    logger.info(f"🔌 Transport: {mode}")

def configure_server_for_http():
    """
    Configures the authentication provider for HTTP transport.
    This must be called BEFORE server.run().
    """
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
            logger.warning("⚠️  OAuth 2.1 enabled but OAuth credentials not configured")
            return
        logger.info("🔐 OAuth 2.1 enabled for Microsoft Teams authentication")
    else:
        logger.info("OAuth 2.0 mode - Server will use legacy authentication.")
        server.auth = None

# --- Custom Routes ---
@server.custom_route("/health", methods=["GET"])
async def health_check(request: Request):
    try:
        version = metadata.version("teams-mcp-server")
    except metadata.PackageNotFoundError:
        version = "dev"
    return JSONResponse({
        "status": "healthy",
        "service": "teams-mcp-server",
        "version": version,
        "transport": get_transport_mode()
    })

@server.custom_route("/oauth2callback", methods=["GET"])
async def oauth2_callback(request: Request) -> HTMLResponse:
    """Handle OAuth 2.0 callback from Microsoft."""
    logger.info("Received OAuth callback from Microsoft")
    
    # Extract authorization code and state from query parameters
    code = request.query_params.get("code")
    state = request.query_params.get("state")
    error = request.query_params.get("error")
    
    if error:
        logger.error(f"OAuth error received: {error}")
        return create_error_response(
            error_message=f"Authentication failed: {error}. {request.query_params.get('error_description', '')}"
        )
    
    if not code:
        logger.error("No authorization code received in callback")
        return create_error_response(
            error_message="No authorization code received. The OAuth callback did not include the required authorization code."
        )
    
    try:
        # Get session ID from context if available
        session_id = None
        try:
            from core.context import get_fastmcp_session_id
            session_id = get_fastmcp_session_id()
        except Exception as e:
            logger.debug(f"Could not get session ID from context: {e}")
        
        # Try to extract session from request headers if not found
        if not session_id:
            try:
                from auth.oauth21_session_store import extract_session_from_headers
                headers = dict(request.headers)
                session_id = extract_session_from_headers(headers)
                logger.debug(f"Extracted session ID from headers: {session_id}")
            except Exception as e:
                logger.debug(f"Could not extract session from headers: {e}")
        
        # Fallback: use state parameter as session identifier
        if not session_id and state:
            session_id = f"oauth_state_{state}"
            logger.debug(f"Using OAuth state as session ID: {session_id}")
        
        # Handle the OAuth callback
        success, message, user_email = await handle_auth_callback(code, state, session_id)
        
        if success and user_email:
            logger.info(f"Successfully authenticated user: {user_email}")
            return create_success_response(
                verified_user_id=user_email
            )
        else:
            logger.error(f"Authentication failed: {message}")
            return create_error_response(
                error_message=message or "Unknown error occurred during authentication"
            )
            
    except Exception as e:
        logger.error(f"Unexpected error during OAuth callback: {e}", exc_info=True)
        return create_server_error_response(
            error_detail=str(e)
        )

@server.custom_route("/start_auth", methods=["POST"])
async def start_teams_auth(request: Request) -> JSONResponse:
    """Start Microsoft Teams authentication flow."""
    try:
        data = await request.json()
        user_email = data.get("user_email")
        
        if not user_email:
            return JSONResponse(
                {"error": "user_email is required"}, 
                status_code=400
            )
        
        # Start the authentication flow
        auth_url, state = await start_auth_flow(user_email)
        
        return JSONResponse({
            "auth_url": auth_url,
            "state": state,
            "message": f"Authentication started for {user_email}"
        })
        
    except Exception as e:
        logger.error(f"Error starting auth flow: {e}")
        return JSONResponse(
            {"error": f"Failed to start authentication: {str(e)}"}, 
            status_code=500
        )

# --- Server Configuration Validation ---
def validate_server_config():
    """Validate server configuration before starting."""
    config_errors = check_client_secrets()
    if config_errors:
        logger.error(f"Server configuration errors: {config_errors}")
        return False
    return True