import inspect
import logging
from functools import wraps
from typing import Dict, List, Optional, Any, Callable, Union
from datetime import datetime, timedelta
import httpx

from auth.scopes import (
    TEAMS_READ_SCOPE, TEAMS_CHANNELS_READ_SCOPE, TEAMS_MESSAGES_READ_SCOPE,
    TEAMS_CHAT_READ_SCOPE, TEAMS_MEMBERS_READ_SCOPE, USER_READ_SCOPE
)

logger = logging.getLogger(__name__)

# OAuth 2.1 integration is available
OAUTH21_INTEGRATION_AVAILABLE = True

class TeamsAuthenticationError(Exception):
    """Exception raised when Teams authentication fails."""
    pass

class TeamsGraphService:
    """Microsoft Graph API service for Teams operations."""
    
    def __init__(self, access_token: str):
        self.access_token = access_token
        self.base_url = "https://graph.microsoft.com/v1.0"
        self.headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
    
    async def get(self, endpoint: str) -> Dict[str, Any]:
        """Make GET request to Microsoft Graph API."""
        async with httpx.AsyncClient() as client:
            response = await client.get(f"{self.base_url}{endpoint}", headers=self.headers)
            response.raise_for_status()
            return response.json()
    
    async def post(self, endpoint: str, data: Dict[str, Any]) -> Dict[str, Any]:
        """Make POST request to Microsoft Graph API."""
        async with httpx.AsyncClient() as client:
            response = await client.post(f"{self.base_url}{endpoint}", json=data, headers=self.headers)
            response.raise_for_status()
            return response.json()

async def get_authenticated_teams_service_oauth21(
    tool_name: str,
    user_email: str,
    required_scopes: List[str],
    session_id: Optional[str] = None,
    auth_token_email: Optional[str] = None,
    allow_recent_auth: bool = False,
) -> tuple[TeamsGraphService, str]:
    """
    OAuth 2.1 authentication for Microsoft Teams using the session store.
    """
    from auth.oauth21_session_store import get_oauth21_session_store

    store = get_oauth21_session_store()

    # Use validation method to ensure session can only access its own credentials
    credentials = store.get_credentials_with_validation(
        requested_user_email=user_email,
        session_id=session_id,
        auth_token_email=auth_token_email,
        allow_recent_auth=allow_recent_auth
    )

    if not credentials:
        raise TeamsAuthenticationError(
            f"Access denied: Cannot retrieve credentials for {user_email}. "
            f"You can only access credentials for your authenticated account."
        )

    # Check scopes (simplified for Microsoft Graph)
    # Note: Microsoft Graph API scopes are checked at token level
    
    # Create Teams Graph service
    service = TeamsGraphService(credentials.token)
    logger.info(f"[{tool_name}] Authenticated Teams service for {user_email}")

    return service, user_email

# Service configuration mapping for Teams
SERVICE_CONFIGS = {
    "teams": {"scopes": ["teams_read"]},
    "chat": {"scopes": ["chat_read"]},
    "user": {"scopes": ["user_read"]}
}

# Scope group definitions for Microsoft Teams
SCOPE_GROUPS = {
    # Teams scopes
    "teams_read": TEAMS_READ_SCOPE,
    "teams_channels": TEAMS_CHANNELS_READ_SCOPE,
    "teams_messages": TEAMS_MESSAGES_READ_SCOPE,
    "teams_chat": TEAMS_CHAT_READ_SCOPE,
    "teams_members": TEAMS_MEMBERS_READ_SCOPE,
    
    # User scopes
    "user_read": USER_READ_SCOPE,
}

# Service cache for performance
_service_cache: Dict[str, tuple[TeamsGraphService, datetime, str]] = {}
_cache_duration = timedelta(minutes=30)

def _is_cache_valid(cached_time: datetime) -> bool:
    """Check if cached service is still valid."""
    return datetime.now() - cached_time < _cache_duration

def _get_cache_key(user_email: str, service_type: str, resolved_scopes: List[str]) -> str:
    """Generate cache key for service."""
    scope_str = "_".join(sorted(resolved_scopes))
    return f"{user_email}_{service_type}_{scope_str}"

def _get_cached_service(cache_key: str) -> Optional[tuple[TeamsGraphService, str]]:
    """Retrieve cached service if valid."""
    if cache_key in _service_cache:
        service, cached_time, user_email = _service_cache[cache_key]
        if _is_cache_valid(cached_time):
            logger.debug(f"Using cached Teams service for key: {cache_key}")
            return service, user_email
        else:
            # Remove expired cache entry
            del _service_cache[cache_key]
            logger.debug(f"Removed expired cache entry: {cache_key}")
    return None

def _cache_service(cache_key: str, service: TeamsGraphService, user_email: str) -> None:
    """Cache a service instance."""
    _service_cache[cache_key] = (service, datetime.now(), user_email)
    logger.debug(f"Cached Teams service for key: {cache_key}")

def _resolve_scopes(scopes: Union[str, List[str]]) -> List[str]:
    """Resolve scope names to actual scope URLs."""
    if isinstance(scopes, str):
        if scopes in SCOPE_GROUPS:
            return [SCOPE_GROUPS[scopes]]
        else:
            return [scopes]

    resolved = []
    for scope in scopes:
        if scope in SCOPE_GROUPS:
            resolved.append(SCOPE_GROUPS[scope])
        else:
            resolved.append(scope)
    return resolved

def require_teams_service(service_type: str, scopes: Union[str, List[str]]):
    """
    Decorator for functions that need Microsoft Teams Graph API service.

    Args:
        service_type: Type of service (e.g., "teams", "chat")
        scopes: Required Microsoft Graph API scopes

    Usage:
        @require_teams_service("teams", "teams_read")
        async def list_teams(service, user_email: str):
            # service is automatically injected
            return await service.get("/me/joinedTeams")
    """
    def decorator(func: Callable) -> Callable:
        # Get original function signature
        sig = inspect.signature(func)
        
        # Create new signature without 'service' parameter
        new_params = []
        for name, param in sig.parameters.items():
            if name != 'service':  # Remove service parameter from signature
                new_params.append(param)
        
        wrapper_sig = sig.replace(parameters=new_params)

        @wraps(func)
        async def wrapper(*args, **kwargs):
            # Extract user_email from arguments
            sig = inspect.signature(func)
            param_names = list(sig.parameters.keys())

            user_email = None
            if 'user_email' in kwargs:
                user_email = kwargs['user_email']
            else:
                try:
                    user_email_index = param_names.index('user_email')
                    # Skip the 'service' parameter when counting
                    adjusted_index = user_email_index - 1 if user_email_index > 0 else 0
                    if adjusted_index < len(args):
                        user_email = args[adjusted_index]
                except (ValueError, IndexError):
                    pass

            if not user_email:
                raise TeamsAuthenticationError("user_email parameter is required")

            # Resolve scopes
            resolved_scopes = _resolve_scopes(scopes)
            
            # Try cache first
            service = None
            actual_user_email = user_email
            
            cache_key = _get_cache_key(user_email, service_type, resolved_scopes)
            cached_result = _get_cached_service(cache_key)
            if cached_result:
                service, actual_user_email = cached_result

            if service is None:
                try:
                    tool_name = func.__name__

                    # Get authenticated user from context
                    authenticated_user = None
                    auth_method = None
                    mcp_session_id = None

                    try:
                        from fastmcp.server.dependencies import get_context
                        ctx = get_context()
                        if ctx:
                            authenticated_user = ctx.get_state("authenticated_user_email")
                            auth_method = ctx.get_state("authenticated_via")

                            if hasattr(ctx, 'session_id'):
                                mcp_session_id = ctx.session_id

                            logger.debug(f"[{tool_name}] Auth from middleware: {authenticated_user} via {auth_method}")
                    except Exception as e:
                        logger.debug(f"[{tool_name}] Could not get FastMCP context: {e}")

                    # Log authentication status
                    logger.debug(f"[{tool_name}] Auth: {authenticated_user or 'none'} via {auth_method or 'none'} (session: {mcp_session_id[:8] if mcp_session_id else 'none'})")

                    # Get Teams service
                    service, actual_user_email = await get_authenticated_teams_service_oauth21(
                        tool_name=tool_name,
                        user_email=user_email,
                        required_scopes=resolved_scopes,
                        session_id=mcp_session_id,
                        auth_token_email=authenticated_user,
                        allow_recent_auth=False
                    )

                    # Cache the service
                    _cache_service(cache_key, service, actual_user_email)

                except Exception as e:
                    logger.error(f"[{tool_name}] Failed to get Teams service: {e}")
                    raise TeamsAuthenticationError(f"Failed to authenticate Teams service: {e}")

            # Call the original function with the service object injected
            try:
                return await func(service, *args, **kwargs)
            except Exception as e:
                error_message = f"Teams API error for {actual_user_email}: {str(e)}"
                logger.error(f"[{func.__name__}] {error_message}")
                raise Exception(error_message)

        # Set the wrapper's signature to the one without 'service'
        wrapper.__signature__ = wrapper_sig
        return wrapper
    return decorator

def require_multiple_teams_services(service_configs: List[Dict[str, Any]]):
    """
    Decorator for functions that need multiple Teams services.

    Args:
        service_configs: List of service configurations

    Usage:
        @require_multiple_teams_services([
            {"service_type": "teams", "scopes": "teams_read", "param_name": "teams_service"},
            {"service_type": "chat", "scopes": "chat_read", "param_name": "chat_service"}
        ])
        async def get_teams_and_chats(teams_service, chat_service, user_email: str):
            # Both services are automatically injected
    """
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        async def wrapper(*args, **kwargs):
            # Extract user_email
            sig = inspect.signature(func)
            param_names = list(sig.parameters.keys())

            user_email = None
            if 'user_email' in kwargs:
                user_email = kwargs['user_email']
            else:
                try:
                    user_email_index = param_names.index('user_email')
                    if user_email_index < len(args):
                        user_email = args[user_email_index]
                except ValueError:
                    pass

            if not user_email:
                raise TeamsAuthenticationError("user_email parameter is required")

            # Get all required services
            services = {}
            for config in service_configs:
                service_type = config["service_type"]
                scopes = config["scopes"]
                param_name = config.get("param_name", f"{service_type}_service")
                
                resolved_scopes = _resolve_scopes(scopes)
                
                service, _ = await get_authenticated_teams_service_oauth21(
                    tool_name=func.__name__,
                    user_email=user_email,
                    required_scopes=resolved_scopes,
                    allow_recent_auth=False
                )
                
                services[param_name] = service

            # Inject services into function call
            modified_kwargs = kwargs.copy()
            modified_kwargs.update(services)

            return await func(*args, **modified_kwargs)

        return wrapper
    return decorator
