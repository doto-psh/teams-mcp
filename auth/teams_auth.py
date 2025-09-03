# auth/teams_auth.py

import asyncio
import json
import jwt
import logging
import os
import httpx
from datetime import datetime, timedelta
from typing import List, Optional, Tuple, Dict, Any

import msal
from auth.scopes import SCOPES
from auth.oauth21_session_store import get_oauth21_session_store
from core.config import (
    TEAMS_MCP_PORT,
    TEAMS_MCP_BASE_URI,
    get_transport_mode,
    get_oauth_redirect_uri,
)
from core.context import get_fastmcp_session_id

# Try to import FastMCP dependencies (may not be available in all environments)
try:
    from fastmcp.server.dependencies import get_context as get_fastmcp_context
except ImportError:
    get_fastmcp_context = None

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


# Constants
def get_default_credentials_dir():
    """Get the default credentials directory path, preferring user-specific locations."""
    # Check for explicit environment variable override
    if os.getenv("MICROSOFT_MCP_CREDENTIALS_DIR"):
        return os.getenv("MICROSOFT_MCP_CREDENTIALS_DIR")

    # Use user home directory for credentials storage
    home_dir = os.path.expanduser("~")
    if home_dir and home_dir != "~":  # Valid home directory found
        return os.path.join(home_dir, ".microsoft_teams_mcp", "credentials")

    # Fallback to current working directory if home directory is not accessible
    return os.path.join(os.getcwd(), ".credentials")


DEFAULT_CREDENTIALS_DIR = get_default_credentials_dir()

# Microsoft OAuth Configuration
MICROSOFT_AUTHORITY = "https://login.microsoftonline.com/"
MICROSOFT_GRAPH_ENDPOINT = "https://graph.microsoft.com/v1.0"


class TeamsCredentials:
    """Microsoft Teams credentials wrapper similar to Google Credentials."""
    
    def __init__(self, token=None, refresh_token=None, token_uri=None, 
                 client_id=None, client_secret=None, scopes=None, expiry=None,
                 tenant_id=None):
        self.token = token
        self.refresh_token = refresh_token
        self.token_uri = token_uri
        self.client_id = client_id
        self.client_secret = client_secret
        self.scopes = scopes or []
        self.expiry = expiry
        self.tenant_id = tenant_id

    @property
    def expired(self):
        """Check if the token is expired."""
        if not self.expiry:
            return True
        return datetime.now() >= self.expiry

    @property
    def valid(self):
        """Check if credentials are valid."""
        return self.token and not self.expired

    def refresh(self, request=None):
        """Refresh the access token using the refresh token."""
        if not self.refresh_token:
            raise Exception("No refresh token available")
        
        try:
            app = msal.ConfidentialClientApplication(
                client_id=self.client_id,
                client_credential=self.client_secret,
                authority=f"{MICROSOFT_AUTHORITY}{self.tenant_id}"
            )
            
            result = app.acquire_token_by_refresh_token(
                refresh_token=self.refresh_token,
                scopes=self.scopes
            )
            
            if "access_token" in result:
                self.token = result["access_token"]
                if "expires_in" in result:
                    self.expiry = datetime.now() + timedelta(seconds=result["expires_in"])
                if "refresh_token" in result:
                    self.refresh_token = result["refresh_token"]
                logger.info("Successfully refreshed Microsoft Teams credentials")
            else:
                error_msg = result.get("error_description", result.get("error", "Unknown error"))
                raise Exception(f"Failed to refresh token: {error_msg}")
                
        except Exception as e:
            logger.error(f"Error refreshing Microsoft Teams credentials: {e}")
            raise


def _find_any_credentials(
    base_dir: str = DEFAULT_CREDENTIALS_DIR,
) -> Optional[TeamsCredentials]:
    """
    Find and load any valid credentials from the credentials directory.
    Used in single-user mode to bypass session-to-OAuth mapping.

    Returns:
        First valid TeamsCredentials object found, or None if none exist.
    """
    if not os.path.exists(base_dir):
        logger.info(f"[single-user] Credentials directory not found: {base_dir}")
        return None

    # Scan for any .json credential files
    for filename in os.listdir(base_dir):
        if filename.endswith(".json"):
            filepath = os.path.join(base_dir, filename)
            try:
                with open(filepath, "r") as f:
                    creds_data = json.load(f)
                
                expiry = None
                if creds_data.get("expiry"):
                    expiry = datetime.fromisoformat(creds_data["expiry"])
                
                credentials = TeamsCredentials(
                    token=creds_data.get("token"),
                    refresh_token=creds_data.get("refresh_token"),
                    token_uri=creds_data.get("token_uri"),
                    client_id=creds_data.get("client_id"),
                    client_secret=creds_data.get("client_secret"),
                    scopes=creds_data.get("scopes"),
                    expiry=expiry,
                    tenant_id=creds_data.get("tenant_id")
                )
                logger.info(f"[single-user] Found credentials in {filepath}")
                return credentials
            except (IOError, json.JSONDecodeError, KeyError) as e:
                logger.warning(
                    f"[single-user] Error loading credentials from {filepath}: {e}"
                )
                continue

    logger.info(f"[single-user] No valid credentials found in {base_dir}")
    return None


def _get_user_credential_path(
    user_email: str, base_dir: str = DEFAULT_CREDENTIALS_DIR
) -> str:
    """Constructs the path to a user's credential file."""
    if not os.path.exists(base_dir):
        os.makedirs(base_dir)
        logger.info(f"Created credentials directory: {base_dir}")
    return os.path.join(base_dir, f"{user_email}.json")


def save_credentials_to_file(
    user_email: str,
    credentials: TeamsCredentials,
    base_dir: str = DEFAULT_CREDENTIALS_DIR,
):
    """Saves user credentials to a file."""
    creds_path = _get_user_credential_path(user_email, base_dir)
    creds_data = {
        "token": credentials.token,
        "refresh_token": credentials.refresh_token,
        "token_uri": credentials.token_uri,
        "client_id": credentials.client_id,
        "client_secret": credentials.client_secret,
        "scopes": credentials.scopes,
        "expiry": credentials.expiry.isoformat() if credentials.expiry else None,
        "tenant_id": credentials.tenant_id,
    }
    try:
        with open(creds_path, "w") as f:
            json.dump(creds_data, f)
        logger.info(f"Credentials saved for user {user_email} to {creds_path}")
    except IOError as e:
        logger.error(
            f"Error saving credentials for user {user_email} to {creds_path}: {e}"
        )
        raise


def load_credentials_from_file(
    user_email: str, base_dir: str = DEFAULT_CREDENTIALS_DIR
) -> Optional[TeamsCredentials]:
    """Loads user credentials from a file."""
    creds_path = _get_user_credential_path(user_email, base_dir)
    if not os.path.exists(creds_path):
        logger.info(f"No credentials file found for user {user_email}")
        return None

    try:
        with open(creds_path, "r") as f:
            creds_data = json.load(f)
        
        expiry = None
        if creds_data.get("expiry"):
            expiry = datetime.fromisoformat(creds_data["expiry"])
        
        credentials = TeamsCredentials(
            token=creds_data.get("token"),
            refresh_token=creds_data.get("refresh_token"),
            token_uri=creds_data.get("token_uri"),
            client_id=creds_data.get("client_id"),
            client_secret=creds_data.get("client_secret"),
            scopes=creds_data.get("scopes"),
            expiry=expiry,
            tenant_id=creds_data.get("tenant_id")
        )
        logger.info(f"Credentials loaded for user {user_email}")
        return credentials
    except (IOError, json.JSONDecodeError, KeyError) as e:
        logger.error(f"Error loading credentials for user {user_email}: {e}")
        return None


def create_msal_app(tenant_id: str = None) -> msal.ConfidentialClientApplication:
    """Create MSAL application instance."""
    client_id = os.getenv("MICROSOFT_OAUTH_CLIENT_ID")
    client_secret = os.getenv("MICROSOFT_OAUTH_CLIENT_SECRET")
    tenant_id = tenant_id or os.getenv("MICROSOFT_TENANT_ID", "common")
    
    if not client_id or not client_secret:
        raise ValueError("Microsoft OAuth credentials not configured. Please set MICROSOFT_OAUTH_CLIENT_ID and MICROSOFT_OAUTH_CLIENT_SECRET")
    
    authority = f"{MICROSOFT_AUTHORITY}{tenant_id}"
    
    return msal.ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=authority
    )


def get_authorization_url() -> Tuple[str, str]:
    """Get authorization URL for Microsoft OAuth flow."""
    app = create_msal_app()
    redirect_uri = get_oauth_redirect_uri()
    
    # Generate authorization URL with PKCE
    auth_url = app.get_authorization_request_url(
        scopes=SCOPES,
        redirect_uri=redirect_uri,
        response_type="code"
    )
    
    # Return URL and state (MSAL handles state internally)
    return auth_url, ""


def exchange_code_for_credentials(code: str, state: str = None) -> TeamsCredentials:
    """Exchange authorization code for credentials."""
    app = create_msal_app()
    redirect_uri = get_oauth_redirect_uri()
    
    result = app.acquire_token_by_authorization_code(
        code=code,
        scopes=SCOPES,
        redirect_uri=redirect_uri
    )
    
    if "access_token" not in result:
        error_msg = result.get("error_description", result.get("error", "Unknown error"))
        raise Exception(f"Failed to exchange code for token: {error_msg}")
    
    expiry = None
    if "expires_in" in result:
        expiry = datetime.now() + timedelta(seconds=result["expires_in"])
    
    tenant_id = os.getenv("MICROSOFT_TENANT_ID", "common")
    
    credentials = TeamsCredentials(
        token=result["access_token"],
        refresh_token=result.get("refresh_token"),
        token_uri=f"{MICROSOFT_AUTHORITY}{tenant_id}/oauth2/v2.0/token",
        client_id=os.getenv("MICROSOFT_OAUTH_CLIENT_ID"),
        client_secret=os.getenv("MICROSOFT_OAUTH_CLIENT_SECRET"),
        scopes=result.get("scope", SCOPES),
        expiry=expiry,
        tenant_id=tenant_id
    )
    
    return credentials


async def get_user_info(credentials: TeamsCredentials) -> Dict[str, Any]:
    """Get user information from Microsoft Graph API."""
    if not credentials.valid:
        if credentials.refresh_token:
            credentials.refresh()
        else:
            raise Exception("Invalid credentials and no refresh token available")
    
    headers = {
        "Authorization": f"Bearer {credentials.token}",
        "Content-Type": "application/json"
    }
    
    async with httpx.AsyncClient() as client:
        response = await client.get(
            f"{MICROSOFT_GRAPH_ENDPOINT}/me",
            headers=headers
        )
        
        if response.status_code == 200:
            return response.json()
        else:
            raise Exception(f"Failed to get user info: {response.status_code} {response.text}")


def get_cached_user_credentials(user_email: str) -> Optional[TeamsCredentials]:
    """
    Get cached credentials for a user.
    
    Args:
        user_email: User's email address
        
    Returns:
        TeamsCredentials object if found and valid, None otherwise
    """
    # Check if running in single-user mode
    if os.getenv('MCP_SINGLE_USER_MODE') == '1':
        logger.info("[single-user] Loading any available credentials")
        return _find_any_credentials()
    
    # Multi-user mode: load user-specific credentials
    credentials = load_credentials_from_file(user_email)
    if credentials and credentials.valid:
        return credentials
    elif credentials and credentials.refresh_token:
        try:
            credentials.refresh()
            save_credentials_to_file(user_email, credentials)
            return credentials
        except Exception as e:
            logger.error(f"Failed to refresh credentials for {user_email}: {e}")
            return None
    
    return None


# Backward compatibility functions
def get_oauth_flow():
    """Create OAuth flow for Microsoft Teams authentication."""
    return create_msal_app()


def credentials_to_dict(credentials: TeamsCredentials) -> Dict[str, Any]:
    """Convert TeamsCredentials to dictionary."""
    return {
        "token": credentials.token,
        "refresh_token": credentials.refresh_token,
        "token_uri": credentials.token_uri,
        "client_id": credentials.client_id,
        "client_secret": credentials.client_secret,
        "scopes": credentials.scopes,
        "expiry": credentials.expiry.isoformat() if credentials.expiry else None,
        "tenant_id": credentials.tenant_id,
    }


def credentials_from_dict(data: Dict[str, Any]) -> TeamsCredentials:
    """Create TeamsCredentials from dictionary."""
    expiry = None
    if data.get("expiry"):
        expiry = datetime.fromisoformat(data["expiry"])
    
    return TeamsCredentials(
        token=data.get("token"),
        refresh_token=data.get("refresh_token"),
        token_uri=data.get("token_uri"),
        client_id=data.get("client_id"),
        client_secret=data.get("client_secret"),
        scopes=data.get("scopes"),
        expiry=expiry,
        tenant_id=data.get("tenant_id")
    )


# Session management functions
async def save_credentials_to_session(session_id: str, credentials: TeamsCredentials):
    """Saves user credentials using OAuth21SessionStore."""
    # Get user email from credentials if possible
    user_email = None
    if credentials and credentials.token:
        try:
            # Use Microsoft Graph API to get user info
            user_info = await get_user_info(credentials)
            user_email = user_info.get("mail") or user_info.get("userPrincipalName")
        except Exception as e:
            logger.debug(f"Could not get user email from credentials: {e}")
    
    if user_email:
        store = get_oauth21_session_store()
        store.store_session(
            user_email=user_email,
            access_token=credentials.token,
            refresh_token=credentials.refresh_token,
            token_uri=credentials.token_uri,
            client_id=credentials.client_id,
            client_secret=credentials.client_secret,
            scopes=credentials.scopes,
            expiry=credentials.expiry,
            mcp_session_id=session_id
        )
        logger.debug(f"Credentials saved to OAuth21SessionStore for session_id: {session_id}, user: {user_email}")
    else:
        logger.warning(f"Could not save credentials to session store - no user email found for session: {session_id}")


def load_credentials_from_session(session_id: str) -> Optional[TeamsCredentials]:
    """Loads user credentials from OAuth21SessionStore."""
    store = get_oauth21_session_store()
    credentials = store.get_credentials_by_mcp_session(session_id)
    if credentials:
        logger.debug(
            f"Credentials loaded from OAuth21SessionStore for session_id: {session_id}"
        )
        # Convert to TeamsCredentials
        return TeamsCredentials(
            token=credentials.token,
            refresh_token=credentials.refresh_token,
            token_uri=credentials.token_uri,
            client_id=credentials.client_id,
            client_secret=credentials.client_secret,
            scopes=credentials.scopes,
            expiry=credentials.expiry,
            tenant_id=os.getenv("MICROSOFT_TENANT_ID", "common")
        )
    else:
        logger.debug(
            f"No credentials found in OAuth21SessionStore for session_id: {session_id}"
        )
    return None


def load_client_secrets_from_env() -> Optional[Dict[str, Any]]:
    """
    Loads the client secrets from environment variables.

    Environment variables used:
        - MICROSOFT_OAUTH_CLIENT_ID: OAuth 2.0 client ID
        - MICROSOFT_OAUTH_CLIENT_SECRET: OAuth 2.0 client secret
        - MICROSOFT_TENANT_ID: Microsoft tenant ID

    Returns:
        Client secrets configuration dict compatible with Microsoft OAuth library,
        or None if required environment variables are not set.
    """
    client_id = os.getenv("MICROSOFT_OAUTH_CLIENT_ID")
    client_secret = os.getenv("MICROSOFT_OAUTH_CLIENT_SECRET")
    tenant_id = os.getenv("MICROSOFT_TENANT_ID", "common")

    if client_id and client_secret:
        # Create config structure for Microsoft OAuth
        config = {
            "client_id": client_id,
            "client_secret": client_secret,
            "authority": f"{MICROSOFT_AUTHORITY}{tenant_id}",
            "tenant_id": tenant_id
        }

        logger.info("Loaded Microsoft OAuth client credentials from environment variables")
        return config

    logger.debug("Microsoft OAuth client credentials not found in environment variables")
    return None


def check_client_secrets() -> Optional[str]:
    """
    Check if Microsoft OAuth client secrets are properly configured.
    
    Returns:
        None if configuration is valid, error message string if invalid
    """
    # Try environment variables first
    env_config = load_client_secrets_from_env()
    if env_config:
        return None  # Environment variables are valid
    
    # Return error message if no valid configuration found
    return "Microsoft OAuth credentials not configured. Please set MICROSOFT_OAUTH_CLIENT_ID and MICROSOFT_OAUTH_CLIENT_SECRET environment variables."


async def start_auth_flow(user_email: str) -> Tuple[str, str]:
    """
    Start Microsoft OAuth authentication flow.
    
    Args:
        user_email: User's email address for authentication
        
    Returns:
        Tuple of (authorization_url, state)
    """
    try:
        app = create_msal_app()
        redirect_uri = get_oauth_redirect_uri()
        
        # Generate authorization URL
        auth_url = app.get_authorization_request_url(
            scopes=SCOPES,
            redirect_uri=redirect_uri,
            response_type="code",
            login_hint=user_email
        )
        
        logger.info(f"Generated auth URL for user {user_email}")
        return auth_url, ""  # MSAL handles state internally
        
    except Exception as e:
        logger.error(f"Error starting auth flow for {user_email}: {e}")
        raise


async def handle_auth_callback(code: str, state: Optional[str] = None, session_id: Optional[str] = None) -> Tuple[bool, str, Optional[str]]:
    """
    Handle OAuth callback and exchange code for credentials.
    
    Args:
        code: Authorization code from callback
        state: State parameter (optional for Microsoft OAuth)
        session_id: MCP session ID
        
    Returns:
        Tuple of (success, message, user_email)
    """
    try:
        # Exchange code for credentials
        credentials = exchange_code_for_credentials(code, state or "")
        
        # Get user info
        user_info = await get_user_info(credentials)
        user_email = user_info.get("mail") or user_info.get("userPrincipalName")
        
        if not user_email:
            return False, "Failed to get user email from Microsoft Graph", None
        
        # Save credentials to file
        save_credentials_to_file(user_email, credentials)
        
        # Save to OAuth21 session store with proper binding
        from auth.oauth21_session_store import get_oauth21_session_store
        store = get_oauth21_session_store()
        
        # Create session info
        expiry = None
        if hasattr(credentials, 'expiry') and credentials.expiry:
            expiry = credentials.expiry
        
        # Store in OAuth21 session with session binding
        store.store_session(
            user_email=user_email,
            access_token=credentials.token,
            refresh_token=credentials.refresh_token,
            token_uri=credentials.token_uri,
            client_id=credentials.client_id,
            client_secret=credentials.client_secret,
            scopes=credentials.scopes,
            expiry=expiry,
            session_id=state or f"oauth_{user_email}",  # Use OAuth state as session ID
            mcp_session_id=session_id,  # Bind to MCP session if provided
            issuer=f"https://login.microsoftonline.com/{getattr(credentials, 'tenant_id', 'common')}/v2.0"
        )
        
        # Also bind session directly for fallback
        if session_id:
            logger.info(f"Binding session {session_id} to user {user_email}")
            # Create additional binding entry
            store._session_auth_binding[session_id] = user_email
        
        # Save to legacy session if session_id provided
        if session_id:
            await save_credentials_to_session(session_id, credentials)
        
        logger.info(f"Successfully authenticated user {user_email}")
        return True, f"Successfully authenticated {user_email}", user_email
        
    except Exception as e:
        logger.error(f"Error handling auth callback: {e}")
        return False, f"Authentication failed: {str(e)}", None


def get_credentials(user_email: str = None, session_id: str = None) -> Optional[TeamsCredentials]:
    """
    Get user credentials from various sources.
    
    Args:
        user_email: User's email address
        session_id: MCP session ID
        
    Returns:
        TeamsCredentials object if found, None otherwise
    """
    # Check if running in single-user mode
    if os.getenv('MCP_SINGLE_USER_MODE') == '1':
        logger.info("[single-user] Loading any available credentials")
        return _find_any_credentials()
    
    # Try session-based credentials first
    if session_id:
        credentials = load_credentials_from_session(session_id)
        if credentials and credentials.valid:
            return credentials
    
    # Try user-specific credentials
    if user_email:
        credentials = load_credentials_from_file(user_email)
        if credentials and credentials.valid:
            return credentials
        elif credentials and credentials.refresh_token:
            try:
                credentials.refresh()
                save_credentials_to_file(user_email, credentials)
                return credentials
            except Exception as e:
                logger.error(f"Failed to refresh credentials for {user_email}: {e}")
    
    return None


class TeamsAuthenticationError(Exception):
    """Exception raised for Microsoft Teams authentication errors."""
    pass


async def get_authenticated_teams_service(user_email: str) -> TeamsCredentials:
    """
    Get authenticated Microsoft Teams credentials for a user.
    
    Args:
        user_email: User's email address
        
    Returns:
        TeamsCredentials object
        
    Raises:
        TeamsAuthenticationError: If authentication fails
    """
    session_id = get_fastmcp_session_id()
    credentials = get_credentials(user_email=user_email, session_id=session_id)
    
    if not credentials:
        raise TeamsAuthenticationError(f"No valid credentials found for user {user_email}")
    
    if not credentials.valid:
        if credentials.refresh_token:
            try:
                credentials.refresh()
                save_credentials_to_file(user_email, credentials)
            except Exception as e:
                raise TeamsAuthenticationError(f"Failed to refresh credentials for {user_email}: {e}")
        else:
            raise TeamsAuthenticationError(f"Invalid credentials and no refresh token for {user_email}")
    
    return credentials


# Helper function for making authenticated requests
async def make_graph_request(
    credentials: TeamsCredentials, 
    endpoint: str, 
    method: str = "GET", 
    data: Dict[str, Any] = None
) -> Dict[str, Any]:
    """Make authenticated request to Microsoft Graph API."""
    if not credentials.valid:
        if credentials.refresh_token:
            credentials.refresh()
        else:
            raise Exception("Invalid credentials and no refresh token available")
    
    headers = {
        "Authorization": f"Bearer {credentials.token}",
        "Content-Type": "application/json"
    }
    
    url = f"{MICROSOFT_GRAPH_ENDPOINT}/{endpoint.lstrip('/')}"
    
    async with httpx.AsyncClient() as client:
        if method.upper() == "GET":
            response = await client.get(url, headers=headers)
        elif method.upper() == "POST":
            response = await client.post(url, headers=headers, json=data)
        elif method.upper() == "PUT":
            response = await client.put(url, headers=headers, json=data)
        elif method.upper() == "DELETE":
            response = await client.delete(url, headers=headers)
        else:
            raise ValueError(f"Unsupported HTTP method: {method}")
        
        if response.status_code in [200, 201, 204]:
            if response.content:
                return response.json()
            return {}
        else:
            raise Exception(f"Graph API request failed: {response.status_code} {response.text}")
