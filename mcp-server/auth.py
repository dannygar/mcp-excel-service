"""
Microsoft Entra ID Token Validation for MCP Server

This module provides JWT token validation for Microsoft Entra ID (Azure AD)
authentication. It validates bearer tokens sent by AI Foundry agents using
the Project Managed Identity or Agent Identity.

Usage:
    from auth import EntraTokenValidator, EntraAuthMiddleware, create_validator_from_env
    
    validator = create_validator_from_env()
    if validator:
        app_with_auth = EntraAuthMiddleware(app, validator=validator)
"""

import os
import logging
from typing import Optional, Dict, Any, Callable
from functools import wraps

import jwt
from jwt import PyJWKClient, PyJWKClientError
from jwt.exceptions import InvalidTokenError, ExpiredSignatureError, InvalidAudienceError, InvalidIssuerError
from starlette.requests import Request
from starlette.responses import JSONResponse

logger = logging.getLogger("entra-auth")

# =============================================================================
# Configuration (read from environment)
# =============================================================================

# Get configuration for incoming token validation
# Check ENTRA_* first (set by Bicep for token validation), then fall back to AZURE_* (for Graph API)
ENTRA_TENANT_ID = os.getenv("ENTRA_TENANT_ID") or os.getenv("AZURE_TENANT_ID", "")
ENTRA_CLIENT_ID = os.getenv("ENTRA_CLIENT_ID") or os.getenv("AZURE_CLIENT_ID", "")  # This is the audience

# Auth can be explicitly disabled for local development (e.g., MCP Inspector testing)
# Set DISABLE_ENTRA_AUTH=true to skip token validation
_auth_explicitly_disabled = os.getenv("DISABLE_ENTRA_AUTH", "").lower() in ("true", "1", "yes")
_has_entra_config = bool(ENTRA_TENANT_ID and ENTRA_CLIENT_ID)
ENTRA_AUTH_ENABLED = _has_entra_config and not _auth_explicitly_disabled


# =============================================================================
# Token Validator Class
# =============================================================================

class EntraTokenValidator:
    """
    Validates Microsoft Entra ID JWT tokens for Azure AI Foundry integration.
    
    This validator:
    - Fetches and caches JWKS (JSON Web Key Sets) from Microsoft
    - Validates token signature, audience, issuer, and expiration
    - Supports both v1.0 and v2.0 token endpoints
    """
    
    def __init__(
        self,
        tenant_id: str,
        client_id: str,
        allowed_audiences: Optional[list] = None,
        cache_ttl: int = 3600
    ):
        """
        Initialize the token validator.
        
        Args:
            tenant_id: Microsoft Entra tenant ID
            client_id: Application (client) ID of the app registration
            allowed_audiences: List of allowed audience values (defaults to client_id and api://{client_id})
            cache_ttl: Time in seconds to cache JWKS (default: 1 hour)
        """
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.allowed_audiences = allowed_audiences or [
            client_id,
            f"api://{client_id}"
        ]
        self.cache_ttl = cache_ttl
        
        # JWKS endpoint for Microsoft Entra ID
        self.jwks_uri = f"https://login.microsoftonline.com/{tenant_id}/discovery/v2.0/keys"
        
        # Valid issuers - support both v1.0 and v2.0 tokens
        self.issuer_v2 = f"https://login.microsoftonline.com/{tenant_id}/v2.0"
        self.issuer_v1 = f"https://sts.windows.net/{tenant_id}/"
        
        # Initialize PyJWKClient with caching
        self._jwks_client = PyJWKClient(
            self.jwks_uri,
            cache_jwk_set=True,
            lifespan=cache_ttl
        )
    
    async def validate_token(self, token: str) -> tuple[bool, Optional[Dict[str, Any]], Optional[str]]:
        """
        Validate a JWT token from Microsoft Entra ID.
        
        Args:
            token: The JWT bearer token (without 'Bearer ' prefix)
            
        Returns:
            Tuple of (is_valid, claims, error_message)
        """
        try:
            # Log token info for debugging (without exposing sensitive data)
            try:
                unverified_claims = jwt.decode(token, options={"verify_signature": False})
                logger.info(f"Token audience (aud): {unverified_claims.get('aud')}")
                logger.info(f"Token issuer (iss): {unverified_claims.get('iss')}")
                logger.info(f"Token app id (appid/azp): {unverified_claims.get('appid') or unverified_claims.get('azp')}")
                logger.info(f"Allowed audiences: {self.allowed_audiences}")
            except Exception as e:
                logger.warning(f"Could not decode unverified claims: {e}")
            
            # Get the signing key from JWKS using PyJWKClient
            try:
                signing_key = self._jwks_client.get_signing_key_from_jwt(token)
                logger.info(f"Successfully retrieved signing key")
            except PyJWKClientError as e:
                logger.error(f"Failed to get signing key from JWKS: {e}")
                return False, None, f"Unable to find signing key for token: {e}"
            except Exception as e:
                logger.error(f"Unexpected error getting signing key: {e}")
                return False, None, f"Error retrieving signing key: {e}"
            
            # Try v1.0 issuer first (common for managed identities), then v2.0
            claims = None
            last_error = None
            
            for issuer in [self.issuer_v1, self.issuer_v2]:
                try:
                    claims = jwt.decode(
                        token,
                        signing_key.key,
                        algorithms=["RS256"],
                        audience=self.allowed_audiences,
                        issuer=issuer,
                        options={
                            "verify_aud": True,
                            "verify_iss": True,
                            "verify_exp": True,
                            "verify_nbf": True,
                        }
                    )
                    logger.info(f"Token validated successfully with issuer: {issuer}")
                    break
                except InvalidIssuerError as e:
                    last_error = str(e)
                    continue
                except (InvalidTokenError, ExpiredSignatureError, InvalidAudienceError) as e:
                    last_error = str(e)
                    logger.warning(f"Token validation failed with issuer {issuer}: {e}")
                    break  # Don't try other issuers for non-issuer errors
            
            if claims is None:
                return False, None, f"Token validation failed: {last_error}"
            
            # Log successful validation
            logger.info(f"Token validated for app: {claims.get('appid') or claims.get('azp')}")
            
            return True, claims, None
            
        except Exception as e:
            logger.error(f"Token validation error: {e}")
            return False, None, str(e)


# =============================================================================
# Helper Functions
# =============================================================================

def get_token_from_request(request: Request) -> Optional[str]:
    """Extract bearer token from request Authorization header."""
    auth_header = request.headers.get("Authorization", "")
    
    if not auth_header.startswith("Bearer "):
        return None
    
    return auth_header[7:]  # Remove "Bearer " prefix


def require_auth(validator: EntraTokenValidator):
    """
    Decorator factory that creates an authentication middleware.
    
    Usage:
        @require_auth(validator)
        async def my_endpoint(request):
            ...
    """
    def decorator(func: Callable):
        @wraps(func)
        async def wrapper(request: Request, *args, **kwargs):
            token = get_token_from_request(request)
            
            if not token:
                logger.warning("No bearer token in request")
                return JSONResponse(
                    {"error": "unauthorized", "message": "Missing bearer token"},
                    status_code=401
                )
            
            is_valid, claims, error = await validator.validate_token(token)
            
            if not is_valid:
                logger.warning(f"Token validation failed: {error}")
                return JSONResponse(
                    {"error": "unauthorized", "message": error},
                    status_code=401
                )
            
            # Attach claims to request state for downstream use
            request.state.token_claims = claims
            
            return await func(request, *args, **kwargs)
        
        return wrapper
    return decorator


# =============================================================================
# ASGI Middleware
# =============================================================================

class EntraAuthMiddleware:
    """
    ASGI middleware for Microsoft Entra ID authentication.
    
    This middleware validates bearer tokens on all requests except
    for specified excluded paths (like /health).
    
    Note: This is a raw ASGI middleware, not a Starlette BaseHTTPMiddleware,
    which allows it to work correctly with FastMCP's lifespan management.
    """
    
    def __init__(
        self,
        app,
        validator: EntraTokenValidator,
        excluded_paths: Optional[list] = None
    ):
        """
        Initialize the auth middleware.
        
        Args:
            app: The ASGI application to wrap
            validator: EntraTokenValidator instance
            excluded_paths: List of paths to exclude from auth (e.g., ["/health"])
        """
        self.app = app
        self.validator = validator
        self.excluded_paths = excluded_paths or ["/health"]
    
    async def __call__(self, scope, receive, send):
        if scope["type"] != "http":
            await self.app(scope, receive, send)
            return
        
        request = Request(scope, receive)
        path = request.url.path
        
        # Skip auth for excluded paths
        if any(path.startswith(excluded) for excluded in self.excluded_paths):
            await self.app(scope, receive, send)
            return
        
        # Validate token
        token = get_token_from_request(request)
        
        if not token:
            logger.warning(f"No bearer token in request to {path}")
            response = JSONResponse(
                {"error": "unauthorized", "message": "Missing bearer token"},
                status_code=401
            )
            await response(scope, receive, send)
            return
        
        logger.info(f"Validating token for request to {path}")
        is_valid, claims, error = await self.validator.validate_token(token)
        
        if not is_valid:
            logger.warning(f"Token validation failed for {path}: {error}")
            response = JSONResponse(
                {"error": "unauthorized", "message": f"Authentication failed - check MCP server Entra ID configuration. Details: {error}"},
                status_code=401
            )
            await response(scope, receive, send)
            return
        
        logger.info(f"Token validated successfully for {path}")
        # Store claims in scope for downstream use
        scope["state"] = scope.get("state", {})
        scope["state"]["token_claims"] = claims
        
        await self.app(scope, receive, send)


# =============================================================================
# Factory Function
# =============================================================================

def create_validator_from_env() -> Optional[EntraTokenValidator]:
    """
    Create an EntraTokenValidator from environment variables.
    
    Required environment variables:
        ENTRA_TENANT_ID or AZURE_TENANT_ID: Microsoft Entra tenant ID
        ENTRA_CLIENT_ID or AZURE_CLIENT_ID: Application (client) ID
        
    Optional:
        ENTRA_ALLOWED_AUDIENCES: Comma-separated list of allowed audiences
        DISABLE_ENTRA_AUTH: Set to "true" to disable authentication
    
    Returns:
        EntraTokenValidator if auth is configured and enabled, None otherwise
    """
    # Check if auth is explicitly disabled
    if not ENTRA_AUTH_ENABLED:
        if _auth_explicitly_disabled:
            logger.info("Entra auth explicitly disabled (DISABLE_ENTRA_AUTH=true)")
        else:
            logger.info("Entra auth not configured (missing credentials)")
        return None
    
    tenant_id = ENTRA_TENANT_ID
    client_id = ENTRA_CLIENT_ID
    
    allowed_audiences = None
    audiences_env = os.getenv("ENTRA_ALLOWED_AUDIENCES")
    if audiences_env:
        allowed_audiences = [a.strip() for a in audiences_env.split(",")]
    
    logger.info(f"Entra auth configured for tenant: {tenant_id}, client: {client_id}")
    
    return EntraTokenValidator(
        tenant_id=tenant_id,
        client_id=client_id,
        allowed_audiences=allowed_audiences
    )


def configure_auth_middleware():
    """Log authentication middleware configuration status."""
    if ENTRA_AUTH_ENABLED:
        logger.info("Entra ID authentication enabled")
        logger.info(f"  Tenant ID: {ENTRA_TENANT_ID}")
        logger.info(f"  Client ID (Audience): {ENTRA_CLIENT_ID}")
    else:
        if _auth_explicitly_disabled:
            logger.warning("Entra ID authentication DISABLED (DISABLE_ENTRA_AUTH=true)")
        else:
            logger.warning("Entra ID authentication DISABLED - MCP endpoints are public")

