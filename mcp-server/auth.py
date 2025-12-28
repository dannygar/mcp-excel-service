"""
Microsoft Entra ID Authentication Module

Provides token validation and middleware for authenticating incoming requests
from Azure AI Foundry agents using Project Managed Identity.

Authentication Flow:
    1. Extract bearer token from Authorization header
    2. Decode JWT to get claims (kid, issuer, audience, expiration)
    3. Fetch JWKS keys from Microsoft Entra ID
    4. Validate token signature using PyJWT
    5. Verify issuer matches configured tenant
    6. Verify audience matches configured client ID
"""

import os
import json
import logging
import time
from typing import Optional

import httpx
from starlette.requests import Request
from starlette.responses import JSONResponse
from starlette.middleware.base import BaseHTTPMiddleware

logger = logging.getLogger("mcp-excel-server")

# =============================================================================
# Configuration
# =============================================================================

# Get configuration for incoming token validation
# Check ENTRA_* first (set by Bicep for token validation), then fall back to AZURE_* (for Graph API)
ENTRA_TENANT_ID = os.getenv("ENTRA_TENANT_ID") or os.getenv("AZURE_TENANT_ID", "")
ENTRA_CLIENT_ID = os.getenv("ENTRA_CLIENT_ID") or os.getenv("AZURE_CLIENT_ID", "")  # This is the audience

# Auth is enabled if we have both tenant and client ID configured
# Note: ENABLE_ENTRA_AUTH env var is no longer needed - we just check if credentials are present
_has_entra_config = bool(ENTRA_TENANT_ID and ENTRA_CLIENT_ID)
ENTRA_AUTH_ENABLED = _has_entra_config

# OIDC/JWT validation settings
ENTRA_ISSUER_V1 = f"https://sts.windows.net/{ENTRA_TENANT_ID}/"
ENTRA_ISSUER_V2 = f"https://login.microsoftonline.com/{ENTRA_TENANT_ID}/v2.0"
ENTRA_JWKS_URI = f"https://login.microsoftonline.com/{ENTRA_TENANT_ID}/discovery/v2.0/keys"

# Cache for JWKS keys
_jwks_cache = {
    "keys": None,
    "fetched_at": 0,
    "cache_duration": 3600,  # Refresh keys every hour
}


# =============================================================================
# Token Validation Functions
# =============================================================================

async def get_jwks_keys() -> dict:
    """
    Fetch and cache the JSON Web Key Set (JWKS) from Microsoft Entra ID.
    Used for validating JWT token signatures.
    """
    global _jwks_cache
    
    current_time = time.time()
    if _jwks_cache["keys"] and current_time < _jwks_cache["fetched_at"] + _jwks_cache["cache_duration"]:
        return _jwks_cache["keys"]
    
    logger.info(f"Fetching JWKS from {ENTRA_JWKS_URI}")
    
    async with httpx.AsyncClient() as client:
        response = await client.get(ENTRA_JWKS_URI, timeout=30.0)
        if response.status_code == 200:
            _jwks_cache["keys"] = response.json()
            _jwks_cache["fetched_at"] = current_time
            logger.info("Successfully fetched and cached JWKS keys")
            return _jwks_cache["keys"]
        else:
            logger.error(f"Failed to fetch JWKS: {response.status_code}")
            raise ValueError(f"Failed to fetch JWKS: {response.status_code}")


def decode_jwt_without_verification(token: str) -> dict:
    """
    Decode a JWT token without verification to extract header and claims.
    Used to get the key ID (kid) for signature verification.
    """
    import base64
    
    parts = token.split('.')
    if len(parts) != 3:
        raise ValueError("Invalid JWT format")
    
    def decode_part(part: str) -> dict:
        # Add padding if needed
        padding = 4 - len(part) % 4
        if padding != 4:
            part += '=' * padding
        decoded = base64.urlsafe_b64decode(part)
        return json.loads(decoded)
    
    header = decode_part(parts[0])
    payload = decode_part(parts[1])
    
    return {"header": header, "payload": payload}


async def validate_entra_token(token: str) -> dict:
    """
    Validate a Microsoft Entra ID bearer token.
    
    Validates:
    - Token signature using JWKS
    - Token expiration
    - Token issuer (must be from configured tenant)
    - Token audience (must match AZURE_CLIENT_ID)
    
    Returns:
        Token claims if valid
        
    Raises:
        ValueError: If token is invalid
    """
    try:
        # Decode token to get claims (without signature verification first)
        decoded = decode_jwt_without_verification(token)
        header = decoded["header"]
        claims = decoded["payload"]
        
        # Validate basic structure
        if "alg" not in header or "kid" not in header:
            raise ValueError("Invalid token header")
        
        # Get current time for expiration check
        current_time = int(time.time())
        
        # Validate expiration
        exp = claims.get("exp", 0)
        if current_time >= exp:
            raise ValueError("Token has expired")
        
        # Validate not before (if present)
        nbf = claims.get("nbf", 0)
        if current_time < nbf:
            raise ValueError("Token is not yet valid")
        
        # Validate issuer - accept both v1 and v2 endpoints
        issuer = claims.get("iss", "")
        if issuer not in [ENTRA_ISSUER_V1, ENTRA_ISSUER_V2]:
            logger.warning(f"Invalid issuer: {issuer}. Expected: {ENTRA_ISSUER_V1} or {ENTRA_ISSUER_V2}")
            raise ValueError(f"Invalid token issuer: {issuer}")
        
        # Validate audience - must match our client ID
        aud = claims.get("aud", "")
        # Audience can be the client ID or the Application ID URI (api://<client-id>)
        valid_audiences = [ENTRA_CLIENT_ID, f"api://{ENTRA_CLIENT_ID}"]
        if aud not in valid_audiences:
            logger.warning(f"Invalid audience: {aud}. Expected one of: {valid_audiences}")
            raise ValueError(f"Invalid token audience: {aud}")
        
        # For production, verify signature with JWKS
        # This requires the cryptography and PyJWT libraries
        try:
            import jwt
            from jwt import PyJWKClient
            
            jwks_client = PyJWKClient(ENTRA_JWKS_URI)
            signing_key = jwks_client.get_signing_key_from_jwt(token)
            
            # Verify the token with proper signature validation
            verified_claims = jwt.decode(
                token,
                signing_key.key,
                algorithms=["RS256"],
                audience=valid_audiences,
                issuer=[ENTRA_ISSUER_V1, ENTRA_ISSUER_V2],
                options={"verify_exp": True, "verify_nbf": True}
            )
            
            logger.info(f"Token validated successfully. Subject: {verified_claims.get('sub', 'unknown')}")
            return verified_claims
            
        except ImportError:
            # PyJWT not installed - use basic validation only
            logger.warning("PyJWT not installed - using basic token validation without signature verification")
            logger.info(f"Token claims validated. Subject: {claims.get('sub', 'unknown')}")
            return claims
            
    except ValueError:
        raise
    except Exception as e:
        logger.error(f"Token validation error: {e}")
        raise ValueError(f"Token validation failed: {str(e)}")


def extract_bearer_token(request: Request) -> Optional[str]:
    """
    Extract the bearer token from the Authorization header.
    """
    auth_header = request.headers.get("Authorization", "")
    if auth_header.startswith("Bearer "):
        return auth_header[7:]
    return None


# =============================================================================
# Authentication Middleware
# =============================================================================

class EntraAuthMiddleware(BaseHTTPMiddleware):
    """
    Middleware to validate Microsoft Entra ID bearer tokens on incoming requests.
    
    - Health endpoint (/health) is excluded for Container Apps probes
    - All other endpoints require valid Entra ID token when auth is enabled
    - Returns 401 Unauthorized for invalid or missing tokens
    """
    
    async def dispatch(self, request: Request, call_next):
        # Allow health checks without authentication
        if request.url.path == "/health":
            return await call_next(request)
        
        # Skip auth if disabled
        if not ENTRA_AUTH_ENABLED:
            logger.debug("Entra auth disabled - allowing request without token validation")
            return await call_next(request)
        
        # Check for required configuration
        if not ENTRA_TENANT_ID or not ENTRA_CLIENT_ID:
            logger.error("Entra auth enabled but AZURE_TENANT_ID or AZURE_CLIENT_ID not configured")
            return JSONResponse(
                status_code=500,
                content={"error": "Server authentication not configured properly"}
            )
        
        # Extract bearer token
        token = extract_bearer_token(request)
        if not token:
            logger.warning(f"Missing bearer token for {request.url.path}")
            return JSONResponse(
                status_code=401,
                content={
                    "error": "Unauthorized",
                    "message": "Bearer token required. Use Microsoft Entra authentication with Project Managed Identity."
                }
            )
        
        # Validate token
        try:
            claims = await validate_entra_token(token)
            # Store claims in request state for potential use by handlers
            request.state.token_claims = claims
            request.state.authenticated = True
        except ValueError as e:
            logger.warning(f"Token validation failed: {e}")
            return JSONResponse(
                status_code=401,
                content={
                    "error": "Unauthorized",
                    "message": str(e)
                }
            )
        except Exception as e:
            logger.error(f"Unexpected error during token validation: {e}")
            return JSONResponse(
                status_code=401,
                content={
                    "error": "Unauthorized",
                    "message": "Token validation failed"
                }
            )
        
        return await call_next(request)


def configure_auth_middleware():
    """Configure authentication middleware for the MCP server."""
    if ENTRA_AUTH_ENABLED:
        logger.info("Entra ID authentication enabled")
        logger.info(f"  Tenant ID: {ENTRA_TENANT_ID}")
        logger.info(f"  Client ID (Audience): {ENTRA_CLIENT_ID}")
        logger.info(f"  Valid issuers: {ENTRA_ISSUER_V1}, {ENTRA_ISSUER_V2}")
    else:
        logger.warning("Entra ID authentication DISABLED - MCP endpoints are public")
