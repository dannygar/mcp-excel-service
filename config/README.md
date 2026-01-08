# Configuration Directory

This directory holds all environment configuration files for the MCP Excel Service.

## Files

| File | Purpose |
|------|---------|
| `.env.example` | Template with all available settings (committed to git) |
| `.env.local` | Local development settings with `DISABLE_ENTRA_AUTH=true` (git-ignored) |
| `.env.dev` | Development/staging deployment settings with auth enabled (git-ignored) |
| `.env.prod` | Production deployment settings (optional, git-ignored) |

**Note:** All `.env.*` files except `.env.example` are excluded from git via `.gitignore`.

## Setup

### Local Development

Run the app registration script to create credentials:

```pwsh
.\scripts\register-app.ps1
```

This will create `config/.env.local` with your Azure AD credentials and local development settings (including `DISABLE_ENTRA_AUTH=true` for MCP Inspector testing).

### Azure Deployment

For Azure Container Apps deployment, the `deploy-mcp-server.ps1` script will:
1. Read credentials from `config/.env.{EnvironmentName}` (defaults to `dev`)
2. Fall back to `config/.env.local` if environment-specific file doesn't exist
3. Inject credentials as Container App secrets

```pwsh
# Deploy to dev environment (default)
.\scripts\deploy-mcp-server.ps1

# Deploy to prod environment
.\scripts\deploy-mcp-server.ps1 -EnvironmentName prod
```

## Configuration Loading Priority

The server uses `MCP_ENV` environment variable to select the config file (defaults to `local`):

1. `config/.env.{MCP_ENV}` - Environment-specific config (e.g., `.env.local`, `.env.dev`, `.env.prod`)
2. `config/.env.local` - Local development fallback
3. Environment variables - Container Apps secrets / OS environment

### Setting MCP_ENV

- **Local development:** Not needed (defaults to `local`, uses `.env.local`)
- **Azure Container Apps:** Set via deployment script based on `-EnvironmentName` parameter
