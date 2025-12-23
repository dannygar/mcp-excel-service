# AI coding agent guide for this repo

## Big picture
- Purpose: Azure Container Apps-based MCP (Model Context Protocol) server providing company earnings data for AI agents.
- Architecture: Remote MCP server using Azure Container Apps with FastMCP (Python 3.12+), deployed via Azure Developer CLI (azd).
- Key modules:
  - `container-app/server.py`: MCP server with 2 company earnings tools:
    - get_company_revenue: Get quarterly revenue in USD millions (via Alpha Vantage INCOME_STATEMENT)
    - get_company_free_cash_flow: Get quarterly free cash flow in USD millions (via Alpha Vantage CASH_FLOW)
  - Supported companies: Microsoft (MSFT), Tesla (TSLA), NVIDIA (NVDA)
- Config/secrets: API key stored as Container App secret and injected via environment variable
- Required API key: ALPHAVANTAGE_API_KEY only
- Infrastructure: `infra/container-app/main.bicep` defines Azure resources (Container App, Container Registry, Log Analytics); `azure.yaml` configures azd deployment.

## Runtime & workflows (Windows/PowerShell)
- Python: 3.11+ (3.12 recommended). Uses `uv` for dependency management.
- Local MCP server (Container App):
  - **Docker**: `docker build -t mcp-finance-server -f container-app/Dockerfile container-app/ && docker run -p 3000:3000 -e ALPHAVANTAGE_API_KEY="your_key" mcp-finance-server`
  - **Direct Python**: `cd container-app && $env:ALPHAVANTAGE_API_KEY="your_key" && uv run python server.py`
  - MCP endpoint: `http://localhost:3000/mcp`
  - Health endpoint: `http://localhost:3000/health`
  - Connect via MCP Inspector (`yarn inspector`) or VS Code Copilot agent mode
- MCP Inspector testing:
  - Install: `yarn install`
  - Launch inspector: `yarn inspector` (opens web UI at http://localhost:5173)
  - Test tools interactively with custom parameters and view request/response payloads
- Deploy to Azure:
  - Set environment variable: `$env:ALPHAVANTAGE_API_KEY="your_key"`
  - Deploy: `azd up` or `.\scripts\deploy-container-app.ps1`
  - Redeploy code only: `azd deploy`
  - Teardown: `azd down`
- No formal test suite; validation via running tools and observing logs.
- Debugging:
  - Local: Run server directly, attach VS Code debugger
  - Azure: `az containerapp logs show --name <app-name> --resource-group <rg> --follow`

## Environment configuration
- Production: Environment variables injected via Azure Container App secrets
- Local development: Set `ALPHAVANTAGE_API_KEY` environment variable directly

## Secrets and configuration
- Required env var: `ALPHAVANTAGE_API_KEY` - for company earnings data (INCOME_STATEMENT and CASH_FLOW endpoints)
- In Azure: Stored as Container App secret, referenced in container env vars
- Locally: Set as environment variable before running the server

## Project conventions to follow
- Use `logging` with `basicConfig(level=logging.INFO)` for server logs.
- Keep secrets out of code; access via environment variables only.
- Use FastMCP decorators (`@mcp.tool()`) for tool definitions.
- Return JSON strings from tools for consistent parsing.

## Typical task templates (examples)
- Add a new tool to `container-app/server.py`:
  1) Import required libraries
  2) Define async function with `@mcp.tool()` decorator
  3) Add docstring describing the tool
  4) Make API calls with try-except error handling
  5) Return result as JSON string using `json.dumps()`
- Example: Adding a new market data tool:
  ```python
  @mcp.tool()
  async def get_stock_price(symbol: str) -> str:
      """Get current stock price for a given symbol."""
      try:
          api_key = get_api_key()
          # Make API call...
          return json.dumps({"symbol": symbol, "price": price})
      except Exception as e:
          return json.dumps({"status": "error", "message": str(e)})
  ```

## API Provider Details
- **Alpha Vantage**: Market data for stocks, ETFs, indices. Free tier: 25 requests/day, 5 requests/minute.

## Debugging tips
- Missing API key errors: Check `ALPHAVANTAGE_API_KEY` env var is set
- Container not starting: Check `az containerapp logs show` for errors
- HTTP 401/403: Verify API key is valid and not expired
- Rate limit errors: Respect Alpha Vantage limits (25/day free tier)
- Server not responding: Check health endpoint first (`/health`)

## Deployment
- Use `.\scripts\deploy-container-app.ps1` for full deployment with verification
- Or use `azd up` for standard Azure Developer CLI deployment
- Container App runs on port 3000 with HTTP (Streamable HTTP) transport
- MCP endpoint: `https://<container-app-fqdn>/mcp`
- Health endpoint: `https://<container-app-fqdn>/health`
