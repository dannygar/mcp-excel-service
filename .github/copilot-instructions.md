# AI coding agent guide for this repo

## Big picture
- Purpose: Azure Container Apps-based MCP (Model Context Protocol) server providing Excel file manipulation capabilities for AI agents.
- Architecture: Remote MCP server using Azure Container Apps with FastMCP (Python 3.12+), deployed via Azure Developer CLI (azd).
- **Dual Protocol Support**: MCP (Streamable HTTP) for AI agents + REST API (OpenAPI) for direct HTTP clients
- Key modules:
  - `mcp-server/server.py`: MCP server with 4 Excel tools:
    - excel.updateRange: Update a range of cells in an Excel worksheet
    - excel.logTrades: Log multiple trades to an Excel workbook
    - excel.updateTradeWithDelta: Update delta value for a trade
    - excel.closeTrade: Close a trade by updating close date/time
- REST API endpoints (same functionality as MCP tools):
  - POST /api/v1/updateRange
  - POST /api/v1/logTrades
  - POST /api/v1/updateTradeWithDelta
  - POST /api/v1/closeTrade
  - GET /api/v1/openapi.json (OpenAPI specification)
- Config/secrets: Azure AD credentials stored as Container App secrets and injected via environment variables
- Required credentials: AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET (service principal)
- Infrastructure: `infra/mcp-server/main.bicep` defines Azure resources (Container App, Container Registry, Log Analytics); `azure.yaml` configures azd deployment.

## Runtime & workflows (Windows/PowerShell)
- Python: 3.11+ (3.12 recommended). Uses `uv` for dependency management.
- App Registration: Run `.\scripts\register-app.ps1` to create Entra ID app with Graph API permissions
- Local MCP server (Container App):
  - **Docker**: `docker build -t mcp-excel-server -f mcp-server/Dockerfile mcp-server/ && docker run -p 3000:3000 --env-file config/.env.local mcp-excel-server`
  - **Direct Python**: `cd mcp-server && uv run python server.py` (uses config/.env.local automatically)
  - MCP endpoint: `http://localhost:3000/mcp`
  - REST API: `http://localhost:3000/api/v1/*`
  - OpenAPI spec: `http://localhost:3000/api/v1/openapi.json`
  - Health endpoint: `http://localhost:3000/health`
  - Connect via MCP Inspector (`yarn inspector`) or VS Code Copilot agent mode
- MCP Inspector testing:
  - Install: `yarn install`
  - Launch inspector: `yarn inspector` (opens web UI at http://localhost:5173)
  - Test tools interactively with custom parameters and view request/response payloads
- REST API testing:
  - Use curl, Postman, or any HTTP client
  - Example: `curl -X POST http://localhost:3000/api/v1/logTrades -H "Content-Type: application/json" -d '{"url": "...", "file_name": "...", "sheet_name": "...", "trades": "..."}'`
- Deploy to Azure:
  - First time: `.\scripts\deploy-mcp-server.ps1` (creates App Registration + deploys)
  - With existing credentials: `azd up`
  - Redeploy code only: `azd deploy`
  - Teardown: `azd down`
- No formal test suite; validation via running tools and observing logs.
- Debugging:
  - Local: Run server directly, attach VS Code debugger
  - Azure: `az containerapp logs show --name <app-name> --resource-group <rg> --follow`

## Environment configuration
- All environment files are in `config/` folder (no duplicates in mcp-server/)
- Local development: Uses `config/.env.local` (MCP_ENV=local, default)
- Azure dev deployment: Uses `config/.env.dev` (MCP_ENV=dev)
- Azure prod deployment: Uses `config/.env.prod` (MCP_ENV=prod)
- Environment is auto-created by `register-app.ps1` script

## Secrets and configuration
- Required env vars:
  - `AZURE_TENANT_ID` - Azure AD tenant ID
  - `AZURE_CLIENT_ID` - App Registration client ID
  - `AZURE_CLIENT_SECRET` - App Registration client secret
- In Azure: Stored as Container App secrets, referenced in container env vars
- Locally: Set in `config/.env.local` file or environment variables

## Project conventions to follow
- Use `logging` with `basicConfig(level=logging.INFO)` for server logs.
- Keep secrets out of code; access via environment variables only.
- Use FastMCP decorators (`@mcp.tool()`) for tool definitions.
- Use FastMCP decorators (`@mcp.custom_route()`) for REST API endpoints.
- Return JSON strings from tools for consistent parsing.
- Use async/await for all Graph API calls.
- REST API endpoints should call the same underlying implementation as MCP tools.

## Typical task templates (examples)
- Add a new tool to `mcp-server/server.py`:
  1) Import required libraries
  2) Define async function with `@mcp.tool()` decorator
  3) Add docstring describing the tool
  4) Get headers via `headers = await get_graph_headers()`
  5) Make API calls with try-except error handling
  6) Return result as JSON string using `json.dumps()`
  7) Add corresponding REST API endpoint using `@mcp.custom_route()`
- Example: Adding a new Excel tool with REST API:
  ```python
  @mcp.tool(name="excel.getRange")
  async def excel_get_range(drive_id: str, item_id: str, sheet_name: str, address: str) -> str:
      """Get values from a range in an Excel worksheet."""
      try:
          headers = await get_graph_headers()
          workbook_url = build_workbook_url(drive_id, item_id)
          url = f"{workbook_url}/worksheets/{sheet_name}/range(address='{address}')"
          async with httpx.AsyncClient() as client:
              response = await client.get(url, headers=headers, timeout=30.0)
              if response.status_code == 200:
                  return json.dumps(response.json(), indent=2)
              return json.dumps({"status": "error", "message": response.text})
      except Exception as e:
          return json.dumps({"status": "error", "message": str(e)})
  ```

## API Provider Details
- **Microsoft Graph API**: Excel operations via `/drives/{drive-id}/items/{item-id}/workbook` endpoints
- Authentication: Client credentials flow (service principal)
- Token caching: Tokens are cached and refreshed 5 minutes before expiration

## Debugging tips
- Missing credentials: Check AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET env vars
- Token acquisition failures: Verify app registration has correct permissions and admin consent
- Container not starting: Check `az containerapp logs show` for errors
- HTTP 401/403: Verify app has Files.ReadWrite.All and Sites.ReadWrite.All permissions with admin consent
- Server not responding: Check health endpoint first (`/health`)

## Deployment
- Use `.\scripts\deploy-mcp-server.ps1` for full deployment with App Registration
- Use `.\scripts\register-app.ps1` to create/update App Registration only
- Or use `azd up` for standard Azure Developer CLI deployment (requires existing credentials)
- Container App runs on port 3000 with HTTP (Streamable HTTP) transport
- MCP endpoint: `https://<container-app-fqdn>/mcp`
- Health endpoint: `https://<container-app-fqdn>/health`
