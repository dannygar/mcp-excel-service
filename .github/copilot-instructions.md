# AI coding agent guide for this repo

## Big picture
- Purpose: Azure Container Apps-based MCP (Model Context Protocol) server providing Excel file manipulation capabilities for AI agents.
- Architecture: Remote MCP server using Azure Container Apps with FastMCP (Python 3.12+), deployed via Azure Developer CLI (azd).
- Key modules:
  - `mcp-server/server.py`: MCP server with 2 Excel tools:
    - excel.appendRows: Append rows to an Excel table in SharePoint/OneDrive
    - excel.updateRange: Update a range of cells in an Excel worksheet
- Config/secrets: Azure AD credentials stored as Container App secrets and injected via environment variables
- Required credentials: AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET (service principal)
- Infrastructure: `infra/mcp-server/main.bicep` defines Azure resources (Container App, Container Registry, Log Analytics); `azure.yaml` configures azd deployment.

## Runtime & workflows (Windows/PowerShell)
- Python: 3.11+ (3.12 recommended). Uses `uv` for dependency management.
- App Registration: Run `.\scripts\register-app.ps1` to create Entra ID app with Graph API permissions
- Local MCP server (Container App):
  - **Docker**: `docker build -t mcp-excel-server -f mcp-server/Dockerfile mcp-server/ && docker run -p 3000:3000 --env-file mcp-server/.env mcp-excel-server`
  - **Direct Python**: `cd mcp-server && uv run python server.py` (uses .env file automatically)
  - MCP endpoint: `http://localhost:3000/mcp`
  - Health endpoint: `http://localhost:3000/health`
  - Connect via MCP Inspector (`yarn inspector`) or VS Code Copilot agent mode
- MCP Inspector testing:
  - Install: `yarn install`
  - Launch inspector: `yarn inspector` (opens web UI at http://localhost:5173)
  - Test tools interactively with custom parameters and view request/response payloads
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
- Production: Environment variables injected via Azure Container App secrets
- Local development: Uses `.env` file in mcp-server folder (auto-created by register-app.ps1)

## Secrets and configuration
- Required env vars:
  - `AZURE_TENANT_ID` - Azure AD tenant ID
  - `AZURE_CLIENT_ID` - App Registration client ID
  - `AZURE_CLIENT_SECRET` - App Registration client secret
- In Azure: Stored as Container App secrets, referenced in container env vars
- Locally: Set in `mcp-server/.env` file or environment variables

## Project conventions to follow
- Use `logging` with `basicConfig(level=logging.INFO)` for server logs.
- Keep secrets out of code; access via environment variables only.
- Use FastMCP decorators (`@mcp.tool()`) for tool definitions.
- Return JSON strings from tools for consistent parsing.
- Use async/await for all Graph API calls.

## Typical task templates (examples)
- Add a new tool to `mcp-server/server.py`:
  1) Import required libraries
  2) Define async function with `@mcp.tool()` decorator
  3) Add docstring describing the tool
  4) Get headers via `headers = await get_graph_headers()`
  5) Make API calls with try-except error handling
  6) Return result as JSON string using `json.dumps()`
- Example: Adding a new Excel tool:
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
