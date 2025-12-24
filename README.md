# MCP Excel Service

Azure Container Apps-based MCP (Model Context Protocol) server providing Excel file manipulation capabilities for AI agents via Microsoft Graph API.

## Architecture

- **Remote MCP server** using Azure Container Apps (FastMCP + Python 3.12)
- **Streamable HTTP Transport** for Azure AI Foundry integration
- **2 Excel manipulation tools** for updating cells and finding/updating rows
- **Data provider**: Microsoft Graph API (SharePoint/OneDrive)
- **Authentication**: Azure AD service principal (client credentials flow)
- **Auto-scaling**: 1-5 replicas based on HTTP load
- **Deployed via Azure Developer CLI** (`azd`)

## Prerequisites

- **Python 3.11+** (3.12 recommended)
- **Docker** (for local development and container builds)
- **Azure Developer CLI** (`azd`) for deployment
- **Azure CLI** (`az`) for resource management
- **Microsoft 365** with SharePoint/OneDrive access

**Setup:**
```pwsh
# Install uv (if not already installed)
irm https://astral.sh/uv/install.ps1 | iex

# Sync dependencies
cd mcp-server
uv sync
```

## Quick Start

### Local Development

**Option 1: Run the Container App locally with Docker**

```pwsh
# Build the container
docker build -t mcp-excel-server -f mcp-server/Dockerfile mcp-server/

# Run with Azure AD credentials
docker run -p 3000:3000 --env-file mcp-server/.env mcp-excel-server

# Server available at http://localhost:3000/mcp
```

**Option 2: Run directly with Python**

```pwsh
cd mcp-server

# Ensure .env file exists with Azure AD credentials (created by register-app.ps1)
# Or set environment variables manually:
$env:AZURE_TENANT_ID = "your-tenant-id"
$env:AZURE_CLIENT_ID = "your-client-id"
$env:AZURE_CLIENT_SECRET = "your-client-secret"

# Run the server
uv run python server.py

# Server available at http://localhost:3000/mcp
```

### Testing with MCP Inspector

```pwsh
# Install MCP Inspector
yarn install

# Start the server first (in another terminal), then launch inspector
yarn inspector
```

The inspector opens at `http://localhost:5173` where you can:
- Browse available tools (`excel.updateRowByLookup`, `excel.updateRange`)
- Test tool invocations with custom parameters
- View request/response payloads in real-time

---

## Deploy to Azure

### 1. Register App & Configure Credentials

```pwsh
# Create Entra ID App Registration with Graph API permissions
.\scripts\register-app.ps1

# This creates:
# - App Registration with Files.ReadWrite.All and Sites.ReadWrite.All permissions
# - Client secret for authentication
# - .env file with credentials for local development
# - Foundry-compatible authentication (Application ID URI)
```

> ⚠️ **Important**: After running the script, grant admin consent for the API permissions in the Azure Portal.

### 2. Deploy with Deployment Script (Recommended)

```pwsh
# Full deployment with Foundry integration
.\scripts\deploy-mcp-server.ps1

# This script will:
# - Create/update App Registration (if needed)
# - Auto-discover Azure AI Foundry projects
# - Let you select which projects should access the MCP server
# - Deploy infrastructure via Bicep
# - Build and push Docker image to ACR
# - Configure Container App with secrets
# - Display Foundry integration instructions
```

### 3. Deploy with Azure Developer CLI

```pwsh
# Alternative: Deploy infrastructure + container directly
azd up
```

### 4. Verify Deployment

```pwsh
# Check health endpoint
Invoke-WebRequest -Uri "https://<your-container-app>.azurecontainerapps.io/health"

# The MCP endpoint is available at:
# https://<your-container-app>.azurecontainerapps.io/mcp
```

### 5. Teardown

```pwsh
# Remove all Azure resources
azd down
```

---

## Connect to Azure AI Foundry

### Add MCP Server to Foundry Agent

> **📖 For detailed instructions, see [FOUNDRY_INTEGRATION.md](docs/FOUNDRY_INTEGRATION.md)**

**Quick Setup:**

1. Navigate to [Azure AI Foundry](https://ai.azure.com)
2. Go to your project → **Build** → **Create agent**
3. Click **+ Add** in the Tools section
4. Select **Custom** → **Model Context Protocol** → **Create**
5. Configure the connection:

| Field | Value |
|-------|-------|
| **Name** | `MCP Excel Service` |
| **Remote MCP Server** | `https://<your-container-app>.azurecontainerapps.io/mcp` |
| **Authentication** | Microsoft Entra |
| **Type** | Project Managed Identity |
| **Audience** | `<your-client-id>` (from deployment output) |

6. Click **Connect**

### Test in Chat Playground

Try these prompts:
- "Append sales data to my Excel file in SharePoint"
- "Update cells A1:B5 in my inventory spreadsheet"
- "Add a new row to the Products table in my workbook"

---

## Connect to VS Code GitHub Copilot

Add to your `.vscode/mcp.json`:

```json
{
  "servers": {
    "mcp-excel-remote": {
      "type": "http",
      "url": "https://<your-container-app>.azurecontainerapps.io/mcp"
    },
    "mcp-excel-local": {
      "type": "http",
      "url": "http://localhost:3000/mcp"
    }
  }
}
```

---

## MCP Tools

Both tools accept a single JSON object as input, making them fully compatible with Foundry Agent schema validation.

### `excel.updateRowByLookup`

Find a row by looking up a reference value in a column and update specific cells in that row (or a relative row).

**Parameters:**
| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `url` | string | ✓ | SharePoint/OneDrive URL to document library |
| `file_name` | string | ✓ | Excel file name with .xlsx extension |
| `sheet_name` | string | ✓ | Worksheet name |
| `search_column` | string | ✓ | Column letter to search (e.g., "A", "C") |
| `reference_value` | string | ✓ | Value to find (supports dates like "12/22/2025") |
| `target_columns` | string | ✓ | **JSON array** of column letters (e.g., `'["D", "F", "H"]'`) |
| `values` | string | ✓ | **JSON array** of values to write (e.g., `'["value1", 123, true]'`) |
| `row_offset` | integer | | Rows below found row to update (default: 0) |

> **Note**: `target_columns` and `values` must be valid JSON strings. This ensures compatibility with all MCP clients including Azure AI Foundry.

**Example Request:**
```json
{
  "url": "https://contoso.sharepoint.com/Shared%20Documents/Forms/AllItems.aspx",
  "file_name": "2025 Trade Tracker.xlsx",
  "sheet_name": "December",
  "search_column": "C",
  "reference_value": "12/23/2025",
  "target_columns": "[\"C\", \"E\", \"I\", \"J\", \"L\"]",
  "values": "[\"12/23/2025\", \"11:36 AM\", \"VPCS\", 0.25, 25]",
  "row_offset": 1
}
```

**Response:**
```json
{
  "status": "success",
  "message": "Successfully updated 5 cells in row 15",
  "sheet_name": "December",
  "found_row": 14,
  "target_row": 15,
  "row_offset": 1,
  "updated_cells": ["C15", "E15", "I15", "J15", "L15"]
}
```

---

### `excel.updateRange`

Update a range of cells in an Excel worksheet with a 2D array of values.

**Parameters:**
| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `url` | string | ✓ | SharePoint/OneDrive URL to document library |
| `file_name` | string | ✓ | Excel file name with .xlsx extension |
| `sheet_name` | string | ✓ | Worksheet name |
| `address` | string | ✓ | Cell range address (e.g., "A1:C3") |
| `values` | string | ✓ | **JSON 2D array** where each inner array is a row |

> **Note**: `values` must be a valid JSON string containing a 2D array.

**Example Request:**
```json
{
  "url": "https://contoso.sharepoint.com/sites/Sales/Shared%20Documents",
  "file_name": "Sales.xlsx",
  "sheet_name": "Sheet1",
  "address": "A1:C2",
  "values": "[[\"Name\", \"Quantity\", \"Price\"], [\"Widget\", 100, 9.99]]"
}
```

**Response:**
```json
{
  "status": "success",
  "message": "Successfully updated range 'A1:C2' in sheet 'Sheet1'",
  "file_name": "Sales.xlsx",
  "sheet_name": "Sheet1",
  "address": "A1:C2",
  "row_count": 2,
  "column_count": 3
}
```

---

### `excel.logTrades`

**High-level tool** for logging multiple trades to a pre-configured trade tracker spreadsheet. The spreadsheet URL and file name are configured via environment variables, simplifying agent prompts.

**Environment Variables (required):**
| Variable | Description | Example |
|----------|-------------|---------|
| `TRADE_TRACKER_URL` | SharePoint/OneDrive URL | `https://contoso.sharepoint.com/Shared%20Documents/Forms/AllItems.aspx` |
| `TRADE_TRACKER_FILE` | Excel workbook filename | `2025 Trade Tracker.xlsx` |

**Parameters:**
| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `trades` | string | ✓ | **JSON array** of trade objects (see below) |
| `reference_date` | string | | Date to search for in column C (default: last date in column C) |
| `sheet_name` | string | | Worksheet name (default: current month, e.g., "December") |

**Trade Object Fields:**
| Field | Column | Type | Description |
|-------|--------|------|-------------|
| `open_date` | C | string | Date when trade was opened (e.g., "12/23/2025") |
| `open_time` | E | string | Time when trade was opened (e.g., "10:30 AM") |
| `close_date` | F | string | Date when trade was closed, if known |
| `close_time` | G | string | Time when trade was closed (e.g., "4:00 PM") |
| `strategy` | I | string | Strategy name (e.g., "VPCS", "VCCS", "IC") |
| `credit` | J | number | Credit received when opening |
| `debit` | K | number | Debit paid if closed before expiration |
| `contracts` | L | integer | Number of contracts |
| `open_fees` | N | number | Total fees paid when trade was opened |
| `close_fees` | O | number | Total fees paid if closed before expiration |
| `sold_call_strike` | Q | number | Strike price for sold calls |
| `sold_put_strike` | R | number | Strike price for sold puts |
| `width` | T | number | Width in USD between sold and bought strikes |
| `expired` | - | boolean | If `true`, auto-fills close_date=open_date, close_time="4:00 PM", debit=0 |

> **Note**: For backward compatibility, `date` and `time` are aliases for `open_date` and `open_time`, and `fees` is an alias for `open_fees`.

**Expired Trade Handling:**
For 0DTE SPX options that expire worthless, set `"expired": true` in the trade object. The tool will automatically populate:
- `close_date` = `open_date` (same-day expiration)
- `close_time` = "4:00 PM" (market close)
- `debit` = 0 (expired worthless)

**Example Request (with expired trades):**
```json
{
  "trades": "[{\"open_date\": \"12/23/2025\", \"open_time\": \"10:42 AM\", \"strategy\": \"VCCS\", \"credit\": 0.10, \"contracts\": 25, \"open_fees\": 88.27, \"sold_call_strike\": 6920, \"width\": 15, \"expired\": true}, {\"open_date\": \"12/23/2025\", \"open_time\": \"11:36 AM\", \"strategy\": \"VPCS\", \"credit\": 0.25, \"contracts\": 25, \"open_fees\": 88.27, \"sold_put_strike\": 6860, \"width\": 15, \"expired\": true}]",
  "reference_date": "12/22/2025",
  "sheet_name": "December"
}
```

**Response:**
```json
{
  "status": "success",
  "message": "Successfully logged 2 trades",
  "trades_logged": 2,
  "file_name": "2025 Trade Tracker.xlsx",
  "sheet_name": "December",
  "reference_date": "12/22/2025",
  "results": [
    {
      "trade_index": 1,
      "row": 15,
      "open_date": "12/23/2025",
      "open_time": "10:42 AM",
      "close_date": "12/23/2025",
      "close_time": "4:00 PM",
      "strategy": "VCCS",
      "credit": 0.10,
      "debit": 0,
      "contracts": 25,
      "expired": true
    },
    {
      "trade_index": 2,
      "row": 16,
      "open_date": "12/23/2025",
      "open_time": "11:36 AM",
      "close_date": "12/23/2025",
      "close_time": "4:00 PM",
      "strategy": "VPCS",
      "credit": 0.25,
      "debit": 0,
      "contracts": 25,
      "expired": true
    }
  ]
}
```

**Simplified Foundry Agent Prompt:**
```
Show me all SPX trades from today that expired, then log them to my trade tracker.
```

---

## Supported SharePoint URL Formats

The MCP server automatically resolves various SharePoint/OneDrive URL formats:

| URL Type | Example |
|----------|---------|
| SharePoint Site | `https://contoso.sharepoint.com/sites/Sales/Shared%20Documents` |
| Document Library View | `https://contoso.sharepoint.com/Shared%20Documents/Forms/AllItems.aspx` |
| OneDrive for Business | `https://contoso-my.sharepoint.com/personal/user/Documents` |

---

## Project Structure

```
├── mcp-server/
│   ├── server.py              # MCP server (FastMCP + Streamable HTTP)
│   ├── config.py              # Configuration (strategy mapping, trade tracker URL)
│   ├── Dockerfile             # Container image definition
│   ├── requirements.txt       # Python dependencies
│   └── pyproject.toml         # Project metadata
├── infra/
│   └── mcp-server/
│       ├── main.bicep         # Azure infrastructure (ACR, Log Analytics, etc.)
│       └── container-app.bicep # Container App definition (scaling 1-5 replicas)
├── scripts/
│   ├── deploy-mcp-server.ps1  # Full deployment with Foundry integration
│   └── register-app.ps1       # Entra ID App Registration
├── docs/
│   ├── FOUNDRY_INTEGRATION.md # Step-by-step Foundry setup guide
│   └── AZURE_DEPLOYMENT.md    # Detailed Azure setup guide
├── azure.yaml                 # azd configuration
├── package.json               # MCP Inspector dependencies
└── .vscode/
    └── mcp.json               # VS Code MCP configuration
```

---

## Configuration

### Environment Variables

| Variable | Required | Description |
|----------|----------|-------------|
| `AZURE_TENANT_ID` | Yes | Azure AD tenant ID |
| `AZURE_CLIENT_ID` | Yes | App Registration client ID |
| `AZURE_CLIENT_SECRET` | Yes | App Registration client secret |
| `TRADE_TRACKER_URL` | For logTrades | SharePoint/OneDrive URL for trade tracker |
| `TRADE_TRACKER_FILE` | For logTrades | Excel filename for trade tracker |
| `PORT` | No | Server port (default: 3000) |
| `HOST` | No | Server host (default: 0.0.0.0) |

### Required Graph API Permissions

| Permission | Type | Description |
|------------|------|-------------|
| `Files.ReadWrite.All` | Application | Read/write files in SharePoint/OneDrive |
| `Sites.ReadWrite.All` | Application | Read/write items in all site collections |

> ⚠️ These permissions require **admin consent** after App Registration is created.

### Scaling Configuration

The Container App is configured for auto-scaling:
- **Minimum replicas**: 1 (always running)
- **Maximum replicas**: 5
- **Scale trigger**: 50 concurrent HTTP requests

---

## Testing

### Run Test Suite

The test suite validates schema compatibility and tool functionality:

```pwsh
# Start the MCP server first
cd mcp-server
uv run python server.py

# In another terminal, run all tests
cd mcp-server
uv run python test_server.py

# Run specific tests
uv run python test_server.py --test health
uv run python test_server.py --test list_tools
uv run python test_server.py --test update_row
uv run python test_server.py --test update_range
```

### Integration Test (with real SharePoint)

```pwsh
uv run python test_server.py --test integration `
  --sharepoint-url "https://contoso.sharepoint.com/Shared%20Documents" `
  --file-name "MyWorkbook.xlsx" `
  --sheet-name "Sheet1"
```

### Test Against Deployed Server

```pwsh
uv run python test_server.py --url "https://<your-container-app>.azurecontainerapps.io"
```

---

## Debugging

### View Container Logs

```pwsh
# Stream logs from Azure Container Apps
az containerapp logs show `
  --name <container-app-name> `
  --resource-group <resource-group> `
  --follow

# Or use azd
azd monitor --logs
```

### Local Debugging

1. Start the server: `cd mcp-server && uv run python server.py`
2. Set breakpoints in `server.py`
3. Attach debugger (VS Code: Python: Attach to Local Process)

### Common Issues

| Issue | Solution |
|-------|----------|
| 401 Unauthorized | Check Graph API permissions have admin consent |
| Token acquisition failed | Verify AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET |
| File not found | Verify SharePoint URL format and file name |
| Health check fails | Ensure Container App is running and PORT is 3000 |
| `expired` field not working | Check logs - the tool handles both boolean `true` and string `"true"` |
| Close date/time not populated | Ensure `expired: true` is set in the trade object |
| Module 'config' not found | Ensure `config.py` is copied in Dockerfile |

---

## Documentation

- [Foundry Integration Guide](docs/FOUNDRY_INTEGRATION.md)
- [Azure Deployment Guide](docs/AZURE_DEPLOYMENT.md)
- [Model Context Protocol Documentation](https://modelcontextprotocol.io/)
- [FastMCP Documentation](https://gofastmcp.com/)
- [Microsoft Graph API - Excel](https://learn.microsoft.com/graph/api/resources/excel)

---

## License

MIT
