# MCP Finance Data Server

Azure Container Apps-based MCP (Model Context Protocol) server providing company earnings data for AI agents.

## Architecture

- **Remote MCP server** using Azure Container Apps (FastMCP + Python 3.12)
- **HTTP Transport** (Streamable HTTP) for Azure AI Foundry integration
- **2 company earnings tools** for revenue and free cash flow analysis
- **Data provider**: Alpha Vantage API (SEC filings)
- **Supported companies**: Microsoft (MSFT), Tesla (TSLA), NVIDIA (NVDA)
- **Deployed via Azure Developer CLI** (`azd`)

## Prerequisites

- **Python 3.11+** (3.12 recommended)
- **Docker** (for local development and container builds)
- **Azure Developer CLI** (azd) for deployment
- **Azure CLI** for resource management

**Setup:**
```pwsh
# Install uv (if not already installed)
irm https://astral.sh/uv/install.ps1 | iex

# Sync dependencies
uv sync
```

## Quick Start

### Local Development

**Option 1: Run the Container App locally with Docker**

```pwsh
# Build the container
docker build -t mcp-excel-server -f mcp-server/Dockerfile mcp-server/

# Run with your API key
docker run -p 3000:3000 -e ALPHAVANTAGE_API_KEY="your_key" mcp-finance-server

# Server available at http://localhost:3000/mcp
```

**Option 2: Run directly with Python**

```pwsh
cd mcp-server

# Set your access token
$env:GRAPH_ACCESS_TOKEN = "your_token"

# Run the server
uv run python server.py

# Server available at http://localhost:3000/mcp
```

### Testing with MCP Inspector

```pwsh
# Install MCP Inspector
yarn install

# Start the server first, then launch inspector
yarn inspector
```

The inspector opens at `http://localhost:5173` where you can:
- Browse available tools (`get_company_revenue`, `get_company_free_cash_flow`)
- Test tool invocations with custom parameters
- View request/response payloads in real-time

---

## Deploy to Azure

### 1. Configure API Keys

```pwsh
# Required: Set your Alpha Vantage API key
$env:ALPHAVANTAGE_API_KEY = "your_alpha_vantage_key"
```

**Get a free API key:** https://www.alphavantage.co/support/#api-key

### 2. Deploy with Azure Developer CLI

```pwsh
# Deploy infrastructure + container
azd up

# Or use the deployment script for more control
.\scripts\deploy-mcp-server.ps1
```

### 3. Verify Deployment

```pwsh
# Check health endpoint
Invoke-WebRequest -Uri "https://<your-container-app>.azurecontainerapps.io/health"

# The MCP endpoint is available at:
# https://<your-container-app>.azurecontainerapps.io/mcp
```

### 4. Teardown

```pwsh
# Remove all Azure resources
azd down
```

---

## Connect to Azure AI Foundry

### Add MCP Server to Foundry Agent

1. Navigate to your AI Foundry project: https://ai.azure.com
2. Go to **Agent** → **Tools** → **MCP Servers**
3. Click **Add MCP Server**
4. Configure:
   - **Name**: `mcp-finance`
   - **Transport Type**: Streamable HTTP
   - **Endpoint URL**: `https://<your-container-app>.azurecontainerapps.io/mcp`
   - **Authentication**: None
5. Click **Save**

### Test in Chat Playground

Try these prompts:
- "What was Microsoft's revenue for Q4 FY2024?"
- "Get Tesla's free cash flow for Q2 FY2024"
- "Compare NVIDIA's revenue across Q1-Q4 FY2024"

---

## Connect to VS Code GitHub Copilot

Add to your `.vscode/mcp.json`:

```json
{
  "servers": {
    "mcp-finance-remote": {
      "type": "http",
      "url": "https://<your-container-app>.azurecontainerapps.io/mcp"
    },
    "mcp-finance-local": {
      "type": "http",
      "url": "http://localhost:3000/mcp"
    }
  }
}
```

---

## MCP Tools

### `get_company_revenue`

Get total revenue in USD millions for a company's quarterly earnings report.

**Parameters:**
| Parameter | Type | Description |
|-----------|------|-------------|
| `company_symbol` | string | Company ticker (MSFT, TSLA, or NVDA) |
| `fiscal_year` | number | Fiscal year (e.g., 2024) |
| `fiscal_quarter` | number | Fiscal quarter (1, 2, 3, or 4) |

**Example Response:**
```json
{
  "company": "MSFT",
  "fiscal_year": 2024,
  "fiscal_quarter": 4,
  "fiscal_date_ending": "2024-06-30",
  "total_revenue_usd_millions": 64725.00,
  "currency": "USD",
  "data_source": "Alpha Vantage"
}
```

---

### `get_company_free_cash_flow`

Get free cash flow in USD millions for a company's quarterly earnings report.

**Parameters:**
| Parameter | Type | Description |
|-----------|------|-------------|
| `company_symbol` | string | Company ticker (MSFT, TSLA, or NVDA) |
| `fiscal_year` | number | Fiscal year (e.g., 2024) |
| `fiscal_quarter` | number | Fiscal quarter (1, 2, 3, or 4) |

**Example Response:**
```json
{
  "company": "MSFT",
  "fiscal_year": 2024,
  "fiscal_quarter": 4,
  "fiscal_date_ending": "2024-06-30",
  "free_cash_flow_usd_millions": 23255.00,
  "operating_cash_flow_usd_millions": 28515.00,
  "capital_expenditures_usd_millions": 5260.00,
  "currency": "USD",
  "data_source": "Alpha Vantage"
}
```

---

## Supported Companies

| Symbol | Company Name |
|--------|--------------|
| MSFT | Microsoft Corporation |
| TSLA | Tesla, Inc. |
| NVDA | NVIDIA Corporation |

---

## Project Structure

```
├── mcp-server/
│   ├── server.py              # MCP server (FastMCP + Streamable HTTP)
│   ├── Dockerfile             # Container image definition
│   ├── requirements.txt       # Python dependencies
│   └── pyproject.toml         # Project metadata
├── infra/
│   └── mcp-server/
│       ├── main.bicep         # Azure infrastructure
│       └── container-app.bicep # Container App definition
├── config/
│   ├── .env.prod              # Production API keys
│   └── .env.dev               # Development API keys
├── scripts/
│   └── deploy-mcp-server.ps1  # Deployment script
├── docs/
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
| `ALPHAVANTAGE_API_KEY` | Yes | Alpha Vantage API key for financial data |
| `PORT` | No | Server port (default: 3000) |
| `HOST` | No | Server host (default: 0.0.0.0) |

### API Rate Limits

Alpha Vantage free tier:
- 25 requests per day
- 5 requests per minute

Consider upgrading for production use.

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

---

## Documentation

- [Azure Deployment Guide](docs/AZURE_DEPLOYMENT.md)
- [Model Context Protocol Documentation](https://modelcontextprotocol.io/)
- [FastMCP Documentation](https://gofastmcp.com/)
- [Alpha Vantage API](https://www.alphavantage.co/documentation/)

---

## License

MIT
