# Azure Deployment Checklist - Market Intelligence MCP Server

Quick reference guide for deploying the MCP Server to Azure Functions.

## Prerequisites

- [ ] Python 3.11 or 3.12 installed
- [ ] Azure CLI installed and authenticated (`az login`)
- [ ] Azure Developer CLI (azd) installed
- [ ] Azure Functions Core Tools >= 4.0.7030
- [ ] Docker Desktop running (for local testing with Azurite)

## API Keys Required

Before deployment, obtain these API keys:

### Required (Minimum Functionality)
- [ ] **ALPHAVANTAGE_API_KEY** - [Get key](https://www.alphavantage.co/support/#api-key)
  - Free tier: 25 requests/day
  - Used for: Market data snapshots, history, and company earnings data

### Optional (Enhanced Features)
- [ ] **FRED_API_KEY** - [Get key](https://fred.stlouisfed.org/docs/api/api_key.html)
  - Free, generous limits
  - Used for: Macroeconomic data and yield spreads
  
- [ ] **NEWSAPI_KEY** - [Get key](https://newsapi.org/register)
  - Free tier: 100 requests/day
  - Used for: News headlines tool

### Optional (Additional Features)
- [ ] **FINNHUB_API_KEY** - [Get key](https://finnhub.io/register)
  - Free tier: 60 calls/minute
  - Used for: Alternative market data provider
  
- [ ] **RAPIDAPI_KEY** - [Get key](https://rapidapi.com/)
  - Subscription required
  - Used for: Social sentiment analysis (Reddit, StockTwits, Twitter, news)

## Deployment Steps

### 1. Set Environment Variables

```pwsh
# Required API key (minimum)
$env:ALPHAVANTAGE_API_KEY="your_alpha_vantage_key"

# Optional API keys (for enhanced features)
$env:FRED_API_KEY="your_fred_api_key"
$env:NEWSAPI_KEY="your_newsapi_key"
$env:FINNHUB_API_KEY="your_finnhub_key"
$env:RAPIDAPI_KEY="your_rapidapi_key"
```

### 2. Configure azd Environment

```pwsh
# Initialize azd (if not already done)
azd init

# Ensure VNET is disabled (for simpler storage access)
# Edit .azure/<your-env-name>/.env and set:
# VNET_ENABLED="false"
```

### 3. Deploy to Azure

```pwsh
# Full deployment (infrastructure + code)
azd up

# Or code-only redeployment
azd deploy
```

### 4. Verify Deployment

```pwsh
# Get deployment details
azd env get-values

# Test the MCP endpoint
$functionAppName = azd env get-values | Select-String "AZURE_FUNCTION_NAME" | ForEach-Object { ($_ -split '=')[1].Trim('"') }
$mcpEndpoint = "https://$functionAppName.azurewebsites.net/runtime/webhooks/mcp/sse"

Write-Host "MCP Endpoint: $mcpEndpoint"
```

### 5. (Optional) Configure Entra ID Authentication

Only needed if using Azure AI Foundry with managed identity authentication:

```pwsh
# Get AI Foundry managed identity principal ID
$aiFoundryIdentity = az cognitiveservices account show `
  --name <your-ai-foundry-name> `
  --resource-group <your-resource-group> `
  --query identity.principalId -o tsv

# Configure authentication and grant app role
.\scripts\setup-auth.ps1 -AIFoundryManagedIdentityId $aiFoundryIdentity
```

## Post-Deployment Configuration

### Verify API Keys in Function App

```pwsh
$functionAppName = azd env get-values | Select-String "AZURE_FUNCTION_NAME" | ForEach-Object { ($_ -split '=')[1].Trim('"') }
$resourceGroup = azd env get-values | Select-String "AZURE_RESOURCE_GROUP_NAME" | ForEach-Object { ($_ -split '=')[1].Trim('"') }

# List app settings (verify API keys are present)
az functionapp config appsettings list `
  --name $functionAppName `
  --resource-group $resourceGroup `
  --query "[?name=='ALPHAVANTAGE_API_KEY' || name=='FRED_API_KEY' || name=='NEWSAPI_KEY'].{name:name, value:value}" `
  -o table
```

### Update API Keys (if needed)

```pwsh
az functionapp config appsettings set `
  --name $functionAppName `
  --resource-group $resourceGroup `
  --settings "ALPHAVANTAGE_API_KEY=new_key"
```

## Testing Deployed MCP Server

### Using MCP Inspector

```pwsh
# In package.json, update the inspector command with your deployed URL:
# "inspector": "npx @modelcontextprotocol/inspector https://<your-function-app>.azurewebsites.net/runtime/webhooks/mcp/sse"

yarn inspector
```

### Using GitHub Copilot Agent Mode

1. Press `F1` → **MCP: Add Server**
2. Transport type: **HTTP (Server-Sent Events)**
3. URL: `https://<your-function-app>.azurewebsites.net/runtime/webhooks/mcp/sse`
4. Add function key as header if using function key authentication
5. Test with Copilot agent: "Get SPX quote and recent market news"

## Troubleshooting

### API Keys Not Working

```pwsh
# Verify keys are set in Function App
az functionapp config appsettings list --name $functionAppName --resource-group $resourceGroup

# Restart Function App to pick up new settings
az functionapp restart --name $functionAppName --resource-group $resourceGroup
```

### MCP Endpoint Returns 401/403

- Check function keys: `az functionapp keys list --name $functionAppName --resource-group $resourceGroup`
- If using Entra ID auth, verify app registration and role assignments
- Confirm CORS settings allow your client origin

### Function App Not Responding

```pwsh
# Check Function App status
az functionapp show --name $functionAppName --resource-group $resourceGroup --query state

# View logs
az monitor app-insights query `
  --app <app-insights-name> `
  --analytics-query "traces | where timestamp > ago(1h) | order by timestamp desc | take 50"
```

## Cleanup

```pwsh
# Remove all Azure resources
azd down
```

## Additional Resources

- [Full Deployment Guide](./AZURE_DEPLOYMENT.md) - Comprehensive deployment documentation
- [MCP Tools Documentation](../README.md#mcp-tools) - List of all available tools
- [Architecture Diagram](../architecture-diagram.drawio) - Visual architecture overview
