# Azure Deployment Guide - MCP Finance Server

This guide walks you through deploying the MCP Finance Server to Azure Container Apps and connecting it to Azure AI Foundry.

## Overview

The MCP Finance Server is deployed as an Azure Container App with:

- **2 company earnings tools** for AI agents
- **FastMCP framework** with Streamable HTTP transport
- **Azure Container Registry** for Docker images
- **Application Insights** for monitoring
- **No authentication required** (public endpoint)

## Prerequisites

- **Azure CLI** installed and authenticated
- **Azure Developer CLI (azd)** installed
- **Docker** (optional, for local testing)
- **Alpha Vantage API key** ([Get free key](https://www.alphavantage.co/support/#api-key))

## Quick Deployment

### Step 1: Configure API Key

```pwsh
# Set your Alpha Vantage API key
$env:ALPHAVANTAGE_API_KEY = "your_api_key_here"
```

### Step 2: Deploy to Azure

```pwsh
# Deploy everything with one command
azd up

# Or use the deployment script for more control
.\scripts\deploy-container-app.ps1
```

### Step 3: Get Your MCP Endpoint

After deployment completes, note the MCP endpoint URL:

```
https://<container-app-name>.<environment>.azurecontainerapps.io/mcp
```

You can also retrieve it with:

```pwsh
azd env get-values | Select-String "MCP_ENDPOINT"
```

## Manual Deployment Steps

### 1. Initialize azd Environment

```pwsh
# Create a new environment
azd env new mcp-container

# Set required variables
azd env set AZURE_LOCATION eastus2
azd env set ALPHAVANTAGE_API_KEY "your_api_key"
```

### 2. Deploy Infrastructure

```pwsh
# Deploy infrastructure and code
azd up

# Or deploy infrastructure only
azd provision
```

### 3. Build and Push Container Image

If you need to rebuild the container manually:

```pwsh
# Get your ACR name
$acrName = azd env get-values | Select-String 'AZURE_CONTAINER_REGISTRY_NAME="([^"]+)"' | 
  ForEach-Object { $_.Matches.Groups[1].Value }

# Build and push using ACR
az acr build --registry $acrName --image mcp-finance-server:latest `
  --file container-app/Dockerfile container-app/
```

### 4. Verify Deployment

```pwsh
# Check health endpoint
$fqdn = azd env get-values | Select-String 'AZURE_CONTAINER_APP_FQDN="([^"]+)"' | 
  ForEach-Object { $_.Matches.Groups[1].Value }

Invoke-WebRequest -Uri "https://$fqdn/health"
```

## Connect to Azure AI Foundry

### Add MCP Server to Foundry Agent

1. Navigate to [Azure AI Foundry](https://ai.azure.com)
2. Open your AI Project
3. Go to **Agent** → **Tools** → **MCP Servers**
4. Click **Add MCP Server**
5. Configure:

| Setting | Value |
|---------|-------|
| **Name** | `mcp-finance` |
| **Transport Type** | Streamable HTTP |
| **Endpoint URL** | `https://<your-app>.azurecontainerapps.io/mcp` |
| **Authentication** | None |

6. Click **Save**

### Test in Chat Playground

Try these prompts:

- "What was Microsoft's revenue for Q4 FY2024?"
- "Get Tesla's free cash flow for Q2 FY2024"
- "Compare NVIDIA's revenue between Q3 and Q4 FY2024"

## Connect to VS Code GitHub Copilot

Add to `.vscode/mcp.json`:

```json
{
  "servers": {
    "mcp-finance": {
      "type": "http",
      "url": "https://<your-app>.azurecontainerapps.io/mcp"
    }
  }
}
```

## Infrastructure Details

### Resources Created

| Resource | Description |
|----------|-------------|
| Resource Group | `rg-<environment-name>` |
| Container Registry | Stores Docker images |
| Container Apps Environment | Hosts the container app |
| Container App | Runs the MCP server |
| Log Analytics Workspace | Collects logs |
| Application Insights | Monitors performance |

### Container Configuration

| Setting | Value |
|---------|-------|
| **Image** | `mcp-finance-server:latest` |
| **Port** | 3000 |
| **CPU** | 0.5 cores |
| **Memory** | 1 Gi |
| **Min Replicas** | 1 |
| **Max Replicas** | 10 |

### Endpoints

| Endpoint | URL |
|----------|-----|
| **MCP** | `https://<fqdn>/mcp` |
| **Health** | `https://<fqdn>/health` |

## Environment Variables

| Variable | Description | Required |
|----------|-------------|----------|
| `ALPHAVANTAGE_API_KEY` | Alpha Vantage API key | Yes |
| `PORT` | Server port (default: 3000) | No |
| `HOST` | Server host (default: 0.0.0.0) | No |

## Monitoring

### View Logs

```pwsh
# Stream logs
az containerapp logs show `
  --name <container-app-name> `
  --resource-group <resource-group> `
  --follow

# Or use azd
azd monitor --logs
```

### View Metrics in Azure Portal

1. Navigate to your Container App in Azure Portal
2. Go to **Monitoring** → **Metrics**
3. View request counts, response times, CPU/memory usage

### Application Insights

The deployment includes Application Insights for:

- Request tracing
- Exception logging
- Performance metrics
- Custom events

## Scaling

The Container App is configured with HTTP-based autoscaling:

- **Min replicas**: 1 (always running)
- **Max replicas**: 10
- **Scale trigger**: HTTP concurrent requests

To modify scaling:

```pwsh
az containerapp update `
  --name <container-app-name> `
  --resource-group <resource-group> `
  --min-replicas 1 `
  --max-replicas 20
```

## Troubleshooting

### Container Won't Start

```pwsh
# Check logs
az containerapp logs show --name <app> --resource-group <rg> --tail 100

# Check revision health
az containerapp revision list --name <app> --resource-group <rg> -o table
```

### API Key Not Working

1. Verify the secret exists:
   ```pwsh
   az containerapp show --name <app> --resource-group <rg> `
     --query "properties.configuration.secrets[].name"
   ```

2. Verify environment variable mapping:
   ```pwsh
   az containerapp show --name <app> --resource-group <rg> `
     --query "properties.template.containers[0].env"
   ```

3. Update if missing:
   ```pwsh
   az containerapp update --name <app> --resource-group <rg> `
     --set-env-vars "ALPHAVANTAGE_API_KEY=secretref:alphavantage-api-key"
   ```

### MCP Tools Return Errors

- Check Alpha Vantage rate limits (25/day on free tier)
- Verify API key is valid
- Check container logs for detailed error messages

### Health Check Failing

1. Test health endpoint directly:
   ```pwsh
   Invoke-WebRequest -Uri "https://<fqdn>/health"
   ```

2. Check that probes are configured correctly:
   ```pwsh
   az containerapp show --name <app> --resource-group <rg> `
     --query "properties.template.containers[0].probes"
   ```

## Redeployment

### Code-Only Deployment

```pwsh
# Rebuild and deploy new container image
azd deploy
```

### Full Redeployment

```pwsh
# Provision infrastructure + deploy code
azd up
```

### Manual Container Update

```pwsh
# Build new image
az acr build --registry <acr-name> --image mcp-finance-server:v2 `
  --file container-app/Dockerfile container-app/

# Update container app
az containerapp update --name <app> --resource-group <rg> `
  --image <acr-name>.azurecr.io/mcp-finance-server:v2
```

## Clean Up

```pwsh
# Delete all resources
azd down

# Or delete resource group directly
az group delete --name rg-<environment-name> --yes
```

## Security Considerations

### Current Setup (Development)

- Public endpoint (no authentication)
- API key stored as Container App secret
- HTTPS only (TLS 1.2+)
- CORS enabled for all origins

### Production Recommendations

1. **Enable Authentication**: Add Microsoft Entra ID authentication
2. **Restrict CORS**: Limit allowed origins
3. **Use Private Endpoints**: Deploy in a VNet with private endpoints
4. **Key Vault**: Store secrets in Azure Key Vault
5. **Network Restrictions**: Use IP allow lists

### Add Entra ID Authentication (Optional)

```pwsh
# Enable Easy Auth
az containerapp auth microsoft update `
  --name <container-app-name> `
  --resource-group <resource-group> `
  --client-id <app-registration-client-id> `
  --issuer "https://login.microsoftonline.com/<tenant-id>/v2.0" `
  --yes
```

## References

- [Azure Container Apps Documentation](https://learn.microsoft.com/azure/container-apps/)
- [FastMCP Documentation](https://gofastmcp.com/)
- [Model Context Protocol](https://modelcontextprotocol.io/)
- [Azure AI Foundry MCP Integration](https://learn.microsoft.com/azure/ai-foundry/mcp/)
- [Alpha Vantage API](https://www.alphavantage.co/documentation/)
