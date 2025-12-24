# Adding MCP Excel Server to Azure AI Foundry Agent

This guide provides step-by-step instructions for adding the deployed MCP Excel Server to an Azure AI Foundry agent using Entra ID authentication.

## Prerequisites

Before adding the MCP server to a Foundry agent, ensure:

1. **MCP Server is deployed** - Run `.\scripts\deploy-mcp-server.ps1` to deploy to Azure Container Apps
2. **App Registration is configured** - The deployment script automatically creates an Entra ID app registration with Foundry-compatible authentication
3. **You have access to an Azure AI Foundry project** - You need Contributor or Owner access to add MCP tools

## Deployment Information

After running the deployment script, you'll receive these key values:

| Parameter | Description | Example |
|-----------|-------------|---------|
| **MCP Endpoint** | The URL of the deployed MCP server | `https://ca-mcp-xxxxx.eastus2.azurecontainerapps.io/mcp` |
| **Client ID (Audience)** | The App Registration Client ID used for authentication | `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx` |
| **Tenant ID** | Your Azure AD tenant ID | `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx` |

These values are also saved in `.azure/<environment>/deployment-info.json` after deployment.

## Step-by-Step Instructions

### Step 1: Open Azure AI Foundry Portal

Navigate to [https://ai.azure.com](https://ai.azure.com) and sign in with your Azure credentials.

### Step 2: Select Your Project

1. From the Foundry home page, select your AI project
2. You can also navigate directly to: `https://ai.azure.com/nextgen`

### Step 3: Create or Edit an Agent

1. Click on **Build** in the left navigation
2. Select **Create agent** or open an existing agent to edit

### Step 4: Add the MCP Tool

1. In the agent builder, locate the **Tools** section
2. Click **+ Add** to add a new tool
3. Select **+ Add a new tool** if prompted

### Step 5: Configure Model Context Protocol

1. Select the **Custom** tab
2. Select **Model Context Protocol (MCP)**
3. Click **Create**

### Step 6: Enter MCP Connection Details

Configure the MCP connection with the following values:

| Field | Value |
|-------|-------|
| **Name** | `MCP Excel Service` (or your preferred name) |
| **Remote MCP Server** | Your MCP endpoint URL from deployment (e.g., `https://ca-mcp-xxxxx.eastus2.azurecontainerapps.io/mcp`) |
| **Authentication** | Select **Microsoft Entra** |
| **Type** | Select **Project Managed Identity** |
| **Audience** | Your Client ID from deployment (e.g., `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`) |

### Step 7: Connect and Test

1. Click **Connect** to associate the MCP server with your agent
2. The connection will be validated
3. Once connected, you'll see the available tools listed:
   - `excel.appendRows` - Append rows to an Excel table
   - `excel.updateRange` - Update a range of cells in a worksheet

### Step 8: Test the Integration

In the Chat Playground, try prompts like:

- "Append sales data to my Excel file in SharePoint"
- "Update cells A1:B5 in my inventory spreadsheet with the new values"
- "Add a new row to the Products table in my workbook"

## Authentication Flow

When the agent calls the MCP server:

1. **Agent Authentication**: The Foundry agent uses its Project Managed Identity to obtain an access token
2. **Token Acquisition**: The token is requested with your App Registration Client ID as the audience
3. **Token Validation**: The MCP server validates the token using Microsoft Entra ID
4. **Graph API Call**: The MCP server uses its own service principal credentials to call Microsoft Graph API

```
┌─────────────────┐      ┌─────────────────┐      ┌─────────────────┐
│  Foundry Agent  │─────▶│   MCP Server    │─────▶│  Microsoft      │
│  (Chat/Agent)   │      │ (Container App) │      │  Graph API      │
└─────────────────┘      └─────────────────┘      └─────────────────┘
        │                        │                        │
        │ 1. Get token          │                        │
        │    (Managed ID)       │                        │
        │───────────────────────│                        │
        │                       │                        │
        │ 2. Call MCP with     │                        │
        │    token             │                        │
        │──────────────────────▶│                        │
        │                       │ 3. Call Graph API     │
        │                       │    (Service Principal)│
        │                       │───────────────────────▶│
        │                       │                        │
        │                       │◀───────────────────────│
        │◀──────────────────────│                        │
```

## Troubleshooting

### Connection Fails

**Symptom**: Cannot connect to the MCP server from Foundry

**Solutions**:
1. Verify the MCP endpoint URL is correct and accessible
2. Check that the Container App is running: `az containerapp show --name <app-name> --resource-group <rg>`
3. Ensure the health endpoint responds: `https://<fqdn>/health`

### Authentication Errors

**Symptom**: Token validation fails or unauthorized errors

**Solutions**:
1. Verify the Audience value matches the Client ID exactly
2. Ensure "Project Managed Identity" is selected as the Type
3. Check that the App Registration has the Application ID URI configured (`api://<client-id>`)

### Tool Not Working

**Symptom**: Agent connects but tools don't work

**Solutions**:
1. Verify Graph API permissions (Files.ReadWrite.All, Sites.ReadWrite.All) have admin consent
2. Check the MCP server logs: `az containerapp logs show --name <app-name> --resource-group <rg> --follow`
3. Ensure the Excel file path format is correct (drive ID and item ID required)

## Advanced Configuration

### Multiple Foundry Projects

To grant access from multiple Foundry projects:

1. Run the deployment script: it will auto-discover available AI projects
2. Select the projects that should have access
3. Each project's managed identity will be able to authenticate to the MCP server

### Custom Scaling

The MCP server is configured to scale from 1 to 5 replicas based on HTTP load. To modify:

1. Edit `infra/mcp-server/container-app.bicep`
2. Update the `scale` section:
   ```bicep
   scale: {
     minReplicas: 1  // Minimum replicas
     maxReplicas: 5  // Maximum replicas
     rules: [...]
   }
   ```
3. Redeploy with `azd up`

## Security Best Practices

1. **Principle of Least Privilege**: Only grant access to Foundry projects that need it
2. **Secret Rotation**: Rotate the client secret periodically (configured for 365-day expiration by default)
3. **Audit Logging**: Enable Application Insights for request logging and monitoring
4. **Network Security**: Consider adding IP restrictions or Virtual Network integration for production

## Related Documentation

- [Azure AI Foundry - Connect to MCP Servers](https://learn.microsoft.com/azure/ai-foundry/agents/how-to/tools/model-context-protocol)
- [MCP Authentication in Foundry](https://learn.microsoft.com/azure/ai-foundry/agents/how-to/mcp-authentication)
- [Deploy Remote MCP Server to Azure](https://learn.microsoft.com/azure/developer/azure-mcp-server/how-to/deploy-remote-mcp-server-microsoft-foundry)
