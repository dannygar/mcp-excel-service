<#
.SYNOPSIS
    Deploy MCP Excel Server to Azure Container Apps with Foundry Integration

.DESCRIPTION
    This script deploys the MCP Excel Server to Azure Container Apps with:
    - App Registration in Entra ID with Foundry-compatible authentication
    - Infrastructure provisioning via Bicep
    - Docker image build and push to ACR
    - Container App configuration with secrets (scales 1-5 replicas)
    - Auto-discovery of Azure AI Foundry projects for MCP integration
    - Health check and endpoint verification
    - Clear instructions for adding to Foundry agents

.PARAMETER EnvironmentName
    The Azure Developer CLI environment name (default: "mcp-container")

.PARAMETER Location
    The Azure region for deployment (default: "eastus2")

.PARAMETER AzureClientId
    The Azure AD Client ID. If not provided, reads from AZURE_CLIENT_ID env var or creates new app registration

.PARAMETER AzureClientSecret
    The Azure AD Client Secret. If not provided, reads from AZURE_CLIENT_SECRET env var

.PARAMETER SkipAppRegistration
    Skip app registration creation (use existing credentials)

.PARAMETER SkipInfrastructure
    Skip infrastructure provisioning (use for code-only deployments)

.PARAMETER SkipTest
    Skip endpoint testing after deployment

.PARAMETER SkipFoundryIntegration
    Skip Foundry project discovery and integration steps

.EXAMPLE
    .\scripts\deploy-mcp-server.ps1

.EXAMPLE
    .\scripts\deploy-mcp-server.ps1 -EnvironmentName "prod" -Location "westus2"

.EXAMPLE
    .\scripts\deploy-mcp-server.ps1 -SkipAppRegistration -SkipInfrastructure

.EXAMPLE
    .\scripts\deploy-mcp-server.ps1 -SkipFoundryIntegration
#>

param(
    [string]$EnvironmentName = "mcp-container",
    [string]$Location = "eastus2",
    [string]$AzureClientId = "",
    [string]$AzureClientSecret = "",
    [switch]$SkipAppRegistration,
    [switch]$SkipInfrastructure,
    [switch]$SkipTest,
    [switch]$SkipFoundryIntegration
)

$ErrorActionPreference = "Stop"

# Colors for output
function Write-Step { param($Message) Write-Host "`n▶ $Message" -ForegroundColor Cyan }
function Write-Success { param($Message) Write-Host "✓ $Message" -ForegroundColor Green }
function Write-Warning { param($Message) Write-Host "⚠ $Message" -ForegroundColor Yellow }
function Write-ErrorMsg { param($Message) Write-Host "✗ $Message" -ForegroundColor Red }
function Write-Info { param($Message) Write-Host "  $Message" -ForegroundColor Gray }

# Banner
Write-Host @"

╔══════════════════════════════════════════════════════════════╗
║     MCP Excel Server - Container Apps + Foundry Deployment   ║
╚══════════════════════════════════════════════════════════════╝

"@ -ForegroundColor Magenta

# Get script and project root paths
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $ScriptPath

Write-Host "Project Root: $ProjectRoot"
Write-Host "Environment:  $EnvironmentName"
Write-Host "Location:     $Location"

# =============================================================================
# Prerequisites Check
# =============================================================================

Write-Step "Checking prerequisites..."

# Check Azure CLI
if (-not (Get-Command az -ErrorAction SilentlyContinue)) {
    Write-Error "Azure CLI is not installed. Please install from https://docs.microsoft.com/cli/azure/install-azure-cli"
    exit 1
}

# Check Azure login
$azAccount = az account show 2>$null | ConvertFrom-Json
if (-not $azAccount) {
    Write-Warning "Not logged into Azure. Running 'az login'..."
    az login
    $azAccount = az account show | ConvertFrom-Json
}
Write-Success "Logged in as: $($azAccount.user.name)"

# Check Azure Developer CLI
if (-not (Get-Command azd -ErrorAction SilentlyContinue)) {
    Write-Error "Azure Developer CLI (azd) is not installed. Please install from https://aka.ms/azd"
    exit 1
}
Write-Success "Azure Developer CLI found"

# Check Docker (optional but recommended)
$dockerAvailable = Get-Command docker -ErrorAction SilentlyContinue
if ($dockerAvailable) {
    Write-Success "Docker found (for local testing)"
} else {
    Write-Warning "Docker not found. ACR will build the image remotely."
}

# =============================================================================
# App Registration / Credentials Configuration
# =============================================================================

Write-Step "Configuring Azure AD credentials..."

$azureTenantId = ""

# Get tenant ID from current login
$azAccount = az account show | ConvertFrom-Json
$azureTenantId = $azAccount.tenantId

# Get credentials from parameters, environment, or config file
if ([string]::IsNullOrEmpty($AzureClientId)) {
    $AzureClientId = $env:AZURE_CLIENT_ID
}

if ([string]::IsNullOrEmpty($AzureClientSecret)) {
    $AzureClientSecret = $env:AZURE_CLIENT_SECRET
}

# Try to load from config file
if ([string]::IsNullOrEmpty($AzureClientId) -or [string]::IsNullOrEmpty($AzureClientSecret)) {
    $envLocalPath = Join-Path $ProjectRoot "config\.env.local"
    if (Test-Path $envLocalPath) {
        $envContent = Get-Content $envLocalPath -Raw
        if ([string]::IsNullOrEmpty($AzureClientId) -and $envContent -match 'AZURE_CLIENT_ID=(.+)') {
            $AzureClientId = $matches[1].Trim()
        }
        if ([string]::IsNullOrEmpty($AzureClientSecret) -and $envContent -match 'AZURE_CLIENT_SECRET=(.+)') {
            $AzureClientSecret = $matches[1].Trim()
        }
        if ($envContent -match 'AZURE_TENANT_ID=(.+)') {
            $azureTenantId = $matches[1].Trim()
        }
    }
}

# Check if we have credentials
if ([string]::IsNullOrEmpty($AzureClientId) -or [string]::IsNullOrEmpty($AzureClientSecret)) {
    if ($SkipAppRegistration) {
        Write-ErrorMsg "Azure AD credentials not found and -SkipAppRegistration was specified."
        Write-Host "Please provide credentials via:"
        Write-Host "  - Parameters: -AzureClientId 'xxx' -AzureClientSecret 'xxx'"
        Write-Host "  - Environment variables: AZURE_CLIENT_ID, AZURE_CLIENT_SECRET"
        Write-Host "  - Config file: config\.env.local"
        exit 1
    }
    
    Write-Warning "Azure AD credentials not found. Creating new App Registration..."
    
    # Run the app registration script
    $registerScriptPath = Join-Path $ScriptPath "register-app.ps1"
    if (Test-Path $registerScriptPath) {
        $appResult = & $registerScriptPath
        
        if ($appResult) {
            $AzureClientId = $appResult.clientId
            $AzureClientSecret = $appResult.clientSecret
            $azureTenantId = $appResult.tenantId
            Write-Success "App Registration created successfully"
        } else {
            Write-ErrorMsg "Failed to create App Registration"
            exit 1
        }
    } else {
        Write-ErrorMsg "App registration script not found: $registerScriptPath"
        exit 1
    }
} else {
    Write-Success "Azure AD credentials configured"
}

Write-Host "  Tenant ID: $azureTenantId"
Write-Host "  Client ID: $AzureClientId"

# Set as environment variables for azd
$env:AZURE_TENANT_ID = $azureTenantId
$env:AZURE_CLIENT_ID = $AzureClientId
$env:AZURE_CLIENT_SECRET = $AzureClientSecret

# =============================================================================
# Foundry Project Discovery
# =============================================================================

$selectedFoundryProjects = @()

if (-not $SkipFoundryIntegration) {
    Write-Step "Discovering Azure AI Foundry projects..."
    
    try {
        # Query for AI Services (Cognitive Services) resources which include Foundry projects
        $foundryResources = az cognitiveservices account list --query "[?kind=='AIServices' || kind=='OpenAI']" -o json 2>$null | ConvertFrom-Json
        
        if ($foundryResources -and $foundryResources.Count -gt 0) {
            Write-Success "Found $($foundryResources.Count) Azure AI Foundry/AI Services resource(s):"
            
            $index = 1
            foreach ($resource in $foundryResources) {
                Write-Host "  [$index] $($resource.name) ($($resource.location)) - $($resource.kind)" -ForegroundColor White
                Write-Info "      Resource Group: $($resource.resourceGroup)"
                Write-Info "      Endpoint: $($resource.properties.endpoint)"
                $index++
            }
            
            Write-Host ""
            Write-Host "Select Foundry projects to grant access to this MCP Server." -ForegroundColor Yellow
            Write-Host "Enter project numbers separated by commas (e.g., 1,2,3), 'all' for all, or 'none' to skip:" -ForegroundColor Yellow
            $selection = Read-Host "Selection"
            
            if ($selection -eq 'all') {
                $selectedFoundryProjects = $foundryResources
                Write-Success "Selected all $($foundryResources.Count) projects"
            }
            elseif ($selection -ne 'none' -and $selection -ne '') {
                $indices = $selection -split ',' | ForEach-Object { [int]$_.Trim() }
                foreach ($idx in $indices) {
                    if ($idx -ge 1 -and $idx -le $foundryResources.Count) {
                        $selectedFoundryProjects += $foundryResources[$idx - 1]
                    }
                }
                Write-Success "Selected $($selectedFoundryProjects.Count) project(s)"
            }
            else {
                Write-Warning "No Foundry projects selected. You can configure access later manually."
            }
        }
        else {
            Write-Warning "No Azure AI Foundry/AI Services resources found in the current subscription."
            Write-Info "You can still deploy the MCP server and configure Foundry access later."
        }
    }
    catch {
        Write-Warning "Could not query Foundry projects: $($_.Exception.Message)"
        Write-Info "Continuing with deployment..."
    }
}

# =============================================================================
# Infrastructure Deployment
# =============================================================================

if (-not $SkipInfrastructure) {
    Write-Step "Deploying infrastructure with Azure Developer CLI..."
    
    Push-Location $ProjectRoot
    try {
        # Initialize azd environment if not exists
        $azdEnvExists = azd env list 2>$null | Select-String $EnvironmentName
        if (-not $azdEnvExists) {
            Write-Host "Creating azd environment: $EnvironmentName"
            azd env new $EnvironmentName
        }
        
        # Set environment variables
        azd env set AZURE_LOCATION $Location
        azd env set AZURE_TENANT_ID $azureTenantId
        azd env set AZURE_CLIENT_ID $AzureClientId
        azd env set AZURE_CLIENT_SECRET $AzureClientSecret
        
        # Deploy infrastructure and code
        Write-Host "Running 'azd up' - this may take 5-10 minutes..."
        azd up --environment $EnvironmentName
        
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Infrastructure deployment failed"
            exit 1
        }
        
        Write-Success "Infrastructure deployed successfully"
    }
    finally {
        Pop-Location
    }
} else {
    Write-Warning "Skipping infrastructure deployment (--SkipInfrastructure)"
}

# =============================================================================
# Get Deployment Outputs
# =============================================================================

Write-Step "Retrieving deployment information..."

Push-Location $ProjectRoot
try {
    # Get outputs from azd
    $outputs = azd env get-values --environment $EnvironmentName | Out-String
    
    # Parse outputs
    $resourceGroup = ($outputs | Select-String 'AZURE_RESOURCE_GROUP_NAME="([^"]+)"').Matches.Groups[1].Value
    $containerAppName = ($outputs | Select-String 'AZURE_CONTAINER_APP_NAME="([^"]+)"').Matches.Groups[1].Value
    $acrName = ($outputs | Select-String 'AZURE_CONTAINER_REGISTRY_NAME="([^"]+)"').Matches.Groups[1].Value
    $fqdn = ($outputs | Select-String 'AZURE_CONTAINER_APP_FQDN="([^"]+)"').Matches.Groups[1].Value
    $mcpEndpoint = ($outputs | Select-String 'MCP_ENDPOINT="([^"]+)"').Matches.Groups[1].Value
    
    if (-not $fqdn) {
        # Fallback: get from Azure directly
        $appInfo = az containerapp show --name $containerAppName --resource-group $resourceGroup -o json | ConvertFrom-Json
        $fqdn = $appInfo.properties.configuration.ingress.fqdn
        $mcpEndpoint = "https://$fqdn/mcp"
    }
    
    Write-Host ""
    Write-Host "Deployment Information:" -ForegroundColor White
    Write-Host "  Resource Group:    $resourceGroup"
    Write-Host "  Container App:     $containerAppName"
    Write-Host "  Registry:          $acrName"
    Write-Host "  FQDN:              $fqdn"
    Write-Host "  MCP Endpoint:      $mcpEndpoint"
}
finally {
    Pop-Location
}

# =============================================================================
# Verify Environment Variables
# =============================================================================

Write-Step "Verifying Container App configuration..."

$envVars = az containerapp show --name $containerAppName --resource-group $resourceGroup --query "properties.template.containers[0].env" -o json | ConvertFrom-Json

$hasClientId = $envVars | Where-Object { $_.name -eq "AZURE_CLIENT_ID" }
if ($hasClientId) {
    Write-Success "AZURE_CLIENT_ID environment variable configured"
} else {
    Write-Warning "AZURE_CLIENT_ID not found in container env vars"
}

# =============================================================================
# Health Check and Testing
# =============================================================================

if (-not $SkipTest) {
    Write-Step "Testing deployed endpoints..."
    
    # Wait for container to be ready
    Write-Host "Waiting for container to be ready..."
    Start-Sleep -Seconds 10
    
    # Check revision health
    $revisions = az containerapp revision list --name $containerAppName --resource-group $resourceGroup -o json | ConvertFrom-Json
    $activeRevision = $revisions | Where-Object { $_.properties.trafficWeight -gt 0 } | Select-Object -First 1
    
    if ($activeRevision.properties.healthState -eq "Healthy") {
        Write-Success "Container revision is healthy"
    } else {
        Write-Warning "Container revision health: $($activeRevision.properties.healthState)"
    }
    
    # Test health endpoint
    Write-Host "Testing health endpoint..."
    try {
        $healthResponse = Invoke-WebRequest -Uri "https://$fqdn/health" -UseBasicParsing -TimeoutSec 30
        if ($healthResponse.StatusCode -eq 200) {
            Write-Success "Health check passed: $($healthResponse.Content)"
        }
    }
    catch {
        Write-Warning "Health check failed: $($_.Exception.Message)"
    }
    
    # Test MCP endpoint
    Write-Host "Testing MCP endpoint..."
    try {
        $mcpResponse = Invoke-WebRequest -Uri $mcpEndpoint -UseBasicParsing -TimeoutSec 30 -ErrorAction SilentlyContinue
        Write-Success "MCP endpoint accessible (returned $($mcpResponse.StatusCode))"
    }
    catch {
        # MCP endpoint returns error without proper headers - this is expected
        if ($_.Exception.Response.StatusCode.value__ -eq 406) {
            Write-Success "MCP endpoint responding (requires MCP client headers)"
        } else {
            Write-Warning "MCP endpoint test: $($_.Exception.Message)"
        }
    }
}

# =============================================================================
# Summary
# =============================================================================

Write-Host @"

╔══════════════════════════════════════════════════════════════╗
║                    Deployment Complete!                      ║
╚══════════════════════════════════════════════════════════════╝

"@ -ForegroundColor Green

Write-Host "MCP Server Details:" -ForegroundColor White
Write-Host "  MCP Endpoint:     $mcpEndpoint" -ForegroundColor Cyan
Write-Host "  Health Endpoint:  https://$fqdn/health" -ForegroundColor Cyan
Write-Host "  Transport:        Streamable HTTP" -ForegroundColor Cyan
Write-Host "  Scaling:          1-5 replicas (auto-scale)" -ForegroundColor Cyan
Write-Host ""
Write-Host "Available Tools:" -ForegroundColor White
Write-Host "  • excel.appendRows  - Append rows to an Excel table"
Write-Host "  • excel.updateRange - Update a range of cells in a worksheet"
Write-Host ""
Write-Host "Authentication for Foundry:" -ForegroundColor White
Write-Host "  • Tenant ID:       $azureTenantId"
Write-Host "  • Client ID:       $AzureClientId"
Write-Host "  • Auth Type:       Microsoft Entra ID (Project Managed Identity)"
Write-Host ""

# =============================================================================
# Foundry Integration Instructions
# =============================================================================

Write-Host @"
╔══════════════════════════════════════════════════════════════╗
║          Steps to Add MCP Server to Foundry Agent            ║
╚══════════════════════════════════════════════════════════════╝

"@ -ForegroundColor Yellow

Write-Host "1. Open Azure AI Foundry Portal:" -ForegroundColor White
Write-Host "   https://ai.azure.com" -ForegroundColor Cyan
Write-Host ""
Write-Host "2. Navigate to your Foundry project and select 'Build → Create agent'" -ForegroundColor White
Write-Host ""
Write-Host "3. In the agent builder, click '+ Add' in the Tools section" -ForegroundColor White
Write-Host ""
Write-Host "4. Select the 'Custom' tab → 'Model Context Protocol' → 'Create'" -ForegroundColor White
Write-Host ""
Write-Host "5. Configure the MCP connection with these values:" -ForegroundColor White
Write-Host @"
   ┌────────────────────────────────────────────────────────────┐
   │ Name:                MCP Excel Service                     │
   │ Remote MCP Server:   $mcpEndpoint
   │ Authentication:      Microsoft Entra                       │
   │ Type:                Project Managed Identity              │
   │ Audience:            $AzureClientId
   └────────────────────────────────────────────────────────────┘
"@ -ForegroundColor Gray

Write-Host ""
Write-Host "6. Click 'Connect' to associate the MCP server with your agent" -ForegroundColor White
Write-Host ""
Write-Host "7. Test in the Chat Playground with prompts like:" -ForegroundColor White
Write-Host '   "Append sales data to my Excel file in SharePoint"' -ForegroundColor Gray
Write-Host '   "Update cells A1:B5 in my inventory spreadsheet"' -ForegroundColor Gray
Write-Host ""

if ($selectedFoundryProjects.Count -gt 0) {
    Write-Host "Selected Foundry Projects for Integration:" -ForegroundColor White
    foreach ($project in $selectedFoundryProjects) {
        Write-Host "  • $($project.name) - $($project.properties.endpoint)" -ForegroundColor Cyan
    }
    Write-Host ""
}

Write-Host @"
╔══════════════════════════════════════════════════════════════╗
║                    VS Code Integration                       ║
╚══════════════════════════════════════════════════════════════╝

"@ -ForegroundColor Yellow

Write-Host "Add to VS Code GitHub Copilot (.vscode/mcp.json):" -ForegroundColor White
Write-Host @"
{
  "servers": {
    "mcp-excel": {
      "type": "http",
      "url": "$mcpEndpoint"
    }
  }
}
"@ -ForegroundColor Gray
Write-Host ""

# Save deployment info to file
$deploymentInfo = @{
    environment = $EnvironmentName
    resourceGroup = $resourceGroup
    containerApp = $containerAppName
    registry = $acrName
    fqdn = $fqdn
    mcpEndpoint = $mcpEndpoint
    healthEndpoint = "https://$fqdn/health"
    tenantId = $azureTenantId
    clientId = $AzureClientId
    scaling = @{
        minReplicas = 1
        maxReplicas = 5
    }
    foundryConfig = @{
        authType = "Microsoft Entra"
        identityType = "Project Managed Identity"
        audience = $AzureClientId
    }
    selectedProjects = $selectedFoundryProjects | ForEach-Object { @{ name = $_.name; endpoint = $_.properties.endpoint } }
    deployedAt = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
}

$deploymentInfoPath = Join-Path $ProjectRoot ".azure\$EnvironmentName\deployment-info.json"
$deploymentInfoDir = Split-Path -Parent $deploymentInfoPath
if (-not (Test-Path $deploymentInfoDir)) {
    New-Item -ItemType Directory -Path $deploymentInfoDir -Force | Out-Null
}
$deploymentInfo | ConvertTo-Json -Depth 5 | Out-File -FilePath $deploymentInfoPath -Encoding UTF8
Write-Host "Deployment info saved to: $deploymentInfoPath" -ForegroundColor Gray
