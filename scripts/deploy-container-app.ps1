<#
.SYNOPSIS
    Deploy MCP Finance Server to Azure Container Apps

.DESCRIPTION
    This script deploys the MCP Finance Server to Azure Container Apps with:
    - Infrastructure provisioning via Bicep
    - Docker image build and push to ACR
    - Container App configuration with secrets
    - Health check and endpoint verification

.PARAMETER EnvironmentName
    The Azure Developer CLI environment name (default: "mcp-container")

.PARAMETER Location
    The Azure region for deployment (default: "eastus2")

.PARAMETER AlphaVantageApiKey
    The Alpha Vantage API key. If not provided, reads from ALPHAVANTAGE_API_KEY env var

.PARAMETER SkipInfrastructure
    Skip infrastructure provisioning (use for code-only deployments)

.PARAMETER SkipTest
    Skip endpoint testing after deployment

.EXAMPLE
    .\scripts\deploy-container-app.ps1

.EXAMPLE
    .\scripts\deploy-container-app.ps1 -EnvironmentName "prod" -Location "westus2"

.EXAMPLE
    .\scripts\deploy-container-app.ps1 -SkipInfrastructure
#>

param(
    [string]$EnvironmentName = "mcp-container",
    [string]$Location = "eastus2",
    [string]$AlphaVantageApiKey = "",
    [switch]$SkipInfrastructure,
    [switch]$SkipTest
)

$ErrorActionPreference = "Stop"

# Colors for output
function Write-Step { param($Message) Write-Host "`n▶ $Message" -ForegroundColor Cyan }
function Write-Success { param($Message) Write-Host "✓ $Message" -ForegroundColor Green }
function Write-Warning { param($Message) Write-Host "⚠ $Message" -ForegroundColor Yellow }
function Write-Error { param($Message) Write-Host "✗ $Message" -ForegroundColor Red }

# Banner
Write-Host @"

╔══════════════════════════════════════════════════════════════╗
║        MCP Finance Server - Container Apps Deployment        ║
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
# API Key Configuration
# =============================================================================

Write-Step "Configuring API keys..."

# Get API key from parameter, environment, or config file
if ([string]::IsNullOrEmpty($AlphaVantageApiKey)) {
    $AlphaVantageApiKey = $env:ALPHAVANTAGE_API_KEY
}

if ([string]::IsNullOrEmpty($AlphaVantageApiKey)) {
    $envDevPath = Join-Path $ProjectRoot "config\.env.dev"
    if (Test-Path $envDevPath) {
        $envContent = Get-Content $envDevPath -Raw
        if ($envContent -match 'ALPHAVANTAGE_API_KEY=(.+)') {
            $AlphaVantageApiKey = $matches[1].Trim()
        }
    }
}

if ([string]::IsNullOrEmpty($AlphaVantageApiKey)) {
    Write-Error "ALPHAVANTAGE_API_KEY not found. Please set it via:"
    Write-Host "  - Parameter: -AlphaVantageApiKey 'your-key'"
    Write-Host "  - Environment variable: `$env:ALPHAVANTAGE_API_KEY"
    Write-Host "  - Config file: config\.env.dev"
    exit 1
}

Write-Success "Alpha Vantage API key configured"

# Set as environment variable for azd
$env:ALPHAVANTAGE_API_KEY = $AlphaVantageApiKey

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
        azd env set ALPHAVANTAGE_API_KEY $AlphaVantageApiKey
        
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

$hasApiKey = $envVars | Where-Object { $_.name -eq "ALPHAVANTAGE_API_KEY" }
if ($hasApiKey) {
    Write-Success "ALPHAVANTAGE_API_KEY environment variable configured"
} else {
    Write-Warning "ALPHAVANTAGE_API_KEY not found in container env vars"
    Write-Host "  Adding environment variable..."
    
    az containerapp update `
        --name $containerAppName `
        --resource-group $resourceGroup `
        --set-env-vars "ALPHAVANTAGE_API_KEY=secretref:alphavantage-api-key" "PORT=3000"
    
    Write-Success "Environment variable added"
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
Write-Host "  Transport:        HTTP (Streamable HTTP)" -ForegroundColor Cyan
Write-Host ""
Write-Host "Available Tools:" -ForegroundColor White
Write-Host "  • get_company_revenue        - Get quarterly revenue data"
Write-Host "  • get_company_free_cash_flow - Get quarterly free cash flow data"
Write-Host ""
Write-Host "Supported Companies:" -ForegroundColor White
Write-Host "  • MSFT (Microsoft)"
Write-Host "  • TSLA (Tesla)"
Write-Host "  • NVDA (NVIDIA)"
Write-Host ""
Write-Host "Next Steps:" -ForegroundColor Yellow
Write-Host "  1. Add MCP server to Azure AI Foundry:"
Write-Host "     - URL: $mcpEndpoint"
Write-Host "     - Transport: Streamable HTTP"
Write-Host "     - Authentication: None (public endpoint)"
Write-Host ""
Write-Host "  2. Add to VS Code GitHub Copilot (mcp.json):"
Write-Host @"
     {
       "servers": {
         "mcp-finance": {
           "type": "http",
           "url": "$mcpEndpoint"
         }
       }
     }
"@ -ForegroundColor Gray
Write-Host ""
Write-Host "  3. Test in Foundry Chat Playground:"
Write-Host '     "What was Microsoft revenue for Q4 FY2024?"'
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
    deployedAt = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
}

$deploymentInfoPath = Join-Path $ProjectRoot ".azure\$EnvironmentName\deployment-info.json"
$deploymentInfo | ConvertTo-Json | Out-File -FilePath $deploymentInfoPath -Encoding UTF8
Write-Host "Deployment info saved to: $deploymentInfoPath" -ForegroundColor Gray
