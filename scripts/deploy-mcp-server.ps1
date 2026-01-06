<#
.SYNOPSIS
    Deploy MCP Excel Service to Azure Container Apps with Foundry Integration

.DESCRIPTION
    This script deploys the MCP Excel Service to Azure Container Apps with:
    - App Registration in Entra ID with Foundry-compatible authentication
    - Microsoft Entra ID token validation for secure access
    - Infrastructure provisioning via Bicep
    - Docker image build and push to ACR
    - Container App configuration with secrets (scales 1-5 replicas)
    - Auto-discovery of Azure AI Foundry projects for MCP integration
    - Health check and endpoint verification
    - Clear instructions for adding to Foundry agents

.PARAMETER EnvironmentName
    The environment name for deployment (dev or prod). Default: "dev"

.PARAMETER AzureEnvName
    The Azure Developer CLI environment name. Default: "mcp-excel-{EnvironmentName}"

.PARAMETER Location
    The Azure region for deployment (default: "eastus2")

.PARAMETER AzureClientId
    The Azure AD Client ID. If not provided, reads from AZURE_CLIENT_ID env var or creates new app registration

.PARAMETER AzureClientSecret
    The Azure AD Client Secret. If not provided, reads from AZURE_CLIENT_SECRET env var

.PARAMETER NoAuth
    Deploy WITHOUT Microsoft Entra ID authentication (public endpoint).
    By default, Entra auth is enabled. Use this flag to disable it.

.PARAMETER SkipAppRegistration
    Skip app registration creation (use existing credentials from env vars or config)

.PARAMETER SkipInfrastructure
    Skip infrastructure provisioning (use for code-only deployments)

.PARAMETER SkipTest
    Skip endpoint testing after deployment

.PARAMETER SkipFoundryIntegration
    Skip Foundry project discovery and integration steps

.EXAMPLE
    .\scripts\deploy-mcp-server.ps1
    # Deploys with Entra ID authentication (default)

.EXAMPLE
    .\scripts\deploy-mcp-server.ps1 -NoAuth
    # Deploys without authentication (public endpoint)

.EXAMPLE
    .\scripts\deploy-mcp-server.ps1 -EnvironmentName prod -Location "westus2"

.EXAMPLE
    .\scripts\deploy-mcp-server.ps1 -SkipAppRegistration -SkipInfrastructure

.EXAMPLE
    .\scripts\deploy-mcp-server.ps1 -SkipFoundryIntegration
#>

param(
    [ValidateSet("dev", "prod")]
    [string]$EnvironmentName = "dev",
    [string]$AzureEnvName = "",
    [string]$Location = "eastus2",
    [string]$AzureClientId = "",
    [string]$AzureClientSecret = "",
    [switch]$NoAuth,
    [switch]$SkipAppRegistration,
    [switch]$SkipInfrastructure,
    [switch]$SkipTest,
    [switch]$SkipFoundryIntegration
)

# Set AzureEnvName - prompt if not provided
if ([string]::IsNullOrEmpty($AzureEnvName)) {
    $defaultEnvName = "mcp-excel-$EnvironmentName"
    $userInput = Read-Host "Enter Azure environment name (press Enter for default '$defaultEnvName')"
    if ([string]::IsNullOrWhiteSpace($userInput)) {
        $AzureEnvName = $defaultEnvName
    } else {
        $AzureEnvName = $userInput.Trim()
    }
}

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
║   MCP Excel Service - Container Apps + Foundry Deployment    ║
╚══════════════════════════════════════════════════════════════╝

"@ -ForegroundColor Magenta

# Get script and project root paths
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $ScriptPath

Write-Host "Project Root:     $ProjectRoot"
Write-Host "Environment:      $EnvironmentName"
Write-Host "Azure Env Name:   $AzureEnvName"
Write-Host "Location:         $Location"

# =============================================================================
# Prerequisites Check
# =============================================================================

Write-Step "Checking prerequisites..."

# Check Azure CLI
if (-not (Get-Command az -ErrorAction SilentlyContinue)) {
    Write-ErrorMsg "Azure CLI is not installed. Please install from https://docs.microsoft.com/cli/azure/install-azure-cli"
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
    Write-ErrorMsg "Azure Developer CLI (azd) is not installed. Please install from https://aka.ms/azd"
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
# Trade Tracker Configuration (Optional)
# =============================================================================

Write-Step "Loading Trade Tracker configuration..."

# Get trade tracker settings from environment or config file
$TradeTrackerUrl = $env:TRADE_TRACKER_URL
$TradeTrackerFile = $env:TRADE_TRACKER_FILE

$envFilePath = Join-Path $ProjectRoot "config\.env.$EnvironmentName"
if (Test-Path $envFilePath) {
    Write-Info "Loading configuration from: $envFilePath"
    $envContent = Get-Content $envFilePath -Raw
    if ([string]::IsNullOrEmpty($TradeTrackerUrl) -and $envContent -match 'TRADE_TRACKER_URL=(.+)') {
        $TradeTrackerUrl = $matches[1].Trim()
    }
    if ([string]::IsNullOrEmpty($TradeTrackerFile) -and $envContent -match 'TRADE_TRACKER_FILE=(.+)') {
        $TradeTrackerFile = $matches[1].Trim()
    }
}

if (-not [string]::IsNullOrEmpty($TradeTrackerUrl)) {
    Write-Success "Trade Tracker URL configured"
    Write-Info "URL: $TradeTrackerUrl"
} else {
    Write-Info "Trade Tracker URL not configured (optional - can be set later)"
}

if (-not [string]::IsNullOrEmpty($TradeTrackerFile)) {
    Write-Success "Trade Tracker file configured"
    Write-Info "File: $TradeTrackerFile"
} else {
    Write-Info "Trade Tracker file not configured (optional - can be set later)"
}

# =============================================================================
# App Registration / Entra ID Configuration
# =============================================================================

if ($NoAuth) {
    Write-Step "Skipping Entra ID authentication (-NoAuth specified)..."
    Write-Warning "Deploying WITHOUT Entra authentication (MCP endpoint will be public)."
    Write-Host ""
    Write-Host "To enable Entra auth later, run:" -ForegroundColor Yellow
    Write-Host "  .\scripts\register-app.ps1" -ForegroundColor White
    Write-Host "  Then redeploy without -NoAuth flag" -ForegroundColor White
    $AzureClientId = ""
    $azureTenantId = ""
} else {
    Write-Step "Configuring Microsoft Entra ID authentication..."

    $azureTenantId = ""

    # Get tenant ID from current login
    $azAccount = az account show | ConvertFrom-Json
    $azureTenantId = $azAccount.tenantId

    # Get Entra credentials from parameters, environment, or config file
    # Note: For AI Foundry Project Managed Identity auth, we only need TENANT_ID and CLIENT_ID (no secret)
    if ([string]::IsNullOrEmpty($AzureClientId)) {
        $AzureClientId = $env:AZURE_CLIENT_ID
        if ([string]::IsNullOrEmpty($AzureClientId)) {
            $AzureClientId = $env:ENTRA_CLIENT_ID
        }
    }

    if ([string]::IsNullOrEmpty($AzureClientSecret)) {
        $AzureClientSecret = $env:AZURE_CLIENT_SECRET
    }

    # Try to load from environment-specific config file first, then fall back to .env.local
    $envFilePath = Join-Path $ProjectRoot "config\.env.$EnvironmentName"
    $envLocalPath = Join-Path $ProjectRoot "config\.env.local"
    
    $configFilePath = $null
    if (Test-Path $envFilePath) {
        $configFilePath = $envFilePath
        Write-Info "Loading credentials from: config\.env.$EnvironmentName"
    } elseif (Test-Path $envLocalPath) {
        $configFilePath = $envLocalPath
        Write-Info "Loading credentials from: config\.env.local (fallback)"
    }
    
    if ($configFilePath) {
        $envContent = Get-Content $configFilePath -Raw
        if ([string]::IsNullOrEmpty($AzureClientId) -and $envContent -match 'AZURE_CLIENT_ID=(.+)') {
            $AzureClientId = $matches[1].Trim()
        }
        if ([string]::IsNullOrEmpty($AzureClientId) -and $envContent -match 'ENTRA_CLIENT_ID=(.+)') {
            $AzureClientId = $matches[1].Trim()
        }
        if ([string]::IsNullOrEmpty($AzureClientSecret) -and $envContent -match 'AZURE_CLIENT_SECRET=(.+)') {
            $AzureClientSecret = $matches[1].Trim()
        }
        if ($envContent -match 'AZURE_TENANT_ID=(.+)') {
            $azureTenantId = $matches[1].Trim()
        }
        if ($envContent -match 'ENTRA_TENANT_ID=(.+)') {
            $azureTenantId = $matches[1].Trim()
        }
    }

    # Check if we have Entra Client ID (required for Foundry auth)
    if ([string]::IsNullOrEmpty($AzureClientId)) {
        if ($SkipAppRegistration) {
            Write-Warning "Azure AD Client ID not found and -SkipAppRegistration was specified."
            Write-Warning "Deploying WITHOUT Entra authentication (MCP endpoint will be public)."
            Write-Host ""
            Write-Host "To enable Entra auth later, run:" -ForegroundColor Yellow
            Write-Host "  .\scripts\register-app.ps1" -ForegroundColor White
            Write-Host "  Then set: azd env set ENTRA_CLIENT_ID <client-id>" -ForegroundColor White
        } else {
            Write-Warning "Azure AD Client ID not found. Creating new App Registration..."
            
            # Run the app registration script
            $registerScriptPath = Join-Path $ScriptPath "register-app.ps1"
            if (Test-Path $registerScriptPath) {
                Write-Host "Running register-app.ps1..." -ForegroundColor Gray
                & $registerScriptPath -AppName "MCP-Excel-Service-$EnvironmentName"
                
                if ($LASTEXITCODE -eq 0) {
                    # Script succeeded - prompt user to get the values
                    Write-Host ""
                    Write-Host "App Registration created. Please enter the values shown above:" -ForegroundColor Yellow
                    $AzureClientId = Read-Host "Enter ENTRA_CLIENT_ID"
                    
                    if (-not [string]::IsNullOrEmpty($AzureClientId)) {
                        Write-Success "Entra Client ID configured: $AzureClientId"
                    } else {
                        Write-Warning "No Client ID provided. Deploying without Entra auth."
                    }
                } else {
                    Write-Warning "App Registration script did not complete successfully."
                    Write-Warning "Deploying WITHOUT Entra authentication."
                }
            } else {
                Write-Warning "App registration script not found: $registerScriptPath"
                Write-Warning "Deploying WITHOUT Entra authentication."
            }
        }
    } else {
        Write-Success "Entra ID authentication configured"
    }

    if (-not [string]::IsNullOrEmpty($AzureClientId)) {
        Write-Info "Tenant ID: $azureTenantId"
        Write-Info "Client ID: $AzureClientId"
        Write-Info "Auth Mode: Token Validation (no secret needed for AI Foundry)"
    }
}

# Set as environment variables for azd (only if we have values)
$env:AZURE_TENANT_ID = $azureTenantId
if (-not [string]::IsNullOrEmpty($AzureClientId)) {
    $env:AZURE_CLIENT_ID = $AzureClientId
    $env:ENTRA_CLIENT_ID = $AzureClientId
    $env:ENTRA_TENANT_ID = $azureTenantId
}
if (-not [string]::IsNullOrEmpty($AzureClientSecret)) {
    $env:AZURE_CLIENT_SECRET = $AzureClientSecret
}

# =============================================================================
# Foundry Project Discovery
# =============================================================================

$selectedFoundryProjects = @()

if (-not $SkipFoundryIntegration -and -not $NoAuth) {
    Write-Step "Discovering Azure AI Foundry projects..."
    
    try {
        # Query for AI Services (Cognitive Services) resources which include Foundry projects
        # Include identity information for managed identity principal ID
        $foundryResources = az cognitiveservices account list --query "[?kind=='AIServices' || kind=='OpenAI'].{name:name, location:location, kind:kind, resourceGroup:resourceGroup, endpoint:properties.endpoint, principalId:identity.principalId}" -o json 2>$null | ConvertFrom-Json
        
        if ($foundryResources -and $foundryResources.Count -gt 0) {
            # Filter to only resources with managed identity
            $foundryWithIdentity = $foundryResources | Where-Object { $_.principalId }
            $foundryWithoutIdentity = $foundryResources | Where-Object { -not $_.principalId }
            
            if ($foundryWithIdentity.Count -gt 0) {
                Write-Success "Found $($foundryWithIdentity.Count) Azure AI Foundry/AI Services resource(s) with managed identity:"
                
                $index = 1
                foreach ($resource in $foundryWithIdentity) {
                    Write-Host "  [$index] $($resource.name) ($($resource.location)) - $($resource.kind)" -ForegroundColor White
                    Write-Info "      Resource Group: $($resource.resourceGroup)"
                    Write-Info "      Principal ID: $($resource.principalId)"
                    $index++
                }
                
                if ($foundryWithoutIdentity.Count -gt 0) {
                    Write-Host ""
                    Write-Warning "Found $($foundryWithoutIdentity.Count) resource(s) WITHOUT managed identity (cannot be used):"
                    foreach ($resource in $foundryWithoutIdentity) {
                        Write-Host "    - $($resource.name)" -ForegroundColor Gray
                    }
                }
                
                Write-Host ""
                Write-Host "Select Foundry projects to grant access to this MCP Server." -ForegroundColor Yellow
                Write-Host "Enter project numbers separated by commas (e.g., 1,2,3), 'all' for all, or 'none' to skip:" -ForegroundColor Yellow
                $selection = Read-Host "Selection"
                
                if ($selection -eq 'all') {
                    $selectedFoundryProjects = $foundryWithIdentity
                    Write-Success "Selected all $($foundryWithIdentity.Count) projects"
                }
                elseif ($selection -ne 'none' -and $selection -ne '') {
                    $indices = $selection -split ',' | ForEach-Object { [int]$_.Trim() }
                    foreach ($idx in $indices) {
                        if ($idx -ge 1 -and $idx -le $foundryWithIdentity.Count) {
                            $selectedFoundryProjects += $foundryWithIdentity[$idx - 1]
                        }
                    }
                    Write-Success "Selected $($selectedFoundryProjects.Count) project(s)"
                }
                else {
                    Write-Warning "No Foundry projects selected. You can configure access later with assign-app-role.ps1"
                }
            }
            else {
                Write-Warning "Found AI resources but none have managed identity enabled."
                Write-Info "Enable managed identity in Azure Portal → Your AI Resource → Identity → System assigned → On"
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
} elseif ($NoAuth) {
    Write-Info "Skipping Foundry discovery (-NoAuth specified, no authentication needed)"
}

# =============================================================================
# Verify azure.yaml exists
# =============================================================================

Write-Step "Verifying Azure Developer CLI configuration..."

$azureYamlPath = Join-Path $ProjectRoot "azure.yaml"
if (-not (Test-Path $azureYamlPath)) {
    Write-ErrorMsg "azure.yaml not found at $azureYamlPath"
    exit 1
}

Write-Success "azure.yaml found"

# =============================================================================
# Infrastructure Deployment
# =============================================================================

if (-not $SkipInfrastructure) {
    Write-Step "Deploying infrastructure with Azure Developer CLI..."
    
    Push-Location $ProjectRoot
    try {
        # Initialize azd environment if not exists
        $azdEnvExists = azd env list 2>$null | Select-String $AzureEnvName
        if (-not $azdEnvExists) {
            Write-Host "Creating azd environment: $AzureEnvName"
            azd env new $AzureEnvName
        }
        
        # Select the environment
        azd env select $AzureEnvName
        
        # Set environment variables
        azd env set AZURE_LOCATION $Location
        azd env set AZURE_ENV_NAME $AzureEnvName
        
        # Set Entra ID credentials (only if available)
        if (-not [string]::IsNullOrEmpty($azureTenantId)) {
            azd env set AZURE_TENANT_ID $azureTenantId
            azd env set ENTRA_TENANT_ID $azureTenantId
        }
        if (-not [string]::IsNullOrEmpty($AzureClientId)) {
            azd env set AZURE_CLIENT_ID $AzureClientId
            azd env set ENTRA_CLIENT_ID $AzureClientId
            # Enable Entra auth when client ID is provided
            azd env set ENABLE_ENTRA_AUTH true
        } else {
            azd env set ENABLE_ENTRA_AUTH false
        }
        if (-not [string]::IsNullOrEmpty($AzureClientSecret)) {
            azd env set AZURE_CLIENT_SECRET $AzureClientSecret
        }
        
        # Set Trade Tracker configuration (optional)
        if (-not [string]::IsNullOrEmpty($TradeTrackerUrl)) {
            azd env set TRADE_TRACKER_URL $TradeTrackerUrl
        }
        if (-not [string]::IsNullOrEmpty($TradeTrackerFile)) {
            azd env set TRADE_TRACKER_FILE $TradeTrackerFile
        }
        
        # Deploy infrastructure and code
        Write-Host "Running 'azd up' - this may take 5-10 minutes..."
        azd up --environment $AzureEnvName --no-prompt
        
        if ($LASTEXITCODE -ne 0) {
            Write-ErrorMsg "Infrastructure deployment failed"
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
    $outputs = azd env get-values --environment $AzureEnvName | Out-String
    
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

$hasEntraClientId = $envVars | Where-Object { $_.name -eq "ENTRA_CLIENT_ID" }
$hasEntraTenantId = $envVars | Where-Object { $_.name -eq "ENTRA_TENANT_ID" }
if ($hasEntraClientId -and $hasEntraTenantId) {
    Write-Success "Entra authentication configured (ENTRA_CLIENT_ID and ENTRA_TENANT_ID set)"
    $entraAuthEnabled = $true
} elseif (-not [string]::IsNullOrEmpty($AzureClientId)) {
    Write-Warning "Entra env vars not found in container - auth may not be enabled"
    Write-Info "Run 'azd up' again to update the container configuration"
    $entraAuthEnabled = $false
} else {
    Write-Info "Entra authentication not configured (no client ID provided)"
    $entraAuthEnabled = $false
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
        $statusCode = $_.Exception.Response.StatusCode.value__
        # MCP endpoint returns error without proper headers - this is expected
        if ($statusCode -eq 406) {
            Write-Success "MCP endpoint responding (requires MCP client headers)"
        } elseif ($statusCode -eq 401 -and $entraAuthEnabled) {
            Write-Success "MCP endpoint responding (401 expected - Entra auth enabled, requires bearer token)"
        } elseif ($statusCode -eq 401) {
            Write-Warning "MCP endpoint returned 401 Unauthorized (auth may be enabled but not detected)"
        } else {
            Write-Warning "MCP endpoint test: $($_.Exception.Message)"
        }
    }
}

# =============================================================================
# Assign App Roles to Selected Foundry Projects
# =============================================================================

$assignedProjects = @()
$failedProjects = @()

if ($selectedFoundryProjects.Count -gt 0 -and -not [string]::IsNullOrEmpty($AzureClientId)) {
    Write-Step "Assigning app roles to selected Foundry projects..."
    
    # Get service principal for our app registration
    $spJson = az ad sp show --id $AzureClientId 2>$null
    if ($spJson) {
        $sp = $spJson | ConvertFrom-Json
        $spObjectId = $sp.id
        
        # Find the app role
        $appRole = $sp.appRoles | Where-Object { $_.value -eq "Mcp.Tools.ReadWrite.All" }
        if (-not $appRole) {
            Write-Warning "App role 'Mcp.Tools.ReadWrite.All' not found on service principal"
            Write-Info "Run register-app.ps1 to create the app role"
        } else {
            $appRoleId = $appRole.id
            
            foreach ($project in $selectedFoundryProjects) {
                $projectName = $project.name
                $principalId = $project.principalId
                
                Write-Host "  Assigning role to: $projectName..." -NoNewline
                
                try {
                    # Check if assignment already exists
                    $existingAssignments = az rest --method GET `
                        --uri "https://graph.microsoft.com/v1.0/servicePrincipals/$principalId/appRoleAssignments" `
                        2>$null | ConvertFrom-Json
                    
                    $existingAssignment = $existingAssignments.value | Where-Object { 
                        $_.appRoleId -eq $appRoleId -and $_.resourceId -eq $spObjectId 
                    }
                    
                    if ($existingAssignment) {
                        Write-Host " Already assigned" -ForegroundColor Green
                        $assignedProjects += $project
                    } else {
                        # Create new assignment - use temp file for reliable JSON (az rest requires @file syntax)
                        $bodyObject = @{
                            principalId = $principalId
                            resourceId = $spObjectId
                            appRoleId = $appRoleId
                        }
                        $tempFile = [System.IO.Path]::GetTempFileName()
                        $bodyObject | ConvertTo-Json | Set-Content -Path $tempFile -Encoding UTF8
                        
                        $result = az rest --method POST `
                            --uri "https://graph.microsoft.com/v1.0/servicePrincipals/$principalId/appRoleAssignments" `
                            --headers "Content-Type=application/json" `
                            --body "@$tempFile" 2>$null | ConvertFrom-Json
                        
                        Remove-Item $tempFile -ErrorAction SilentlyContinue
                        
                        if ($result.id) {
                            Write-Host " ✓ Assigned" -ForegroundColor Green
                            $assignedProjects += $project
                        } else {
                            Write-Host " ✗ Failed" -ForegroundColor Red
                            $failedProjects += $project
                        }
                    }
                }
                catch {
                    Write-Host " ✗ Error: $($_.Exception.Message)" -ForegroundColor Red
                    $failedProjects += $project
                }
            }
            
            Write-Host ""
            if ($assignedProjects.Count -gt 0) {
                Write-Success "$($assignedProjects.Count) Foundry project(s) granted access to MCP server"
            }
            if ($failedProjects.Count -gt 0) {
                Write-Warning "$($failedProjects.Count) project(s) failed - run assign-app-role.ps1 manually"
            }
        }
    } else {
        Write-Warning "Could not find service principal for $AzureClientId"
        Write-Info "You may need to run assign-app-role.ps1 manually"
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
Write-Host "  MCP Endpoint:      $mcpEndpoint" -ForegroundColor Cyan
Write-Host "  Health Endpoint:   https://$fqdn/health" -ForegroundColor Cyan
Write-Host "  Transport:         Streamable HTTP" -ForegroundColor Cyan
Write-Host "  Scaling:           1-5 replicas (auto-scale)" -ForegroundColor Cyan
Write-Host ""
Write-Host "Available Tools:" -ForegroundColor White
Write-Host "  • excel.appendRows         - Append rows to an Excel table in SharePoint/OneDrive"
Write-Host "  • excel.updateRange        - Update a range of cells in an Excel worksheet"
Write-Host "  • trade.logTrades          - Log option trades to the Trade Tracker workbook"
Write-Host ""

if (-not [string]::IsNullOrEmpty($AzureClientId)) {
    Write-Host "Authentication for Foundry:" -ForegroundColor White
    Write-Host "  • Tenant ID:       $azureTenantId"
    Write-Host "  • Client ID:       $AzureClientId"
    Write-Host "  • Auth Type:       Microsoft Entra ID (Project Managed Identity)"
    Write-Host ""
} else {
    Write-Host "Authentication:" -ForegroundColor White
    Write-Host "  • Auth Type:       None (public endpoint)" -ForegroundColor Yellow
    Write-Host "  • Note:            Run register-app.ps1 to enable Entra auth" -ForegroundColor Yellow
    Write-Host ""
}

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

if (-not [string]::IsNullOrEmpty($AzureClientId)) {
    Write-Host @"
   ┌────────────────────────────────────────────────────────────┐
   │ Name:                MCP Excel Service                     │
   │ Remote MCP Server:   $mcpEndpoint
   │ Authentication:      Microsoft Entra                       │
   │ Type:                Project Managed Identity              │
   │ Audience:            $AzureClientId
   └────────────────────────────────────────────────────────────┘
"@ -ForegroundColor Gray
} else {
    Write-Host @"
   ┌────────────────────────────────────────────────────────────┐
   │ Name:                MCP Excel Service                     │
   │ Remote MCP Server:   $mcpEndpoint
   │ Authentication:      None                                  │
   └────────────────────────────────────────────────────────────┘
"@ -ForegroundColor Gray
    Write-Host ""
    Write-Host "   ⚠ WARNING: Endpoint is public. Enable Entra auth for production!" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "6. Click 'Connect' to associate the MCP server with your agent" -ForegroundColor White
Write-Host ""
Write-Host "7. Test in the Chat Playground with prompts like:" -ForegroundColor White
Write-Host '   "Append sales data to my Excel file in SharePoint"' -ForegroundColor Gray
Write-Host '   "Update cells A1:B5 in my inventory spreadsheet"' -ForegroundColor Gray
Write-Host '   "Log my latest option trade to the trade tracker"' -ForegroundColor Gray
Write-Host ""

if ($assignedProjects.Count -gt 0) {
    Write-Host "Foundry Projects with Access Granted:" -ForegroundColor White
    foreach ($project in $assignedProjects) {
        Write-Host "  ✓ $($project.name)" -ForegroundColor Green
    }
    Write-Host ""
}

if ($failedProjects.Count -gt 0) {
    Write-Host "Foundry Projects Needing Manual Setup:" -ForegroundColor Yellow
    foreach ($project in $failedProjects) {
        Write-Host "  ⚠ $($project.name) - Run: .\scripts\assign-app-role.ps1 -FoundryPrincipalId '$($project.principalId)'" -ForegroundColor Yellow
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
    "mcp-excel-service": {
      "type": "http",
      "url": "$mcpEndpoint"
    }
  }
}
"@ -ForegroundColor Gray
Write-Host ""

# Save deployment info to file
$foundryAuthConfig = @{
    authType = if ([string]::IsNullOrEmpty($AzureClientId)) { "None" } else { "Microsoft Entra" }
    identityType = if ([string]::IsNullOrEmpty($AzureClientId)) { "None" } else { "Project Managed Identity" }
    audience = $AzureClientId
}

$deploymentInfo = @{
    environment = $EnvironmentName
    azureEnvName = $AzureEnvName
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
    foundryConfig = $foundryAuthConfig
    selectedProjects = $selectedFoundryProjects | ForEach-Object { @{ name = $_.name; endpoint = $_.properties.endpoint } }
    deployedAt = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
}

$deploymentInfoPath = Join-Path $ProjectRoot ".azure\$AzureEnvName\deployment-info.json"
$deploymentInfoDir = Split-Path -Parent $deploymentInfoPath
if (-not (Test-Path $deploymentInfoDir)) {
    New-Item -ItemType Directory -Path $deploymentInfoDir -Force | Out-Null
}
$deploymentInfo | ConvertTo-Json -Depth 5 | Out-File -FilePath $deploymentInfoPath -Encoding UTF8
Write-Host "Deployment info saved to: $deploymentInfoPath" -ForegroundColor Gray
