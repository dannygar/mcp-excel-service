<#
.SYNOPSIS
    Deploy Trading AI Assistant Web App to Azure Container Apps

.DESCRIPTION
    This script deploys the Flask-based Trading AI Assistant Web App to Azure Container Apps with:
    - AI Foundry resource discovery and selection
    - Infrastructure provisioning via Bicep (ACR, Container Apps Environment, Log Analytics)
    - Docker image build and push to ACR
    - Container App configuration with AI Foundry connection
    - RBAC role assignment for managed identity
    - Health check and endpoint verification

.PARAMETER EnvironmentName
    The Azure Developer CLI environment name (default: "trading-webapp-dev")

.PARAMETER Location
    The Azure region for deployment (default: "eastus2")

.PARAMETER AIFoundryEndpoint
    The AI Foundry endpoint URL (optional - will be auto-discovered if not provided)

.PARAMETER AgentName
    The AI Foundry agent name (optional - will be auto-discovered if not provided)

.PARAMETER SkipInfrastructure
    Skip infrastructure provisioning (use for code-only deployments)

.PARAMETER SkipTest
    Skip endpoint testing after deployment

.PARAMETER SkipDiscovery
    Skip AI Foundry discovery (use existing configuration)

.PARAMETER ExistingResourceGroup
    Reuse an existing resource group (optional - useful for shared infrastructure)

.PARAMETER NonInteractive
    Run without prompts, automatically selecting first available options

.PARAMETER EnableEntraAuth
    Enable Microsoft Entra ID authentication for the web app

.EXAMPLE
    .\scripts\deploy-webapp.ps1  # Auto-discovers AI Foundry and deploys

.EXAMPLE
    .\scripts\deploy-webapp.ps1 -AIFoundryEndpoint "https://your-endpoint.services.ai.azure.com/api/projects/your-project"

.EXAMPLE
    .\scripts\deploy-webapp.ps1 -EnvironmentName "webapp-prod" -Location "westus2"

.EXAMPLE
    .\scripts\deploy-webapp.ps1 -SkipInfrastructure  # Code-only deployment

.EXAMPLE
    .\scripts\deploy-webapp.ps1 -NonInteractive  # Auto-select all options
#>

param(
    [string]$EnvironmentName = "trading-webapp-dev",
    [string]$Location = "eastus2",
    [string]$AIFoundryEndpoint = "",
    [string]$AgentName = "",
    [switch]$SkipInfrastructure,
    [switch]$SkipTest,
    [switch]$SkipDiscovery,
    [string]$ExistingResourceGroup = "",
    [switch]$NonInteractive,
    [switch]$EnableEntraAuth
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
║     Trading AI Assistant - Web App Container Deployment      ║
╚══════════════════════════════════════════════════════════════╝

"@ -ForegroundColor Magenta

# Get script and project root paths
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $ScriptPath
$WebAppPath = Join-Path $ProjectRoot "web-app"

Write-Host "Project Root:     $ProjectRoot"
Write-Host "Web App Path:     $WebAppPath"
Write-Host "Environment:      $EnvironmentName"
Write-Host "Location:         $Location"
Write-Host "Agent Name:       $AgentName"
Write-Host "Entra Auth:       $($EnableEntraAuth ? 'Enabled' : 'Disabled')"

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
Write-Host "  Subscription: $($azAccount.name)" -ForegroundColor Gray

# Check Azure Developer CLI
if (-not (Get-Command azd -ErrorAction SilentlyContinue)) {
    Write-Error "Azure Developer CLI (azd) is not installed. Please install from https://aka.ms/azd"
    exit 1
}
Write-Success "Azure Developer CLI found"

# Check Docker (optional but recommended)
$dockerAvailable = $false
$dockerCommand = Get-Command docker -ErrorAction SilentlyContinue
if ($dockerCommand) {
    $dockerVersion = docker version --format '{{.Server.Version}}' 2>$null
    if ($LASTEXITCODE -eq 0 -and $dockerVersion) {
        $dockerAvailable = $true
        Write-Success "Docker found (version: $dockerVersion)"
    } else {
        Write-Warning "Docker is installed but not running. Will use ACR cloud build."
    }
} else {
    Write-Warning "Docker not found. Will use ACR cloud build."
}

# =============================================================================
# Pre-Provision Hook: AI Foundry Discovery & Configuration
# =============================================================================

# Store AI Foundry resource info for RBAC assignment later
$AIFoundryResourceGroup = ""
$AIFoundryResourceName = ""

# Store Entra ID configuration
$EntraClientId = ""
$EntraTenantId = ""

$HooksPath = Join-Path $ProjectRoot "hooks"
$PreprovisionScript = Join-Path $HooksPath "preprovision.ps1"

if (-not $SkipDiscovery -and [string]::IsNullOrEmpty($AIFoundryEndpoint)) {
    # Use the hooks-based discovery system
    if (Test-Path $PreprovisionScript) {
        Write-Step "Running pre-provision hook for AI Foundry discovery..."
        
        $preprovisionParams = @{
            EnvironmentName = $EnvironmentName
        }
        if ($NonInteractive) {
            $preprovisionParams.NonInteractive = $true
        }
        
        $discoveryResult = & $PreprovisionScript @preprovisionParams
        
        if ($discoveryResult -and $discoveryResult.Success) {
            $AIFoundryEndpoint = $discoveryResult.AIFoundryEndpoint
            $AIFoundryResourceGroup = $discoveryResult.AIFoundryResourceGroup
            $AIFoundryResourceName = $discoveryResult.AIFoundryResourceName
            
            if ([string]::IsNullOrEmpty($AgentName) -and $discoveryResult.AgentName) {
                $AgentName = $discoveryResult.AgentName
            }
            
            # Get Entra ID config from preprovision result
            if ($discoveryResult.EntraClientId) {
                $EntraClientId = $discoveryResult.EntraClientId
            }
            if ($discoveryResult.EntraTenantId) {
                $EntraTenantId = $discoveryResult.EntraTenantId
            }
            
            Write-Success "Pre-provision hook completed successfully"
        } else {
            Write-Error "Pre-provision hook failed. Check the output above for details."
            exit 1
        }
    } else {
        # Fallback to legacy discovery method
        Write-Warning "Hooks not found at $HooksPath, using legacy discovery..."
        
        # Try to load from web-app/.env first
        $webAppEnvFile = Join-Path $WebAppPath ".env"
        if (Test-Path $webAppEnvFile) {
            Write-Host "  Checking: $webAppEnvFile" -ForegroundColor Gray
            $envContent = Get-Content $webAppEnvFile -Raw
            if ($envContent -match 'AI_FOUNDRY_ENDPOINT=(.+)') {
                $AIFoundryEndpoint = $matches[1].Trim()
            }
            if ($envContent -match 'AGENT_NAME=(.+)' -and [string]::IsNullOrEmpty($AgentName)) {
                $AgentName = $matches[1].Trim()
            }
            if ($envContent -match 'AI_FOUNDRY_RESOURCE_GROUP=(.+)') {
                $AIFoundryResourceGroup = $matches[1].Trim()
            }
            if ($envContent -match 'AI_FOUNDRY_RESOURCE_NAME=(.+)') {
                $AIFoundryResourceName = $matches[1].Trim()
            }
        }
        
        # If still no endpoint, run legacy discovery script
        if ([string]::IsNullOrEmpty($AIFoundryEndpoint)) {
            $discoverScript = Join-Path $ScriptPath "discover-ai-foundry.ps1"
            
            if (Test-Path $discoverScript) {
                if ($NonInteractive) {
                    $legacyResult = & $discoverScript -OutputJson | ConvertFrom-Json
                } else {
                    $legacyResult = & $discoverScript
                }
                
                if ($legacyResult.success) {
                    $AIFoundryEndpoint = $legacyResult.endpoint
                    $AIFoundryResourceGroup = $legacyResult.resourceGroup
                    $AIFoundryResourceName = $legacyResult.resourceName
                    
                    if ([string]::IsNullOrEmpty($AgentName) -and $legacyResult.agentName) {
                        $AgentName = $legacyResult.agentName
                    }
                } else {
                    Write-Error "AI Foundry discovery failed: $($legacyResult.error)"
                    exit 1
                }
            } else {
                Write-Error "No discovery mechanism available. Please provide -AIFoundryEndpoint"
                exit 1
            }
        }
    }
} elseif ($SkipDiscovery) {
    Write-Step "Skipping AI Foundry discovery (using existing configuration)..."
    
    # Load from web-app/.env
    $webAppEnvFile = Join-Path $WebAppPath ".env"
    if (Test-Path $webAppEnvFile) {
        $envContent = Get-Content $webAppEnvFile -Raw
        if ($envContent -match 'AI_FOUNDRY_ENDPOINT=(.+)') {
            $AIFoundryEndpoint = $matches[1].Trim()
        }
        if ($envContent -match 'AGENT_NAME=(.+)' -and [string]::IsNullOrEmpty($AgentName)) {
            $AgentName = $matches[1].Trim()
        }
        if ($envContent -match 'AI_FOUNDRY_RESOURCE_GROUP=(.+)') {
            $AIFoundryResourceGroup = $matches[1].Trim()
        }
        if ($envContent -match 'AI_FOUNDRY_RESOURCE_NAME=(.+)') {
            $AIFoundryResourceName = $matches[1].Trim()
        }
        # Load Entra ID config
        if ($envContent -match 'ENTRA_CLIENT_ID=(.+)') {
            $EntraClientId = $matches[1].Trim()
        }
        if ($envContent -match 'ENTRA_TENANT_ID=(.+)') {
            $EntraTenantId = $matches[1].Trim()
        }
    }
}

# Final validation
if ([string]::IsNullOrEmpty($AIFoundryEndpoint)) {
    Write-Error "AI Foundry endpoint not configured. Run without -SkipDiscovery or provide -AIFoundryEndpoint"
    exit 1
}

# Default agent name if still not set
if ([string]::IsNullOrEmpty($AgentName)) {
    $AgentName = "tastytrade-ledger"
}

# =============================================================================
# Entra ID Authentication Setup
# =============================================================================

if ($EnableEntraAuth) {
    Write-Step "Setting up Microsoft Entra ID authentication..."
    
    # If no Entra config exists, create the app registration
    if ([string]::IsNullOrEmpty($EntraClientId) -or [string]::IsNullOrEmpty($EntraTenantId)) {
        $entraModulePath = Join-Path $HooksPath "modules" "New-EntraAppRegistration.ps1"
        
        if (Test-Path $entraModulePath) {
            # Get tenant ID from current Azure context
            $EntraTenantId = (az account show --query "tenantId" -o tsv)
            $appName = "trading-webapp-$EnvironmentName"
            
            Write-Host "  Creating Entra ID app registration: $appName" -ForegroundColor Gray
            
            $EntraClientId = & $entraModulePath -AppName $appName -TenantId $EntraTenantId -RedirectUris @(
                "http://localhost:5000",
                "http://localhost:5000/auth/callback"
            )
            
            if ($EntraClientId) {
                Write-Success "Entra ID app registration created"
                Write-Host "  Client ID: $EntraClientId" -ForegroundColor Gray
                Write-Host "  Tenant ID: $EntraTenantId" -ForegroundColor Gray
                
                # Save to web-app/.env
                $webAppEnvFile = Join-Path $WebAppPath ".env"
                if (Test-Path $webAppEnvFile) {
                    $envContent = Get-Content $webAppEnvFile -Raw
                    if ($envContent -notmatch 'ENTRA_CLIENT_ID=') {
                        Add-Content -Path $webAppEnvFile -Value "`nENTRA_CLIENT_ID=$EntraClientId"
                    }
                    if ($envContent -notmatch 'ENTRA_TENANT_ID=') {
                        Add-Content -Path $webAppEnvFile -Value "ENTRA_TENANT_ID=$EntraTenantId"
                    }
                }
            } else {
                Write-Warning "Failed to create Entra ID app registration. Authentication will be disabled."
                $EnableEntraAuth = $false
            }
        } else {
            Write-Warning "Entra ID app registration module not found at $entraModulePath"
            Write-Host "  Please run: .\hooks\modules\New-EntraAppRegistration.ps1 manually" -ForegroundColor Gray
            $EnableEntraAuth = $false
        }
    } else {
        Write-Success "Using existing Entra ID configuration"
        Write-Host "  Client ID: $EntraClientId" -ForegroundColor Gray
        Write-Host "  Tenant ID: $EntraTenantId" -ForegroundColor Gray
    }
}

Write-Host ""
Write-Host "AI Foundry Configuration:" -ForegroundColor Cyan
Write-Host "  Endpoint:        $AIFoundryEndpoint" -ForegroundColor Gray
Write-Host "  Agent:           $AgentName" -ForegroundColor Gray
if ($AIFoundryResourceGroup) {
    Write-Host "  Resource Group:  $AIFoundryResourceGroup" -ForegroundColor Gray
}
if ($AIFoundryResourceName) {
    Write-Host "  Resource Name:   $AIFoundryResourceName" -ForegroundColor Gray
}
if ($EnableEntraAuth -and $EntraClientId) {
    Write-Host ""
    Write-Host "Entra ID Authentication:" -ForegroundColor Cyan
    Write-Host "  Client ID:       $EntraClientId" -ForegroundColor Gray
    Write-Host "  Tenant ID:       $EntraTenantId" -ForegroundColor Gray
}

# =============================================================================
# Azure Developer CLI Setup
# =============================================================================

Write-Step "Setting up Azure Developer CLI environment..."

Push-Location $ProjectRoot
try {
    # Check if environment exists
    $azdEnvExists = azd env list 2>$null | Select-String $EnvironmentName
    if (-not $azdEnvExists) {
        Write-Host "  Creating azd environment: $EnvironmentName"
        azd env new $EnvironmentName
    }
    
    # Select the environment
    azd env select $EnvironmentName
    
    # Set environment variables for Bicep parameters
    azd env set AZURE_LOCATION $Location
    azd env set AZURE_ENV_NAME $EnvironmentName
    azd env set AI_FOUNDRY_ENDPOINT $AIFoundryEndpoint
    azd env set AGENT_NAME $AgentName
    azd env set WEB_IMAGE_NAME ""
    azd env set EXISTING_RESOURCE_GROUP ([string]::IsNullOrEmpty($ExistingResourceGroup) ? "" : $ExistingResourceGroup)
    
    # Set Entra ID authentication variables
    if ($EnableEntraAuth -and $EntraClientId -and $EntraTenantId) {
        azd env set enableEntraAuth "true"
        azd env set ENTRA_CLIENT_ID $EntraClientId
        azd env set ENTRA_TENANT_ID $EntraTenantId
    } else {
        azd env set enableEntraAuth "false"
    }
    
    Write-Success "Environment configured"
}
finally {
    Pop-Location
}

# =============================================================================
# Infrastructure Deployment
# =============================================================================

if (-not $SkipInfrastructure) {
    Write-Step "Deploying infrastructure..."
    
    Push-Location $ProjectRoot
    
    try {
        # Deploy using Azure CLI directly (more reliable than azd provision)
        $bicepPath = Join-Path $ProjectRoot "infra/web-app/main.bicep"
        $resourceGroupName = "rg-webapp-$EnvironmentName"
        
        Write-Host "  Deploying Bicep template at subscription scope..." -ForegroundColor Gray
        
        # Deploy infrastructure
        $deploymentOutput = az deployment sub create `
            --name "webapp-deploy-$(Get-Date -Format 'yyyyMMddHHmmss')" `
            --location $Location `
            --template-file $bicepPath `
            --parameters environmentName=$EnvironmentName `
            --parameters location=$Location `
            --parameters aiFoundryEndpoint="$AIFoundryEndpoint" `
            --parameters agentName="$AgentName" `
            --parameters imageName="" `
            --parameters enableEntraAuth=$($EnableEntraAuth ? 'true' : 'false') `
            --parameters entraClientId="$EntraClientId" `
            --parameters entraTenantId="$EntraTenantId" `
            --output json 2>&1
        
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Deployment output: $deploymentOutput" -ForegroundColor Red
            throw "Infrastructure provisioning failed"
        }
        
        # Parse outputs (Bicep output names are uppercase)
        $deployment = $deploymentOutput | ConvertFrom-Json
        $resourceGroup = $deployment.properties.outputs.AZURE_RESOURCE_GROUP_NAME.value
        $acrName = $deployment.properties.outputs.AZURE_CONTAINER_REGISTRY_NAME.value
        $containerAppName = $deployment.properties.outputs.AZURE_CONTAINER_APP_NAME.value
        
        # Store values in azd environment for later use
        azd env set AZURE_RESOURCE_GROUP_NAME $resourceGroup 2>$null
        azd env set AZURE_CONTAINER_REGISTRY_NAME $acrName 2>$null
        azd env set AZURE_CONTAINER_APP_NAME $containerAppName 2>$null
        
        Write-Success "Infrastructure deployed"
        
        Write-Host "  Resource Group:  $resourceGroup" -ForegroundColor Gray
        Write-Host "  ACR:             $acrName" -ForegroundColor Gray
        Write-Host "  Container App:   $containerAppName" -ForegroundColor Gray
    }
    finally {
        Pop-Location
    }
} else {
    Write-Step "Skipping infrastructure (using existing)..."
    
    # Get existing values
    Push-Location $ProjectRoot
    azd env select $EnvironmentName 2>$null
    $resourceGroup = azd env get-value AZURE_RESOURCE_GROUP_NAME 2>$null
    $acrName = azd env get-value AZURE_CONTAINER_REGISTRY_NAME 2>$null
    $containerAppName = azd env get-value AZURE_CONTAINER_APP_NAME 2>$null
    Pop-Location
    
    if ([string]::IsNullOrEmpty($acrName) -or [string]::IsNullOrEmpty($containerAppName)) {
        Write-Error "Could not find existing infrastructure. Run without -SkipInfrastructure first."
        exit 1
    }
    
    Write-Success "Using existing infrastructure"
}

# =============================================================================
# Build and Deploy Container
# =============================================================================

Write-Step "Building and deploying container..."

# Generate unique image tag
$timestamp = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
$imageTag = "webapp-$timestamp"
$imageName = "$acrName.azurecr.io/trading-webapp:$imageTag"

Write-Host "  Image: $imageName" -ForegroundColor Gray

Push-Location $WebAppPath
try {
    if ($dockerAvailable) {
        # Local Docker build
        Write-Host "  Building Docker image locally..." -ForegroundColor Gray
        
        docker build -t $imageName .
        
        if ($LASTEXITCODE -ne 0) {
            throw "Docker build failed"
        }
        
        # Login to ACR and push
        Write-Host "  Logging in to ACR..." -ForegroundColor Gray
        az acr login --name $acrName
        
        Write-Host "  Pushing image to ACR..." -ForegroundColor Gray
        docker push $imageName
        
        if ($LASTEXITCODE -ne 0) {
            throw "Docker push failed"
        }
    } else {
        # ACR cloud build
        Write-Host "  Building image in ACR (cloud build)..." -ForegroundColor Gray
        
        az acr build `
            --registry $acrName `
            --image "trading-webapp:$imageTag" `
            --file ./Dockerfile `
            .
        
        if ($LASTEXITCODE -ne 0) {
            throw "ACR build failed"
        }
    }
    
    Write-Success "Container image built and pushed"
}
finally {
    Pop-Location
}

# =============================================================================
# Update Container App
# =============================================================================

Write-Step "Updating Container App..."

# Update the container app with the new image
# The container name is 'web' as defined in the Bicep template
az containerapp update `
    --name $containerAppName `
    --resource-group $resourceGroup `
    --container-name web `
    --image $imageName

if ($LASTEXITCODE -ne 0) {
    throw "Container App update failed"
}

# Update azd environment with new image name
Push-Location $ProjectRoot
azd env set WEB_IMAGE_NAME $imageName
Pop-Location

Write-Success "Container App updated"

# =============================================================================
# Post-Provision Hook: RBAC Assignment for AI Foundry Access
# =============================================================================

$PostprovisionScript = Join-Path $HooksPath "postprovision.ps1"

if (-not [string]::IsNullOrEmpty($AIFoundryResourceGroup) -and -not [string]::IsNullOrEmpty($AIFoundryResourceName)) {
    Write-Step "Running post-provision hook for RBAC configuration..."
    
    if (Test-Path $PostprovisionScript) {
        try {
            & $PostprovisionScript `
                -ContainerAppName $containerAppName `
                -ResourceGroup $resourceGroup `
                -AIFoundryResourceGroup $AIFoundryResourceGroup `
                -AIFoundryResourceName $AIFoundryResourceName
            
            Write-Success "Post-provision hook completed - Container App can access AI Foundry"
        }
        catch {
            Write-Warning "Post-provision hook may have failed: $_"
            Write-Host "  You may need to manually assign 'Cognitive Services User' role" -ForegroundColor Gray
            Write-Host "  Run: .\hooks\postprovision.ps1 -ContainerAppName $containerAppName -ResourceGroup $resourceGroup -AIFoundryResourceGroup $AIFoundryResourceGroup -AIFoundryResourceName $AIFoundryResourceName" -ForegroundColor Gray
        }
    } else {
        # Fallback to legacy RBAC script
        $rbacScript = Join-Path $ScriptPath "assign-ai-foundry-rbac.ps1"
        
        if (Test-Path $rbacScript) {
            try {
                & $rbacScript `
                    -ContainerAppName $containerAppName `
                    -ResourceGroup $resourceGroup `
                    -AIFoundryResourceGroup $AIFoundryResourceGroup `
                    -AIFoundryResourceName $AIFoundryResourceName
                
                Write-Success "RBAC configured - Container App can access AI Foundry"
            }
            catch {
                Write-Warning "RBAC assignment may have failed: $_"
                Write-Host "  You may need to manually assign 'Cognitive Services User' role" -ForegroundColor Gray
            }
        } else {
            Write-Warning "No RBAC script found. Manual role assignment may be required."
        }
    }
} else {
    Write-Warning "AI Foundry resource info not available. RBAC not configured."
    Write-Host "  To configure RBAC manually, run:" -ForegroundColor Gray
    Write-Host "  .\hooks\postprovision.ps1 -ContainerAppName $containerAppName -ResourceGroup $resourceGroup -AIFoundryResourceGroup <rg> -AIFoundryResourceName <name>" -ForegroundColor Gray
}

# =============================================================================
# Update Entra ID Redirect URIs (if authentication enabled)
# =============================================================================

if ($EnableEntraAuth -and $EntraClientId) {
    Write-Step "Updating Entra ID redirect URIs for deployed app..."
    
    try {
        # Get Container App FQDN
        $containerAppFqdn = az containerapp show `
            --name $containerAppName `
            --resource-group $resourceGroup `
            --query "properties.configuration.ingress.fqdn" `
            --output tsv 2>$null
        
        if ($containerAppFqdn) {
            $deployedUrl = "https://$containerAppFqdn"
            
            # Get existing redirect URIs
            $app = az ad app show --id $EntraClientId 2>$null | ConvertFrom-Json
            
            if ($app) {
                $existingUris = @($app.spa.redirectUris)
                $newUris = @($existingUris)
                
                # Add deployed URLs if not already present
                $urlsToAdd = @(
                    $deployedUrl,
                    "$deployedUrl/auth/callback"
                )
                
                foreach ($uri in $urlsToAdd) {
                    if ($uri -notin $newUris) {
                        $newUris += $uri
                    }
                }
                
                if ($newUris.Count -gt $existingUris.Count) {
                    # Update the app with new redirect URIs
                    $objectId = $app.id
                    
                    $spaBody = @{
                        spa = @{
                            redirectUris = $newUris
                        }
                    } | ConvertTo-Json -Depth 10
                    
                    $tempFile = [System.IO.Path]::GetTempFileName()
                    $spaBody | Out-File -FilePath $tempFile -Encoding utf8
                    
                    az rest --method PATCH `
                        --uri "https://graph.microsoft.com/v1.0/applications/$objectId" `
                        --headers "Content-Type=application/json" `
                        --body "@$tempFile" `
                        | Out-Null
                    
                    Remove-Item $tempFile -ErrorAction SilentlyContinue
                    
                    Write-Success "Added redirect URIs for deployed app:"
                    foreach ($uri in $urlsToAdd) {
                        Write-Host "    $uri" -ForegroundColor Gray
                    }
                } else {
                    Write-Success "Redirect URIs already configured"
                }
            } else {
                Write-Warning "Could not find Entra app registration: $EntraClientId"
            }
        } else {
            Write-Warning "Could not retrieve Container App FQDN"
        }
    }
    catch {
        Write-Warning "Failed to update Entra ID redirect URIs: $_"
        Write-Host "  You may need to manually add the deployed URL to the app registration" -ForegroundColor Gray
    }
}

# =============================================================================
# Get Deployment URL
# =============================================================================

Write-Step "Getting deployment URL..."

$containerAppUrl = az containerapp show `
    --name $containerAppName `
    --resource-group $resourceGroup `
    --query "properties.configuration.ingress.fqdn" `
    -o tsv

$fullUrl = "https://$containerAppUrl"

Write-Success "Web App deployed!"
Write-Host ""
Write-Host "  URL: $fullUrl" -ForegroundColor Green
Write-Host ""

# Update azd environment
Push-Location $ProjectRoot
azd env set WEB_ENDPOINT $fullUrl
Pop-Location

# =============================================================================
# Health Check
# =============================================================================

if (-not $SkipTest) {
    Write-Step "Running health check..."
    
    # Wait for container to start
    Write-Host "  Waiting for container to start (30 seconds)..." -ForegroundColor Gray
    Start-Sleep -Seconds 30
    
    $healthUrl = "$fullUrl/api/health"
    $maxRetries = 5
    $retryCount = 0
    $healthy = $false
    
    while ($retryCount -lt $maxRetries -and -not $healthy) {
        try {
            $response = Invoke-RestMethod -Uri $healthUrl -Method Get -TimeoutSec 10
            if ($response.status -eq "healthy") {
                $healthy = $true
                Write-Success "Health check passed!"
                Write-Host "  Agent: $($response.agent_name)" -ForegroundColor Gray
            }
        }
        catch {
            $retryCount++
            if ($retryCount -lt $maxRetries) {
                Write-Host "  Retry $retryCount/$maxRetries..." -ForegroundColor Yellow
                Start-Sleep -Seconds 10
            }
        }
    }
    
    if (-not $healthy) {
        Write-Warning "Health check did not pass. The container may still be starting."
        Write-Host "  Check logs: az containerapp logs show -n $containerAppName -g $resourceGroup"
    }
}

# =============================================================================
# Summary
# =============================================================================

Write-Host @"

╔══════════════════════════════════════════════════════════════╗
║                    Deployment Complete!                       ║
╚══════════════════════════════════════════════════════════════╝

"@ -ForegroundColor Green

Write-Host "Web App URL:     $fullUrl" -ForegroundColor Cyan
Write-Host "Resource Group:  $resourceGroup" -ForegroundColor Gray
Write-Host "Container App:   $containerAppName" -ForegroundColor Gray
Write-Host "ACR:             $acrName" -ForegroundColor Gray
if ($EnableEntraAuth -and $EntraClientId) {
    Write-Host ""
    Write-Host "Authentication:  Enabled (Entra ID)" -ForegroundColor Cyan
    Write-Host "  Client ID:     $EntraClientId" -ForegroundColor Gray
    Write-Host "  Tenant ID:     $EntraTenantId" -ForegroundColor Gray
} else {
    Write-Host "Authentication:  Disabled" -ForegroundColor Gray
}
Write-Host ""
Write-Host "Useful commands:" -ForegroundColor Yellow
Write-Host "  View logs:     az containerapp logs show -n $containerAppName -g $resourceGroup --follow"
Write-Host "  Restart:       az containerapp revision restart -n $containerAppName -g $resourceGroup"
Write-Host "  Scale:         az containerapp update -n $containerAppName -g $resourceGroup --min-replicas 1 --max-replicas 5"
Write-Host "  Redeploy:      .\scripts\deploy-webapp.ps1 -SkipInfrastructure"
if (-not $EnableEntraAuth) {
    Write-Host "  Enable Auth:   .\scripts\deploy-webapp.ps1 -EnableEntraAuth"
}
Write-Host ""

# Open browser
$openBrowser = Read-Host "Open web app in browser? (Y/n)"
if ($openBrowser -ne "n" -and $openBrowser -ne "N") {
    Start-Process $fullUrl
}
