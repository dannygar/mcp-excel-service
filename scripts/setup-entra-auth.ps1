<#
.SYNOPSIS
    Set up Microsoft Entra ID authentication for MCP Excel Service

.DESCRIPTION
    This script creates a Microsoft Entra ID app registration for authenticating
    AI Foundry agents to the MCP Excel Service.
    
    For Project Managed Identity auth (recommended for AI Foundry):
    - Creates an app registration with an app role
    - The AI Foundry project's managed identity must be granted this app role
    
    Output values should be set as environment variables or passed to deployment.

.PARAMETER AppName
    Display name for the Entra ID app registration. Default: "MCP Excel Service"

.PARAMETER FoundryProjectResourceId
    Azure resource ID of the AI Foundry project (optional - for auto role assignment)

.EXAMPLE
    .\scripts\setup-entra-auth.ps1
    
.EXAMPLE
    .\scripts\setup-entra-auth.ps1 -AppName "My MCP Server"

.EXAMPLE
    .\scripts\setup-entra-auth.ps1 -FoundryProjectResourceId "/subscriptions/.../projects/myproject"
#>

param(
    [string]$AppName = "MCP Excel Service",
    [string]$FoundryProjectResourceId = ""
)

$ErrorActionPreference = "Stop"

# Colors for output
function Write-Step { param($Message) Write-Host "`n▶ $Message" -ForegroundColor Cyan }
function Write-Success { param($Message) Write-Host "✓ $Message" -ForegroundColor Green }
function Write-Warn { param($Message) Write-Host "⚠ $Message" -ForegroundColor Yellow }
function Write-Err { param($Message) Write-Host "✗ $Message" -ForegroundColor Red }

# Banner
Write-Host @"

╔══════════════════════════════════════════════════════════════╗
║      MCP Excel Service - Microsoft Entra ID Setup            ║
║                                                              ║
║  For Azure AI Foundry Project Managed Identity Auth          ║
╚══════════════════════════════════════════════════════════════╝

"@ -ForegroundColor Magenta

# =============================================================================
# Prerequisites Check
# =============================================================================

Write-Step "Checking prerequisites..."

# Check Azure CLI
if (-not (Get-Command az -ErrorAction SilentlyContinue)) {
    Write-Err "Azure CLI is not installed. Please install from https://docs.microsoft.com/cli/azure/install-azure-cli"
    exit 1
}

# Check Azure login
$azAccount = az account show 2>$null | ConvertFrom-Json
if (-not $azAccount) {
    Write-Warn "Not logged into Azure. Running 'az login'..."
    az login
    $azAccount = az account show | ConvertFrom-Json
}
Write-Success "Logged in as: $($azAccount.user.name)"

$tenantId = $azAccount.tenantId
Write-Host "  Tenant ID: $tenantId"

# =============================================================================
# Check for existing app registration
# =============================================================================

Write-Step "Checking for existing app registration..."

$existingApps = az ad app list --display-name $AppName 2>$null | ConvertFrom-Json
$appId = $null
$objectId = $null

if ($existingApps -and $existingApps.Count -gt 0) {
    Write-Warn "Found existing app registration: $($existingApps[0].appId)"
    $useExisting = Read-Host "Use existing app? (y/n)"
    if ($useExisting -eq "y") {
        $appId = $existingApps[0].appId
        $objectId = $existingApps[0].id
        Write-Success "Using existing app: $appId"
    }
}

# =============================================================================
# Create App Registration
# =============================================================================

if (-not $appId) {
    Write-Step "Creating Entra ID app registration..."
    
    # Create the app registration (simple version without problematic flags)
    $appJson = az ad app create --display-name $AppName --sign-in-audience "AzureADMyOrg" 2>&1
    
    if ($LASTEXITCODE -ne 0) {
        Write-Err "Failed to create app registration: $appJson"
        exit 1
    }
    
    $app = $appJson | ConvertFrom-Json
    $appId = $app.appId
    $objectId = $app.id
    
    Write-Success "Created app registration"
    Write-Host "  Application (client) ID: $appId"
    Write-Host "  Object ID: $objectId"
}

# =============================================================================
# Configure App ID URI
# =============================================================================

Write-Step "Configuring Application ID URI..."

$appIdUri = "api://$appId"
az ad app update --id $appId --identifier-uris $appIdUri 2>$null

if ($LASTEXITCODE -eq 0) {
    Write-Success "Set Application ID URI: $appIdUri"
} else {
    Write-Warn "Could not set Application ID URI (may already be set)"
}

# =============================================================================
# Add App Role for MCP Tools Access
# =============================================================================

Write-Step "Adding app role for MCP tool access..."

$appRoleId = [guid]::NewGuid().ToString()
$appRoleJson = @{
    appRoles = @(
        @{
            id = $appRoleId
            displayName = "MCP Excel Service Tools ReadWrite All"
            description = "Application permission for MCP Excel Service tool calls"
            value = "Mcp.Tools.ReadWrite.All"
            isEnabled = $true
            allowedMemberTypes = @("Application")
        }
    )
} | ConvertTo-Json -Depth 10

# Write to temp file for az cli
$tempFile = [System.IO.Path]::GetTempFileName()
$appRoleJson | Out-File -FilePath $tempFile -Encoding utf8

# Update app with app role using Graph API
$graphUrl = "https://graph.microsoft.com/v1.0/applications/$objectId"
$result = az rest --method PATCH --uri $graphUrl --body "@$tempFile" --headers "Content-Type=application/json" 2>&1

Remove-Item $tempFile -Force -ErrorAction SilentlyContinue

if ($LASTEXITCODE -eq 0) {
    Write-Success "Added app role: Mcp.Tools.ReadWrite.All"
    Write-Host "  App Role ID: $appRoleId"
} else {
    Write-Warn "Could not add app role (may already exist): $result"
}

# =============================================================================
# Create Service Principal (if doesn't exist)
# =============================================================================

Write-Step "Ensuring service principal exists..."

$spJson = az ad sp show --id $appId 2>$null
if (-not $spJson) {
    $spResult = az ad sp create --id $appId 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Success "Created service principal"
        $spJson = az ad sp show --id $appId | ConvertFrom-Json
    } else {
        Write-Err "Failed to create service principal: $spResult"
    }
} else {
    Write-Success "Service principal already exists"
    $spJson = $spJson | ConvertFrom-Json
}

$servicePrincipalId = $spJson.id

# =============================================================================
# Grant App Role to Foundry Project (if provided)
# =============================================================================

if ($FoundryProjectResourceId) {
    Write-Step "Granting app role to AI Foundry project..."
    
    # Extract project info from resource ID
    # Format: /subscriptions/{sub}/resourceGroups/{rg}/providers/Microsoft.CognitiveServices/accounts/{account}/projects/{project}
    $parts = $FoundryProjectResourceId -split "/"
    if ($parts.Count -ge 11) {
        $projectSubscription = $parts[2]
        $projectResourceGroup = $parts[4]
        $accountName = $parts[8]
        $projectName = $parts[10]
        
        Write-Host "  Project: $projectName"
        Write-Host "  Account: $accountName"
        
        # Get the project's managed identity principal ID
        $projectJson = az cognitiveservices account show --name $accountName --resource-group $projectResourceGroup --subscription $projectSubscription 2>$null
        if ($projectJson) {
            $project = $projectJson | ConvertFrom-Json
            $projectPrincipalId = $project.identity.principalId
            
            if ($projectPrincipalId) {
                Write-Host "  Project Principal ID: $projectPrincipalId"
                
                # Assign app role to project's managed identity
                $roleAssignmentBody = @{
                    principalId = $projectPrincipalId
                    resourceId = $servicePrincipalId
                    appRoleId = $appRoleId
                } | ConvertTo-Json
                
                $tempFile2 = [System.IO.Path]::GetTempFileName()
                $roleAssignmentBody | Out-File -FilePath $tempFile2 -Encoding utf8
                
                $assignResult = az rest --method POST `
                    --uri "https://graph.microsoft.com/v1.0/servicePrincipals/$servicePrincipalId/appRoleAssignedTo" `
                    --body "@$tempFile2" `
                    --headers "Content-Type=application/json" 2>&1
                
                Remove-Item $tempFile2 -Force -ErrorAction SilentlyContinue
                
                if ($LASTEXITCODE -eq 0) {
                    Write-Success "Granted Mcp.Tools.ReadWrite.All role to Foundry project"
                } else {
                    Write-Warn "Could not assign role (may already be assigned): $assignResult"
                }
            } else {
                Write-Warn "Could not find project's managed identity"
            }
        } else {
            Write-Warn "Could not find Foundry project. You'll need to assign the role manually."
        }
    } else {
        Write-Warn "Invalid Foundry project resource ID format"
    }
}

# =============================================================================
# Output Summary
# =============================================================================

Write-Host @"

╔══════════════════════════════════════════════════════════════╗
║                    SETUP COMPLETE                            ║
╚══════════════════════════════════════════════════════════════╝

"@ -ForegroundColor Green

Write-Host "Configuration Values:" -ForegroundColor Yellow
Write-Host "═══════════════════════════════════════════════════════════════"
Write-Host ""
Write-Host "  ENTRA_TENANT_ID=$tenantId"
Write-Host "  ENTRA_CLIENT_ID=$appId"
Write-Host "  ENTRA_APP_IDENTIFIER_URI=$appIdUri"
Write-Host ""
Write-Host "═══════════════════════════════════════════════════════════════"

Write-Host ""
Write-Host "Next Steps:" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════════════"
Write-Host ""
Write-Host "1. Set environment variables for deployment:"
Write-Host "   `$env:ENTRA_TENANT_ID = '$tenantId'" -ForegroundColor White
Write-Host "   `$env:ENTRA_CLIENT_ID = '$appId'" -ForegroundColor White
Write-Host ""
Write-Host "2. Or save to azd environment:"
Write-Host "   azd env set ENTRA_TENANT_ID '$tenantId'" -ForegroundColor White
Write-Host "   azd env set ENTRA_CLIENT_ID '$appId'" -ForegroundColor White
Write-Host ""
Write-Host "3. Deploy the MCP server:"
Write-Host "   .\scripts\deploy-mcp-server.ps1" -ForegroundColor White
Write-Host ""
Write-Host "4. In AI Foundry, configure MCP tool connection:"
Write-Host "   - Authentication: Microsoft Entra → Project Managed Identity" -ForegroundColor White
Write-Host "   - Audience: $appId" -ForegroundColor White
Write-Host ""

if (-not $FoundryProjectResourceId) {
    Write-Host "5. Grant the app role to your AI Foundry project:" -ForegroundColor Yellow
    Write-Host "   - Go to Azure Portal → Entra ID → App Registrations"
    Write-Host "   - Find '$AppName' → Manage → App roles"
    Write-Host "   - Assign 'Mcp.Tools.ReadWrite.All' to your Foundry project's managed identity"
    Write-Host ""
}

Write-Host "═══════════════════════════════════════════════════════════════"
Write-Host ""

# Output for scripting
Write-Host "# For scripting, use these outputs:" -ForegroundColor DarkGray
Write-Host "# ENTRA_TENANT_ID=$tenantId" -ForegroundColor DarkGray
Write-Host "# ENTRA_CLIENT_ID=$appId" -ForegroundColor DarkGray
Write-Host "# ENTRA_APP_ROLE_ID=$appRoleId" -ForegroundColor DarkGray
Write-Host "# ENTRA_SERVICE_PRINCIPAL_ID=$servicePrincipalId" -ForegroundColor DarkGray
