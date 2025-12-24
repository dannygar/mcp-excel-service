<#
.SYNOPSIS
    Register or update an Entra ID App Registration for MCP Excel Service

.DESCRIPTION
    This script creates or updates an App Registration in Microsoft Entra ID with:
    - Microsoft Graph API permissions for Excel file operations
    - Client secret for service principal authentication
    - Foundry-compatible authentication (Application ID URI for managed identity)
    - Outputs credentials for local development and Azure deployment

.PARAMETER AppName
    The display name for the App Registration (default: "MCP Excel Service")

.PARAMETER CreateSecret
    Whether to create a new client secret (default: true)

.PARAMETER SecretExpirationDays
    Number of days until the client secret expires (default: 365)

.PARAMETER OutputEnvFile
    Path to output the .env file with credentials (default: config/.env.local)

.PARAMETER EnableFoundryAuth
    Enable Foundry Project Managed Identity authentication by configuring Application ID URI

.EXAMPLE
    .\scripts\register-app.ps1

.EXAMPLE
    .\scripts\register-app.ps1 -AppName "MCP Excel Service Prod" -SecretExpirationDays 180

.EXAMPLE
    .\scripts\register-app.ps1 -EnableFoundryAuth

.NOTES
    Required permissions: Application.ReadWrite.All or Application Administrator role
#>

param(
    [string]$AppName = "MCP Excel Service",
    [switch]$CreateSecret = $true,
    [int]$SecretExpirationDays = 365,
    [string]$OutputEnvFile = "",
    [switch]$EnableFoundryAuth = $true
)

$ErrorActionPreference = "Stop"

# Colors for output
function Write-Step { param($Message) Write-Host "`n▶ $Message" -ForegroundColor Cyan }
function Write-Success { param($Message) Write-Host "✓ $Message" -ForegroundColor Green }
function Write-Warning { param($Message) Write-Host "⚠ $Message" -ForegroundColor Yellow }
function Write-ErrorMsg { param($Message) Write-Host "✗ $Message" -ForegroundColor Red }

# Get script and project root paths
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $ScriptPath

if ([string]::IsNullOrEmpty($OutputEnvFile)) {
    $OutputEnvFile = Join-Path $ProjectRoot "config\.env.local"
}

# Banner
Write-Host @"

╔══════════════════════════════════════════════════════════════╗
║       MCP Excel Service - Entra ID App Registration          ║
╚══════════════════════════════════════════════════════════════╝

"@ -ForegroundColor Magenta

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

$tenantId = $azAccount.tenantId
Write-Host "  Tenant ID: $tenantId"

# =============================================================================
# Define Required Microsoft Graph Permissions
# =============================================================================

# Microsoft Graph API App ID (constant)
$graphApiAppId = "00000003-0000-0000-c000-000000000000"

# Required permissions for Excel operations
# Reference: https://learn.microsoft.com/en-us/graph/permissions-reference
$requiredPermissions = @(
    @{
        # Files.ReadWrite.All - Read and write files in all site collections
        Id = "75359482-378d-4052-8f01-80520e7db3cd"
        Type = "Role"  # Application permission
    },
    @{
        # Sites.ReadWrite.All - Read and write items in all site collections
        Id = "9492366f-7969-46a4-8d15-ed1a20078fff"
        Type = "Role"  # Application permission
    }
    # Optional: Sites.Selected - Access selected site collections (for more restrictive scenarios)
    # Add this if you want to use site-specific permissions instead:
    # @{
    #     Id = "883ea226-0bf2-4a8f-9f9d-92c9162a727d"
    #     Type = "Role"
    # }
)

# =============================================================================
# Check for Existing App Registration
# =============================================================================

Write-Step "Checking for existing App Registration..."

$existingApp = az ad app list --display-name $AppName --query "[0]" -o json 2>$null | ConvertFrom-Json

if ($existingApp) {
    Write-Warning "Found existing App Registration: $($existingApp.appId)"
    $appId = $existingApp.appId
    $objectId = $existingApp.id
    
    $updateChoice = Read-Host "Do you want to update the existing app? (y/n)"
    if ($updateChoice -ne 'y') {
        Write-Host "Exiting without changes."
        exit 0
    }
} else {
    Write-Host "No existing app found. Creating new App Registration..."
}

# =============================================================================
# Create or Update App Registration
# =============================================================================

Write-Step "Creating/Updating App Registration..."

if (-not $existingApp) {
    # Create new app registration
    $newApp = az ad app create `
        --display-name $AppName `
        --sign-in-audience "AzureADMyOrg" `
        --query "{appId:appId, id:id}" `
        -o json | ConvertFrom-Json
    
    $appId = $newApp.appId
    $objectId = $newApp.id
    
    Write-Success "Created App Registration"
    Write-Host "  App (Client) ID: $appId"
    Write-Host "  Object ID: $objectId"
    
    # Wait for propagation
    Write-Host "Waiting for Azure AD propagation..."
    Start-Sleep -Seconds 5
}

# =============================================================================
# Create Service Principal (if not exists)
# =============================================================================

Write-Step "Ensuring Service Principal exists..."

$existingSp = az ad sp show --id $appId 2>$null | ConvertFrom-Json
if (-not $existingSp) {
    Write-Host "Creating Service Principal..."
    az ad sp create --id $appId | Out-Null
    Write-Success "Service Principal created"
    Start-Sleep -Seconds 3
} else {
    Write-Success "Service Principal already exists"
}

# =============================================================================
# Configure Foundry Authentication (Application ID URI)
# =============================================================================

$identifierUri = ""

if ($EnableFoundryAuth) {
    Write-Step "Configuring Foundry Authentication..."
    
    # Set Application ID URI for Foundry Project Managed Identity authentication
    # This is required for Foundry agents to authenticate using their managed identity
    $identifierUri = "api://$appId"
    
    try {
        # Update app with identifier URI
        az ad app update `
            --id $appId `
            --identifier-uris $identifierUri
        
        if ($LASTEXITCODE -eq 0) {
            Write-Success "Application ID URI configured: $identifierUri"
            Write-Host "  This URI is used as the 'Audience' when configuring Foundry MCP connection"
        } else {
            Write-Warning "Could not set Application ID URI. You may need to configure this manually."
        }
        
        # Configure the app to accept tokens from Azure AD
        # This enables the Foundry managed identity to acquire tokens for this app
        Write-Host "Enabling app for access token issuance..."
        
        # Get current app manifest to update oauth2Permissions
        $appManifest = az ad app show --id $appId -o json | ConvertFrom-Json
        
        # Check if we need to add a default scope for the API
        $hasDefaultScope = $appManifest.api.oauth2PermissionScopes | Where-Object { $_.value -eq "user_impersonation" }
        
        if (-not $hasDefaultScope) {
            # Create a default scope for the API
            $defaultScope = @{
                adminConsentDescription = "Allow the application to access MCP Excel Service on behalf of the signed-in user."
                adminConsentDisplayName = "Access MCP Excel Service"
                id = [guid]::NewGuid().ToString()
                isEnabled = $true
                type = "User"
                userConsentDescription = "Allow the application to access MCP Excel Service on your behalf."
                userConsentDisplayName = "Access MCP Excel Service"
                value = "user_impersonation"
            }
            
            $apiConfig = @{
                oauth2PermissionScopes = @($defaultScope)
            }
            
            $apiJsonFile = [System.IO.Path]::GetTempFileName()
            $apiConfig | ConvertTo-Json -Depth 10 | Out-File -FilePath $apiJsonFile -Encoding UTF8
            
            try {
                az ad app update --id $appId --set api="@$apiJsonFile" 2>$null
                Write-Success "Default API scope configured"
            }
            catch {
                Write-Warning "Could not configure API scope. Foundry may still work with managed identity."
            }
            finally {
                if (Test-Path $apiJsonFile) {
                    Remove-Item $apiJsonFile -Force
                }
            }
        }
        
        Write-Success "Foundry authentication configured"
        Write-Host ""
        Write-Host "  IMPORTANT: When adding this MCP server to a Foundry agent:" -ForegroundColor Yellow
        Write-Host "    - Authentication: Microsoft Entra" -ForegroundColor White
        Write-Host "    - Type:           Project Managed Identity" -ForegroundColor White
        Write-Host "    - Audience:       $appId" -ForegroundColor Cyan
        Write-Host ""
    }
    catch {
        Write-Warning "Could not configure Foundry authentication: $($_.Exception.Message)"
        Write-Host "  You can configure this manually in the Azure Portal."
    }
}

# =============================================================================
# Configure API Permissions
# =============================================================================

Write-Step "Configuring Microsoft Graph API permissions..."

# Build the required resource access JSON
$resourceAccess = $requiredPermissions | ForEach-Object {
    @{
        id = $_.Id
        type = $_.Type
    }
}

$requiredResourceAccess = @(
    @{
        resourceAppId = $graphApiAppId
        resourceAccess = $resourceAccess
    }
)

# Write JSON to a temp file to avoid PowerShell quoting issues with Azure CLI
$tempJsonFile = [System.IO.Path]::GetTempFileName()
$requiredResourceAccess | ConvertTo-Json -Depth 10 | Out-File -FilePath $tempJsonFile -Encoding UTF8

try {
    # Update app with required permissions using the temp file
    az ad app update `
        --id $appId `
        --required-resource-accesses "@$tempJsonFile"
    
    if ($LASTEXITCODE -ne 0) {
        Write-Warning "Failed to configure permissions via Azure CLI. Please configure manually in Azure Portal."
    }
}
finally {
    # Clean up temp file
    if (Test-Path $tempJsonFile) {
        Remove-Item $tempJsonFile -Force
    }
}

Write-Success "API permissions configured:"
Write-Host "  • Files.ReadWrite.All (Application)"
Write-Host "  • Sites.ReadWrite.All (Application)"

# =============================================================================
# Grant Admin Consent
# =============================================================================

Write-Step "Granting admin consent for API permissions..."

Write-Warning "Admin consent is required for application permissions."
Write-Host "Attempting to grant admin consent automatically..."

try {
    # Grant admin consent for Microsoft Graph
    az ad app permission admin-consent --id $appId 2>$null
    Write-Success "Admin consent granted successfully"
}
catch {
    Write-Warning "Could not grant admin consent automatically."
    Write-Host ""
    Write-Host "Please grant admin consent manually:" -ForegroundColor Yellow
    Write-Host "  1. Go to: https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/$appId"
    Write-Host "  2. Click 'Grant admin consent for [Your Tenant]'"
    Write-Host "  3. Confirm the consent dialog"
    Write-Host ""
    Read-Host "Press Enter after granting admin consent to continue..."
}

# =============================================================================
# Create Client Secret
# =============================================================================

$clientSecret = ""

if ($CreateSecret) {
    Write-Step "Creating client secret..."
    
    $endDate = (Get-Date).AddDays($SecretExpirationDays).ToString("yyyy-MM-dd")
    
    $secretResult = az ad app credential reset `
        --id $appId `
        --append `
        --display-name "MCP Excel Service Secret" `
        --end-date $endDate `
        --query "{password:password}" `
        -o json | ConvertFrom-Json
    
    $clientSecret = $secretResult.password
    
    Write-Success "Client secret created (expires: $endDate)"
    Write-Warning "IMPORTANT: Save this secret now - it cannot be retrieved later!"
}

# =============================================================================
# Output Configuration
# =============================================================================

Write-Step "Saving configuration..."

# Ensure config directory exists
$configDir = Split-Path -Parent $OutputEnvFile
if (-not (Test-Path $configDir)) {
    New-Item -ItemType Directory -Path $configDir -Force | Out-Null
}

# Create .env file for local development
$envContent = @"
# MCP Excel Service - Entra ID Configuration
# Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
# App Registration: $AppName

# Azure AD / Entra ID Credentials
AZURE_TENANT_ID=$tenantId
AZURE_CLIENT_ID=$appId
AZURE_CLIENT_SECRET=$clientSecret

# Microsoft Graph API Scope (for client credentials flow)
GRAPH_SCOPE=https://graph.microsoft.com/.default

# Foundry Authentication (Application ID URI / Audience)
# Use this value as the 'Audience' when configuring MCP in Foundry
FOUNDRY_AUDIENCE=$appId

# Server Configuration
PORT=3000
HOST=0.0.0.0
"@

$envContent | Out-File -FilePath $OutputEnvFile -Encoding UTF8
Write-Success "Configuration saved to: $OutputEnvFile"

# Also save to .env in mcp-server folder for convenience
$mcpServerEnvPath = Join-Path $ProjectRoot "mcp-server\.env"
$envContent | Out-File -FilePath $mcpServerEnvPath -Encoding UTF8
Write-Success "Configuration also saved to: mcp-server\.env"

# =============================================================================
# Summary
# =============================================================================

Write-Host @"

╔══════════════════════════════════════════════════════════════╗
║              App Registration Complete!                      ║
╚══════════════════════════════════════════════════════════════╝

"@ -ForegroundColor Green

Write-Host "App Registration Details:" -ForegroundColor White
Write-Host "  Display Name:     $AppName" -ForegroundColor Cyan
Write-Host "  Tenant ID:        $tenantId" -ForegroundColor Cyan
Write-Host "  Client ID:        $appId" -ForegroundColor Cyan
if ($clientSecret) {
    Write-Host "  Client Secret:    ****$(($clientSecret).Substring([Math]::Max(0, $clientSecret.Length - 4)))" -ForegroundColor Cyan
}
if ($identifierUri) {
    Write-Host "  App ID URI:       $identifierUri" -ForegroundColor Cyan
}
Write-Host ""

Write-Host "Configured Permissions:" -ForegroundColor White
Write-Host "  • Files.ReadWrite.All   - Read/write files in SharePoint/OneDrive"
Write-Host "  • Sites.ReadWrite.All   - Read/write items in all site collections"
Write-Host ""

if ($EnableFoundryAuth) {
    Write-Host "Foundry Integration:" -ForegroundColor White
    Write-Host "  • Authentication:     Microsoft Entra (Project Managed Identity)" -ForegroundColor Cyan
    Write-Host "  • Audience:           $appId" -ForegroundColor Cyan
    Write-Host ""
}

Write-Host "Configuration Files:" -ForegroundColor White
Write-Host "  • $OutputEnvFile"
Write-Host "  • mcp-server\.env"
Write-Host ""

Write-Host "Next Steps:" -ForegroundColor Yellow
Write-Host "  1. Verify admin consent was granted in Azure Portal"
Write-Host "  2. For local development:"
Write-Host "     cd mcp-server && uv run python server.py"
Write-Host ""
Write-Host "  3. For Azure deployment:"
Write-Host "     .\scripts\deploy-mcp-server.ps1"
Write-Host ""

if ($EnableFoundryAuth) {
    Write-Host "  4. To add to Foundry Agent:" -ForegroundColor Yellow
    Write-Host "     - Go to https://ai.azure.com → Your Project → Build → Create agent"
    Write-Host "     - Click '+ Add' in Tools → Custom → Model Context Protocol"
    Write-Host "     - Configure:"
    Write-Host "       • Name: MCP Excel Service"
    Write-Host "       • Authentication: Microsoft Entra"
    Write-Host "       • Type: Project Managed Identity"
    Write-Host "       • Audience: $appId"
    Write-Host ""
}

# Output for automation
$result = @{
    tenantId = $tenantId
    clientId = $appId
    clientSecret = $clientSecret
    appName = $AppName
    envFile = $OutputEnvFile
    foundryAudience = $appId
    identifierUri = $identifierUri
}

# Save as JSON for other scripts to consume
$resultJsonPath = Join-Path $ProjectRoot ".azure\app-registration.json"
$resultDir = Split-Path -Parent $resultJsonPath
if (-not (Test-Path $resultDir)) {
    New-Item -ItemType Directory -Path $resultDir -Force | Out-Null
}
$result | ConvertTo-Json | Out-File -FilePath $resultJsonPath -Encoding UTF8
Write-Host "JSON output saved to: $resultJsonPath" -ForegroundColor Gray

# Return the result for use in other scripts
return $result
