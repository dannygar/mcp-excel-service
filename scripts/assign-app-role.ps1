<#
.SYNOPSIS
    Assigns an app role to an AI Foundry project's managed identity

.DESCRIPTION
    This script assigns a specified app role (default: Mcp.Tools.ReadWrite.All) 
    to an AI Foundry project's managed identity, allowing the project to 
    authenticate with the MCP server.

.PARAMETER FoundryPrincipalId
    The Object (Principal) ID of the AI Foundry project's managed identity.
    Find this in: Azure Portal → AI Foundry → Your Project → Settings → Identity

.PARAMETER AppClientId
    The Application (Client) ID of the MCP server's app registration.
    Default: Uses ENTRA_CLIENT_ID from azd environment or environment variable.

.PARAMETER AppRoleName
    The name of the app role to assign.
    Default: "Mcp.Tools.ReadWrite.All"

.EXAMPLE
    .\scripts\assign-app-role.ps1 -FoundryPrincipalId "6e8745b1-0624-4cf9-847e-d21c3399dc20"

.EXAMPLE
    .\scripts\assign-app-role.ps1 -FoundryPrincipalId "6e8745b1-..." -AppClientId "4cb6750d-..."

.EXAMPLE
    .\scripts\assign-app-role.ps1 -FoundryPrincipalId "6e8745b1-..." -AppRoleName "Mcp.Tools.Read"
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$FoundryPrincipalId,
    
    [Parameter(Mandatory = $false)]
    [string]$AppClientId = "",
    
    [Parameter(Mandatory = $false)]
    [string]$AppRoleName = "Mcp.Tools.ReadWrite.All"
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
║          Assign App Role to Foundry Managed Identity         ║
╚══════════════════════════════════════════════════════════════╝

"@ -ForegroundColor Magenta

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

# =============================================================================
# Get App Client ID
# =============================================================================

Write-Step "Resolving App Client ID..."

if ([string]::IsNullOrEmpty($AppClientId)) {
    # Try environment variable first
    $AppClientId = $env:ENTRA_CLIENT_ID
    
    # Try azd environment
    if ([string]::IsNullOrEmpty($AppClientId)) {
        $AppClientId = (azd env get-value ENTRA_CLIENT_ID 2>$null)
    }
    
    if ([string]::IsNullOrEmpty($AppClientId)) {
        Write-Error "App Client ID not found. Please provide -AppClientId or set ENTRA_CLIENT_ID"
        Write-Host ""
        Write-Host "You can find this in:"
        Write-Host "  - Azure Portal → Microsoft Entra ID → App Registrations → Your App → Application (client) ID"
        Write-Host "  - Or run: azd env get-value ENTRA_CLIENT_ID"
        exit 1
    }
}

Write-Success "App Client ID: $AppClientId"

# =============================================================================
# Get Service Principal for App Registration
# =============================================================================

Write-Step "Looking up service principal for app registration..."

$spJson = az ad sp show --id $AppClientId 2>$null
if (-not $spJson) {
    Write-Error "Service principal not found for App Client ID: $AppClientId"
    Write-Host "Ensure the app registration exists and has a service principal."
    exit 1
}

$sp = $spJson | ConvertFrom-Json
$spObjectId = $sp.id
$spAppRoles = $sp.appRoles

Write-Success "Found service principal: $($sp.displayName)"
Write-Host "  Object ID: $spObjectId"

# =============================================================================
# Find App Role
# =============================================================================

Write-Step "Finding app role: $AppRoleName..."

$appRole = $spAppRoles | Where-Object { $_.value -eq $AppRoleName }

if (-not $appRole) {
    Write-Error "App role '$AppRoleName' not found on the service principal."
    Write-Host ""
    Write-Host "Available app roles:"
    $spAppRoles | ForEach-Object { Write-Host "  - $($_.value): $($_.displayName)" }
    exit 1
}

$appRoleId = $appRole.id
Write-Success "Found app role: $($appRole.displayName)"
Write-Host "  Role ID: $appRoleId"

# =============================================================================
# Check if Assignment Already Exists
# =============================================================================

Write-Step "Checking for existing role assignment..."

$existingAssignments = az rest --method GET `
    --uri "https://graph.microsoft.com/v1.0/servicePrincipals/$FoundryPrincipalId/appRoleAssignments" `
    2>$null | ConvertFrom-Json

$existingAssignment = $existingAssignments.value | Where-Object { 
    $_.appRoleId -eq $appRoleId -and $_.resourceId -eq $spObjectId 
}

if ($existingAssignment) {
    Write-Success "Role assignment already exists!"
    Write-Host "  Principal ID: $FoundryPrincipalId"
    Write-Host "  App Role: $AppRoleName"
    Write-Host "  Created: $($existingAssignment.createdDateTime)"
    exit 0
}

Write-Host "  No existing assignment found, proceeding..."

# =============================================================================
# Assign App Role
# =============================================================================

Write-Step "Assigning app role to managed identity..."

# Build JSON body and write to temp file (az rest requires @file syntax for reliable JSON)
$bodyObject = @{
    principalId = $FoundryPrincipalId
    resourceId = $spObjectId
    appRoleId = $appRoleId
}
$tempFile = [System.IO.Path]::GetTempFileName()
$bodyObject | ConvertTo-Json | Set-Content -Path $tempFile -Encoding UTF8

try {
    $result = az rest --method POST `
        --uri "https://graph.microsoft.com/v1.0/servicePrincipals/$FoundryPrincipalId/appRoleAssignments" `
        --headers "Content-Type=application/json" `
        --body "@$tempFile" | ConvertFrom-Json
    
    Remove-Item $tempFile -ErrorAction SilentlyContinue
    
    Write-Success "App role assigned successfully!"
    Write-Host ""
    Write-Host "Assignment Details:" -ForegroundColor Yellow
    Write-Host "  Assignment ID: $($result.id)"
    Write-Host "  Principal ID:  $FoundryPrincipalId"
    Write-Host "  Resource:      $($sp.displayName)"
    Write-Host "  App Role:      $AppRoleName"
    Write-Host "  Created:       $($result.createdDateTime)"
}
catch {
    Remove-Item $tempFile -ErrorAction SilentlyContinue
    Write-Error "Failed to assign app role"
    Write-Host ""
    Write-Host "Error details:" -ForegroundColor Yellow
    Write-Host $_.Exception.Message
    Write-Host ""
    Write-Host "Common issues:"
    Write-Host "  - The Foundry Principal ID may be incorrect"
    Write-Host "  - You may not have permission to assign roles"
    Write-Host "  - The managed identity may not exist"
    Write-Host ""
    Write-Host "To find the correct Principal ID:"
    Write-Host "  1. Go to Azure Portal → AI Foundry → Your Project"
    Write-Host "  2. Settings → Identity → Object (principal) ID"
    exit 1
}

# =============================================================================
# Summary
# =============================================================================

Write-Host @"

╔══════════════════════════════════════════════════════════════╗
║                    Assignment Complete!                      ║
╚══════════════════════════════════════════════════════════════╝

"@ -ForegroundColor Green

Write-Host "Next Steps:" -ForegroundColor Yellow
Write-Host "  1. In AI Foundry, add the MCP tool to your agent"
Write-Host "  2. Configure authentication:"
Write-Host "     - Authentication: Microsoft Entra → Project Managed Identity"
Write-Host "     - Audience: $AppClientId"
Write-Host ""
Write-Host "The AI Foundry agent can now authenticate with your MCP server."
