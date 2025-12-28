<#
.SYNOPSIS
    Assign RBAC role to Container App managed identity for AI Foundry access

.DESCRIPTION
    This script assigns the "Cognitive Services User" role to a Container App's
    managed identity, allowing it to access the AI Foundry resource.

.PARAMETER ContainerAppName
    Name of the Container App (required)

.PARAMETER ResourceGroup
    Resource group containing the Container App (required)

.PARAMETER AIFoundryResourceGroup
    Resource group containing the AI Foundry resource (required)

.PARAMETER AIFoundryResourceName
    Name of the AI Foundry resource (required)

.PARAMETER RoleDefinitionId
    Role definition ID to assign. Default: Cognitive Services User (a97b65f3-24c7-4388-baec-2e87135dc908)

.EXAMPLE
    .\scripts\assign-ai-foundry-rbac.ps1 -ContainerAppName "ca-webapp" -ResourceGroup "rg-webapp" -AIFoundryResourceGroup "rg-ai" -AIFoundryResourceName "ai-foundry"

.EXAMPLE
    # With custom role
    .\scripts\assign-ai-foundry-rbac.ps1 -ContainerAppName "ca-webapp" -ResourceGroup "rg-webapp" -AIFoundryResourceGroup "rg-ai" -AIFoundryResourceName "ai-foundry" -RoleDefinitionId "custom-role-id"
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$ContainerAppName,
    
    [Parameter(Mandatory=$true)]
    [string]$ResourceGroup,
    
    [Parameter(Mandatory=$true)]
    [string]$AIFoundryResourceGroup,
    
    [Parameter(Mandatory=$true)]
    [string]$AIFoundryResourceName,
    
    [string]$RoleDefinitionId = "a97b65f3-24c7-4388-baec-2e87135dc908"  # Cognitive Services User
)

$ErrorActionPreference = "Stop"

# Helper functions
function Write-Step { param($Message) Write-Host "`n▶ $Message" -ForegroundColor Cyan }
function Write-Success { param($Message) Write-Host "✓ $Message" -ForegroundColor Green }
function Write-Warning { param($Message) Write-Host "⚠ $Message" -ForegroundColor Yellow }
function Write-Error { param($Message) Write-Host "✗ $Message" -ForegroundColor Red }

Write-Host @"

╔══════════════════════════════════════════════════════════════╗
║           AI Foundry RBAC Role Assignment                     ║
╚══════════════════════════════════════════════════════════════╝

"@ -ForegroundColor Magenta

# =============================================================================
# Step 1: Get Container App Managed Identity
# =============================================================================

Write-Step "Getting Container App managed identity..."

$containerApp = az containerapp show `
    --name $ContainerAppName `
    --resource-group $ResourceGroup `
    2>$null | ConvertFrom-Json

if (-not $containerApp) {
    Write-Error "Container App '$ContainerAppName' not found in resource group '$ResourceGroup'"
    exit 1
}

$principalId = $containerApp.identity.principalId

if (-not $principalId) {
    Write-Error "Container App does not have a system-assigned managed identity"
    Write-Host "Enable managed identity with:" -ForegroundColor Gray
    Write-Host "  az containerapp identity assign --name $ContainerAppName --resource-group $ResourceGroup --system-assigned" -ForegroundColor Gray
    exit 1
}

Write-Success "Found managed identity: $principalId"

# =============================================================================
# Step 2: Get AI Foundry Resource
# =============================================================================

Write-Step "Verifying AI Foundry resource..."

$aiFoundry = az cognitiveservices account show `
    --name $AIFoundryResourceName `
    --resource-group $AIFoundryResourceGroup `
    2>$null | ConvertFrom-Json

if (-not $aiFoundry) {
    Write-Error "AI Foundry resource '$AIFoundryResourceName' not found in resource group '$AIFoundryResourceGroup'"
    exit 1
}

Write-Success "Found AI Foundry resource: $($aiFoundry.name)"
Write-Host "  Location: $($aiFoundry.location)" -ForegroundColor Gray
Write-Host "  Kind: $($aiFoundry.kind)" -ForegroundColor Gray

# =============================================================================
# Step 3: Check Existing Role Assignment
# =============================================================================

Write-Step "Checking existing role assignments..."

$existingAssignment = az role assignment list `
    --assignee $principalId `
    --role $RoleDefinitionId `
    --scope $aiFoundry.id `
    2>$null | ConvertFrom-Json

if ($existingAssignment -and $existingAssignment.Count -gt 0) {
    Write-Success "Role assignment already exists"
    Write-Host "  Principal: $principalId" -ForegroundColor Gray
    Write-Host "  Role: Cognitive Services User" -ForegroundColor Gray
    Write-Host "  Scope: $($aiFoundry.name)" -ForegroundColor Gray
    exit 0
}

# =============================================================================
# Step 4: Create Role Assignment
# =============================================================================

Write-Step "Creating role assignment..."

Write-Host "  Principal: $principalId" -ForegroundColor Gray
Write-Host "  Role: Cognitive Services User ($RoleDefinitionId)" -ForegroundColor Gray
Write-Host "  Scope: $($aiFoundry.id)" -ForegroundColor Gray
Write-Host ""

try {
    az role assignment create `
        --assignee-object-id $principalId `
        --assignee-principal-type ServicePrincipal `
        --role $RoleDefinitionId `
        --scope $aiFoundry.id `
        --output none
    
    if ($LASTEXITCODE -ne 0) {
        throw "Role assignment failed"
    }
    
    Write-Success "Role assignment created successfully!"
}
catch {
    Write-Error "Failed to create role assignment: $_"
    Write-Host ""
    Write-Host "This may be due to insufficient permissions. You need one of:" -ForegroundColor Yellow
    Write-Host "  - Owner role on the AI Foundry resource" -ForegroundColor Gray
    Write-Host "  - User Access Administrator role on the AI Foundry resource" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Ask your subscription admin to run this command:" -ForegroundColor Yellow
    Write-Host "  az role assignment create --assignee-object-id $principalId --assignee-principal-type ServicePrincipal --role '$RoleDefinitionId' --scope $($aiFoundry.id)" -ForegroundColor White
    exit 1
}

# =============================================================================
# Step 5: Verify Assignment
# =============================================================================

Write-Step "Verifying role assignment..."

Start-Sleep -Seconds 5  # Wait for propagation

$verifyAssignment = az role assignment list `
    --assignee $principalId `
    --role $RoleDefinitionId `
    --scope $aiFoundry.id `
    2>$null | ConvertFrom-Json

if ($verifyAssignment -and $verifyAssignment.Count -gt 0) {
    Write-Success "Role assignment verified!"
} else {
    Write-Warning "Role assignment may still be propagating. Check again in a few moments."
}

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "RBAC Configuration Complete" -ForegroundColor Green
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host ""
Write-Host "The Container App's managed identity now has access to:" -ForegroundColor White
Write-Host "  AI Foundry Resource: $AIFoundryResourceName" -ForegroundColor Gray
Write-Host "  Role: Cognitive Services User" -ForegroundColor Gray
Write-Host ""
