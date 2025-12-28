<#
.SYNOPSIS
    Discover Azure AI Foundry resources, projects, and agents

.DESCRIPTION
    This script discovers AI Foundry resources in the current Azure subscription,
    lists projects within them, and discovers agents. It can be used standalone
    or called by deployment scripts.

.PARAMETER NonInteractive
    If set, automatically selects the first available option instead of prompting

.PARAMETER OutputJson
    If set, outputs results as JSON for programmatic consumption

.EXAMPLE
    .\scripts\discover-ai-foundry.ps1

.EXAMPLE
    .\scripts\discover-ai-foundry.ps1 -NonInteractive

.EXAMPLE
    $result = .\scripts\discover-ai-foundry.ps1 -OutputJson | ConvertFrom-Json
#>

param(
    [switch]$NonInteractive,
    [switch]$OutputJson
)

$ErrorActionPreference = "Stop"

# Helper functions for formatted output (suppress in JSON mode)
function Write-Step { 
    param($Message) 
    if (-not $OutputJson) { Write-Host "`n▶ $Message" -ForegroundColor Cyan }
}
function Write-Success { 
    param($Message) 
    if (-not $OutputJson) { Write-Host "✓ $Message" -ForegroundColor Green }
}
function Write-Warning { 
    param($Message) 
    if (-not $OutputJson) { Write-Host "⚠ $Message" -ForegroundColor Yellow }
}
function Write-Info { 
    param($Message) 
    if (-not $OutputJson) { Write-Host "  $Message" -ForegroundColor Gray }
}

# Result object
$result = @{
    success = $false
    resourceGroup = ""
    resourceName = ""
    resourceId = ""
    projectName = ""
    endpoint = ""
    agentName = ""
    agents = @()
}

# Check Azure CLI login
$account = az account show 2>$null | ConvertFrom-Json
if (-not $account) {
    if ($OutputJson) {
        $result.error = "Not logged into Azure. Run 'az login' first."
        $result | ConvertTo-Json -Depth 10
        exit 1
    }
    Write-Error "Not logged into Azure. Run 'az login' first."
    exit 1
}

if (-not $OutputJson) {
    Write-Host @"

╔══════════════════════════════════════════════════════════════╗
║          Azure AI Foundry Resource Discovery                  ║
╚══════════════════════════════════════════════════════════════╝

"@ -ForegroundColor Magenta
    Write-Host "Subscription: $($account.name)" -ForegroundColor Gray
    Write-Host "User:         $($account.user.name)" -ForegroundColor Gray
}

# =============================================================================
# Step 1: Discover AI Foundry Resources
# =============================================================================

Write-Step "Discovering AI Foundry resources..."

# AI Foundry resources are Cognitive Services accounts with kind='AIServices'
$aiFoundryResources = az cognitiveservices account list --query "[?kind=='AIServices']" 2>$null | ConvertFrom-Json

if (-not $aiFoundryResources -or $aiFoundryResources.Count -eq 0) {
    if ($OutputJson) {
        $result.error = "No AI Foundry resources found in subscription"
        $result | ConvertTo-Json -Depth 10
        exit 1
    }
    Write-Error @"
No Azure AI Foundry resources found in subscription.

To use this application, you need an Azure AI Foundry resource with a project and agent.

Option 1 - Create a new AI Foundry resource:
  1. Visit https://ai.azure.com
  2. Create a new AI Foundry resource and project
  3. Create an agent in the project
  4. Run this script again

For more information, visit: https://learn.microsoft.com/azure/ai-foundry
"@
    exit 1
}

# Select resource
$selectedResource = $null

if ($aiFoundryResources.Count -eq 1) {
    $selectedResource = $aiFoundryResources[0]
    Write-Success "Found 1 AI Foundry resource: $($selectedResource.name)"
} else {
    if (-not $OutputJson) {
        Write-Host "Found $($aiFoundryResources.Count) AI Foundry resources:" -ForegroundColor Cyan
        Write-Host ""
        for ($i = 0; $i -lt $aiFoundryResources.Count; $i++) {
            $res = $aiFoundryResources[$i]
            Write-Host "  [$($i+1)] $($res.name)" -ForegroundColor White
            Write-Host "      Resource Group: $($res.resourceGroup)" -ForegroundColor Gray
            Write-Host "      Location: $($res.location)" -ForegroundColor Gray
        }
        Write-Host ""
    }
    
    if ($NonInteractive) {
        $selectedResource = $aiFoundryResources[0]
        Write-Warning "Non-interactive mode: using first resource '$($selectedResource.name)'"
    } else {
        if ($OutputJson) {
            # In JSON mode without interactive, just pick first
            $selectedResource = $aiFoundryResources[0]
        } else {
            Write-Host "Please select which resource to use (1-$($aiFoundryResources.Count)):" -ForegroundColor Yellow -NoNewline
            $selection = Read-Host " "
            
            $selectionNum = 0
            if (-not [int]::TryParse($selection, [ref]$selectionNum) -or $selectionNum -lt 1 -or $selectionNum -gt $aiFoundryResources.Count) {
                Write-Error "Invalid selection. Please enter a number between 1 and $($aiFoundryResources.Count)"
                exit 1
            }
            
            $selectedResource = $aiFoundryResources[$selectionNum - 1]
            Write-Success "Selected: $($selectedResource.name)"
        }
    }
}

$result.resourceGroup = $selectedResource.resourceGroup
$result.resourceName = $selectedResource.name
$result.resourceId = $selectedResource.id

# =============================================================================
# Step 2: Discover Projects
# =============================================================================

Write-Step "Discovering projects in $($selectedResource.name)..."

$resourceId = $selectedResource.id
$projectsUrl = "https://management.azure.com$resourceId/projects?api-version=2025-04-01-preview"
$projects = az rest --method get --url $projectsUrl --query "value" 2>$null | ConvertFrom-Json

if (-not $projects -or $projects.Count -eq 0) {
    if ($OutputJson) {
        $result.error = "No projects found in AI Foundry resource '$($selectedResource.name)'"
        $result | ConvertTo-Json -Depth 10
        exit 1
    }
    Write-Error @"
No projects found in AI Foundry resource '$($selectedResource.name)'.

To use this application, you need to create a project and agent:
  1. Visit https://ai.azure.com
  2. Open resource: $($selectedResource.name)
  3. Create a new project
  4. Create an agent in the project
  5. Run this script again

For more information, visit: https://learn.microsoft.com/azure/ai-foundry/quickstarts/get-started-code
"@
    exit 1
}

$selectedProject = $projects[0]
$projectName = $selectedProject.name.Split('/')[-1]

if ($projects.Count -eq 1) {
    Write-Success "Found 1 project: $projectName"
} else {
    Write-Warning "Found $($projects.Count) projects, using first: $projectName"
}

$result.projectName = $projectName
$result.endpoint = "https://$($selectedResource.name).services.ai.azure.com/api/projects/$projectName"

# =============================================================================
# Step 3: Discover Agents
# =============================================================================

Write-Step "Discovering agents in project '$projectName'..."

# Get access token for AI Foundry API
$accessToken = az account get-access-token --resource "https://cognitiveservices.azure.com" --query accessToken -o tsv 2>$null

if (-not $accessToken) {
    Write-Warning "Could not get access token for agent discovery"
} else {
    try {
        $agentsUrl = "$($result.endpoint)/agents?api-version=v2025-05-01"
        $headers = @{
            "Authorization" = "Bearer $accessToken"
            "Content-Type" = "application/json"
        }
        
        $response = Invoke-RestMethod -Uri $agentsUrl -Method Get -Headers $headers -TimeoutSec 30
        
        if ($response.data -and $response.data.Count -gt 0) {
            $agents = $response.data
            $result.agents = @($agents | ForEach-Object { @{ name = $_.name; id = $_.id } })
            
            if ($agents.Count -eq 1) {
                $result.agentName = $agents[0].name
                Write-Success "Found 1 agent: $($agents[0].name)"
            } else {
                if (-not $OutputJson) {
                    Write-Host "Found $($agents.Count) agents:" -ForegroundColor Cyan
                    for ($i = 0; $i -lt [Math]::Min($agents.Count, 10); $i++) {
                        Write-Host "  [$($i+1)] $($agents[$i].name)" -ForegroundColor Gray
                    }
                    if ($agents.Count -gt 10) {
                        Write-Host "  ... and $($agents.Count - 10) more" -ForegroundColor Gray
                    }
                    Write-Host ""
                }
                
                if ($NonInteractive) {
                    $result.agentName = $agents[0].name
                    Write-Warning "Non-interactive mode: using first agent '$($agents[0].name)'"
                } elseif (-not $OutputJson) {
                    Write-Host "Please select an agent (1-$($agents.Count)) or press Enter for first:" -ForegroundColor Yellow -NoNewline
                    $agentSelection = Read-Host " "
                    
                    if ([string]::IsNullOrEmpty($agentSelection)) {
                        $result.agentName = $agents[0].name
                        Write-Success "Using first agent: $($agents[0].name)"
                    } else {
                        $agentNum = 0
                        if ([int]::TryParse($agentSelection, [ref]$agentNum) -and $agentNum -ge 1 -and $agentNum -le $agents.Count) {
                            $result.agentName = $agents[$agentNum - 1].name
                            Write-Success "Selected agent: $($result.agentName)"
                        } else {
                            $result.agentName = $agents[0].name
                            Write-Warning "Invalid selection, using first agent: $($agents[0].name)"
                        }
                    }
                } else {
                    $result.agentName = $agents[0].name
                }
            }
        } else {
            Write-Warning "No agents found in project"
            if (-not $OutputJson) {
                Write-Host @"

To create an agent:
  1. Visit https://ai.azure.com
  2. Open project: $projectName
  3. Go to 'Agents' and create a new agent
  4. Run this script again

"@ -ForegroundColor Gray
            }
        }
    }
    catch {
        Write-Warning "Could not list agents: $($_.Exception.Message)"
    }
}

$result.success = $true

# =============================================================================
# Output Results
# =============================================================================

if ($OutputJson) {
    $result | ConvertTo-Json -Depth 10
} else {
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Green
    Write-Host "Discovery Complete" -ForegroundColor Green
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Green
    Write-Host ""
    Write-Host "Resource Group:  $($result.resourceGroup)" -ForegroundColor White
    Write-Host "Resource Name:   $($result.resourceName)" -ForegroundColor White
    Write-Host "Project:         $($result.projectName)" -ForegroundColor White
    Write-Host "Endpoint:        $($result.endpoint)" -ForegroundColor White
    if ($result.agentName) {
        Write-Host "Agent:           $($result.agentName)" -ForegroundColor White
    }
    Write-Host ""
    
    # Return the result object for script chaining
    return $result
}
