#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Start the TastyTrade MCP Server locally for development

.DESCRIPTION
    Runs the FastMCP server locally using the mcp-server/server.py.
    This is the local development equivalent of the Azure Container Apps deployment.

.PARAMETER Port
    Port number to run the server on (default: 3000)

.PARAMETER Environment
    Environment to load credentials from: dev or prod (default: prod)
    Loads from config\.env.{Environment}

.PARAMETER Debug
    Enable debug mode with more verbose logging

.EXAMPLE
    .\start-mcp-server-local.ps1
    
.EXAMPLE
    .\start-mcp-server-local.ps1 -Port 8080 -Environment dev

.EXAMPLE
    .\start-mcp-server-local.ps1 -Debug
#>

param(
    [int]$Port = 3000,
    [ValidateSet("local", "dev", "prod")]
    [string]$Environment = "local",
    [switch]$Debug
)

$ErrorActionPreference = "Stop"

# Get the script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$serverDir = Join-Path $scriptDir "mcp-server"
$serverScript = Join-Path $serverDir "server.py"

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Excel MCP Server - Local Development" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# Check if server.py exists
if (-not (Test-Path $serverScript)) {
    Write-Host "Error: server.py not found at $serverScript" -ForegroundColor Red
    exit 1
}

# Check if uv is available
$uvPath = Get-Command uv -ErrorAction SilentlyContinue
if (-not $uvPath) {
    Write-Host "Error: uv is not installed. Install it with: pip install uv" -ForegroundColor Red
    exit 1
}

# Set environment variables
$env:PORT = $Port
$env:MCP_ENV = $Environment

# Clear VIRTUAL_ENV to avoid conflicts with mcp-server's own .venv
$env:VIRTUAL_ENV = $null

# If Debug mode, set log level
if ($Debug) {
    $env:LOG_LEVEL = "DEBUG"
}

Write-Host "Configuration:" -ForegroundColor Yellow
Write-Host "  Port:           $Port"
Write-Host "  Environment:    $Environment (loading from config\.env.$Environment)"
Write-Host "  Debug:          $($Debug.IsPresent)"
Write-Host ""

# Sync dependencies
Write-Host "Syncing dependencies..." -ForegroundColor Yellow
Push-Location $serverDir
try {
    uv sync --quiet
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Error: Failed to sync dependencies" -ForegroundColor Red
        exit 1
    }
    Write-Host "Dependencies synced" -ForegroundColor Green
}
finally {
    Pop-Location
}

Write-Host ""
Write-Host "Starting MCP Server..." -ForegroundColor Green
Write-Host "  MCP Endpoint:    http://localhost:$Port/mcp" -ForegroundColor Cyan
Write-Host "  Health Endpoint: http://localhost:$Port/health" -ForegroundColor Cyan
Write-Host ""
Write-Host "Press Ctrl+C to stop the server" -ForegroundColor Gray
Write-Host ""

# Run the server
Push-Location $serverDir
try {
    uv run python server.py
}
finally {
    Pop-Location
}
