#!/usr/bin/env pwsh
# Launch MCP Inspector with proper Windows environment variable handling

param(
    [string]$Url = "http://localhost:3000/mcp",
    [int]$ClientPort = 5173,
    [int]$ServerPort = 3001
)

Write-Host "Starting MCP Inspector for: $Url" -ForegroundColor Cyan

$nodeModules = Join-Path $PSScriptRoot "node_modules"
$inspectorServerPath = Join-Path $nodeModules "@modelcontextprotocol\inspector\server\build\index.js"
$inspectorClientPath = Join-Path $nodeModules "@modelcontextprotocol\inspector\client\bin\cli.js"

if (-not (Test-Path $inspectorServerPath)) {
    Write-Host "Error: MCP Inspector not found. Run 'yarn install' first." -ForegroundColor Red
    exit 1
}

Write-Host "🔍 MCP Inspector starting at http://localhost:$ClientPort 🚀" -ForegroundColor Green
Write-Host ""
Write-Host "⚠️  Note: When stopping the inspector, restart the MCP server to avoid connection errors" -ForegroundColor Yellow
Write-Host ""

# Start server in background job
$serverJob = Start-Job -ScriptBlock {
    param($ServerPath, $ServerPort, $Url)
    $env:PORT = $ServerPort
    & node $ServerPath --env $Url
} -ArgumentList $inspectorServerPath, $ServerPort, $Url

# Start client in background job with URL query parameters
$clientJob = Start-Job -ScriptBlock {
    param($ClientPath, $ClientPort, $Url)
    $env:PORT = $ClientPort
    # Launch with transport type and URL as query parameters
    & node $ClientPath --url "http://localhost:$ClientPort/?transport=sse&url=$([uri]::EscapeDataString($Url))"
} -ArgumentList $inspectorClientPath, $ClientPort, $Url

Write-Host "Server started (Job ID: $($serverJob.Id))" -ForegroundColor Gray
Write-Host "Client started (Job ID: $($clientJob.Id))" -ForegroundColor Gray
Write-Host ""
Write-Host "Press Ctrl+C to stop the inspector" -ForegroundColor Yellow
Write-Host ""

try {
    # Monitor jobs and stream output
    while ($serverJob.State -eq 'Running' -or $clientJob.State -eq 'Running') {
        # Receive output from jobs
        Receive-Job -Job $serverJob -ErrorAction SilentlyContinue | ForEach-Object {
            Write-Host "[server] $_" -ForegroundColor Cyan
        }
        Receive-Job -Job $clientJob -ErrorAction SilentlyContinue | ForEach-Object {
            Write-Host "[client] $_" -ForegroundColor Green
        }
        Start-Sleep -Milliseconds 100
    }
} finally {
    # Cleanup jobs on exit
    Write-Host ""
    Write-Host "Stopping inspector gracefully..." -ForegroundColor Yellow
    
    # Give jobs time to finish current operations
    Stop-Job -Job $serverJob, $clientJob -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 1
    
    Remove-Job -Job $serverJob, $clientJob -Force -ErrorAction SilentlyContinue
    Write-Host "Inspector stopped." -ForegroundColor Gray
    Write-Host ""
    Write-Host "💡 Tip: Restart the MCP server if you see connection errors" -ForegroundColor Cyan
}
