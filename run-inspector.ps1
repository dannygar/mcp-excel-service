#!/usr/bin/env pwsh
# Launch MCP Inspector (latest version via npx)

param(
    [string]$Url = "http://localhost:3000/mcp"
)

Write-Host "🔍 Starting MCP Inspector (latest version)" -ForegroundColor Cyan
Write-Host "   Server URL: $Url" -ForegroundColor Gray
Write-Host ""
Write-Host "⚠️  Note: Select 'Streamable HTTP' transport and enter the URL above" -ForegroundColor Yellow
Write-Host ""

# Run the latest MCP Inspector via npx
npx @modelcontextprotocol/inspector@latest
