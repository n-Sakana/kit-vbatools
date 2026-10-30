[CmdletBinding()]
param([switch]$PlanOnly)
$ErrorActionPreference = 'Stop'
try {
    . (Join-Path $PSScriptRoot 'menu.ps1')
    if ($PlanOnly) { @(Get-MenuPlan) | ConvertTo-Json -Depth 6; exit 0 }
    Install-StandaloneMenu
    Write-Host ''
    Write-Host '  VBA Tools  |  Menu installed.' -ForegroundColor Cyan
    Write-Host '  Select files or folders > Right-click > Show more options > VBA Tools'
    Write-Host '  Run install.bat again after moving this folder.'
    exit 0
} catch {
    Write-Host ('  Error: ' + $_.Exception.Message) -ForegroundColor Red
    exit 1
}
