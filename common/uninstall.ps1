# Adapted from KeyLoop: back up, delete only the application's HKCU verb, notify Shell.
[CmdletBinding()]
param([switch]$PlanOnly)
$ErrorActionPreference = 'Stop'
try {
    . (Join-Path $PSScriptRoot 'menu.ps1')
    $paths = @('Software\Classes\*\shell\kit-vbatools', 'Software\Classes\Directory\shell\kit-vbatools')
    if ($PlanOnly) { foreach ($path in $paths) { 'HKCU\' + $path }; exit 0 }
    foreach ($path in $paths) {
    $key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($path)
    if ($null -ne $key) {
        $key.Dispose()
        $cache = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'VBATools\menu-backups'
        [void][IO.Directory]::CreateDirectory($cache)
        $context = if ($path.Contains('\Directory\')) { 'folder' } else { 'file' }
        $backup = Join-Path $cache ((Get-Date -Format 'yyyyMMdd-HHmmssfff') + '-' + $context + '.reg')
        & reg.exe export ('HKCU\' + $path) $backup /y | Out-Null
        if ($LASTEXITCODE -ne 0) { throw 'Could not back up the menu registration.' }
        [Microsoft.Win32.Registry]::CurrentUser.DeleteSubKeyTree($path, $false)
        Send-MenuChanged
        Write-Host ('  Backup: ' + $backup)
    }
    }
    Write-Host '  VBA Tools  |  Menu removed.' -ForegroundColor Cyan
    exit 0
} catch {
    Write-Host ('  Error: ' + $_.Exception.Message) -ForegroundColor Red
    exit 1
}
