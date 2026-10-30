# Adapted from vba-devkit/lib/Unlock.ps1. Same DPB= -> DPx= patch, on a copy.
# ZIP workbooks keep their format; no Excel SaveAs conversion or application UI.
param(
    [Parameter(Mandatory)][string]$FilePath,
    [Parameter(Mandatory)][string]$OutputPath
)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'VBAToolkit.psm1') -Force -DisableNameChecking

function Find-DPB([byte[]]$Data) {
    $pattern = [Text.Encoding]::ASCII.GetBytes('DPB=')
    for ($i = 0; $i -le $Data.Length - $pattern.Length; $i++) {
        $match = $true
        for ($j = 0; $j -lt $pattern.Length; $j++) {
            if ($Data[$i + $j] -ne $pattern[$j]) { $match = $false; break }
        }
        if ($match) { return $i }
    }
    return -1
}

$container = Get-VbaProjectBytes $FilePath
if (-not $container.Bytes) { return [pscustomobject]@{ Status = 'NoPassword'; Output = ''; Detail = 'No DPB= marker. Skipped without creating a copy.' } }
$ole = Read-Ole2 $container.Bytes
$entry = $ole.Entries | Where-Object { $_.Name -eq 'PROJECT' -and $_.ObjType -eq 2 } | Select-Object -First 1
if ($null -eq $entry) { return [pscustomobject]@{ Status = 'NoPassword'; Output = ''; Detail = 'No VBA PROJECT / DPB= marker.' } }
$data = Read-Ole2Stream $ole $entry
$position = Find-DPB $data
if ($position -lt 0) { return [pscustomobject]@{ Status = 'NoPassword'; Output = ''; Detail = 'No DPB= marker. Skipped without creating a copy.' } }
$data[$position + 2] = 0x78
Write-Ole2Stream $container.Bytes $ole $entry $data

# Reserve a new name atomically. Only the copy is ever written or cleaned up.
$created = $false
try {
    [IO.File]::Copy($FilePath, $OutputPath, $false)
    $created = $true
    Save-VbaProjectBytes $OutputPath $container.Bytes $container.IsZip
    return [pscustomobject]@{ Status = 'Success'; Output = $OutputPath; Detail = 'Changed DPB= to DPx= in the copy.' }
} catch {
    if ($created -and [IO.File]::Exists($OutputPath)) { [IO.File]::Delete($OutputPath) }
    throw
}
