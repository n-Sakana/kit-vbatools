# Explorer starts one ordinary verb process per selected file. Only the leader opens a TUI.
param(
    [Parameter(Mandatory)][ValidateSet('Diff','Unlock','Markdown','Replace')][string]$Action,
    [Parameter(Mandatory)][string]$Target
)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$cache = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'VBATools'
$queue = Join-Path $cache ('queue\' + $Action)
[void][IO.Directory]::CreateDirectory($queue)
$targetPath = [IO.Path]::GetFullPath($Target)
$sid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$mutex = New-Object Threading.Mutex($false, ('Local\VBATools-' + $sid + '-' + $Action))
$statePath = Join-Path $queue 'state.json'
$leader = $false
$batch = ''
$lockHeld = $false
try {
    try { $lockHeld = $mutex.WaitOne(15000) } catch [Threading.AbandonedMutexException] { $lockHeld = $true }
    if (-not $lockHeld) { throw 'Selection collection timed out. Please try again.' }
    $state = $null
    if ([IO.File]::Exists($statePath)) {
        $state = [IO.File]::ReadAllText($statePath) | ConvertFrom-Json
        # A dead leader never steals the next selection. Only its own recorded PID is inspected.
        $ownerAlive = $true
        try { $owner = [Diagnostics.Process]::GetProcessById([int]$state.Owner); $owner.Dispose() }
        catch [ArgumentException] { $ownerAlive = $false }
        if (-not $ownerAlive -or ([DateTime]::UtcNow.Ticks - [long]$state.Started) -gt 300000000) { $state = $null }
    }
    if ($null -eq $state) {
        $leader = $true
        $batch = [Guid]::NewGuid().ToString('N')
        $state = [pscustomobject]@{ Batch = $batch; Owner = $PID; Started = [DateTime]::UtcNow.Ticks; Last = [DateTime]::UtcNow.Ticks; Paths = @() }
    }
    if (@($state.Paths) -notcontains $targetPath) { $state.Paths = @($state.Paths) + $targetPath }
    $state.Last = [DateTime]::UtcNow.Ticks
    $temporary = $statePath + '.' + $PID + '.tmp'
    [IO.File]::WriteAllText($temporary, ($state | ConvertTo-Json -Depth 4), (New-Object Text.UTF8Encoding($true)))
    if ([IO.File]::Exists($statePath)) { [IO.File]::Delete($statePath) }
    [IO.File]::Move($temporary, $statePath)
    $mutex.ReleaseMutex(); $lockHeld = $false
    if (-not $leader) { exit 0 }

    while ($true) {
        [Threading.Thread]::Sleep(200)
        try { $lockHeld = $mutex.WaitOne(15000) } catch [Threading.AbandonedMutexException] { $lockHeld = $true }
        if (-not $lockHeld) { throw 'Selection collection timed out.' }
        $state = [IO.File]::ReadAllText($statePath) | ConvertFrom-Json
        $now = [DateTime]::UtcNow.Ticks
        # Minimum 2.5 s, then 1.5 s with no arrivals; cap one burst at 15 s.
        $ready = (($now - [long]$state.Started -ge 25000000) -and ($now - [long]$state.Last -ge 15000000)) -or ($now - [long]$state.Started -ge 150000000)
        if ($ready) {
            $batchPath = Join-Path $queue ($batch + '.json')
            $request = [pscustomobject]@{ Action = $Action; Paths = @($state.Paths); CollectorPid = $PID; Batch = $batch }
            [IO.File]::WriteAllText($batchPath, ($request | ConvertTo-Json -Depth 4), (New-Object Text.UTF8Encoding($true)))
            [IO.File]::Delete($statePath)
        }
        $mutex.ReleaseMutex(); $lockHeld = $false
        if ($ready) { break }
    }
    $main = Join-Path $PSScriptRoot 'main.ps1'
    $powershell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $arguments = '-NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File "{0}" -BatchFile "{1}"' -f $main, $batchPath
    if ($env:VBATOOLS_HEADLESS -eq '1') {
        $info = New-Object Diagnostics.ProcessStartInfo($powershell, $arguments)
        $info.UseShellExecute = $false
        $info.CreateNoWindow = $true
        $child = [Diagnostics.Process]::Start($info)
    } else {
        # Like KeyLoop, request the classic console even when Terminal is the default.
        $info = New-Object Diagnostics.ProcessStartInfo((Join-Path $env:SystemRoot 'System32\conhost.exe'), ('"' + $powershell + '" ' + $arguments))
        $info.UseShellExecute = $false
        $child = [Diagnostics.Process]::Start($info)
    }
    [IO.File]::WriteAllText(($batchPath + '.pid'), [string]$child.Id)
    $child.Dispose()
} catch {
    # No dialogs or hidden failures. Diagnostics stay in this application's cache.
    [IO.File]::AppendAllText((Join-Path $cache 'collector-errors.log'), ((Get-Date -Format o) + ' ' + $_.Exception.Message + "`r`n"))
    exit 1
} finally {
    if ($lockHeld) { $mutex.ReleaseMutex() }
    $mutex.Dispose()
}
