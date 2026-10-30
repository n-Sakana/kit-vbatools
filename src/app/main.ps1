param(
    [ValidateSet('Diff','Unlock','Markdown','Replace')][string]$Action = 'Diff',
    [string[]]$Paths = @(),
    [string]$BatchFile = '',
    [switch]$NoPause,
    [string]$FindText = '',
    [AllowEmptyString()][string]$ReplaceText = '',
    [ValidateSet('Word','Partial')][string]$MatchKind = 'Partial',
    [switch]$ReplaceOptionsProvided,
    [string]$ResultFile = ''
)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
[Threading.Thread]::CurrentThread.CurrentUICulture = [Globalization.CultureInfo]::GetCultureInfo('en-US')
. (Join-Path (Split-Path $PSScriptRoot -Parent) 'lib\Runtime.ps1')
$cache = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'VBATools'
[void][IO.Directory]::CreateDirectory((Join-Path $cache 'results'))
$runId = [Guid]::NewGuid().ToString('N')
$results = New-Object Collections.Generic.List[object]
$exitCode = 0
$started = Get-Date
$batch = ''
$selectionInputs = @()
try {
    if ($BatchFile) {
        $fullBatch = [IO.Path]::GetFullPath($BatchFile)
        $queue = [IO.Path]::GetFullPath((Join-Path $cache 'queue')) + [IO.Path]::DirectorySeparatorChar
        if (-not $fullBatch.StartsWith($queue, [StringComparison]::OrdinalIgnoreCase) -or [IO.Path]::GetExtension($fullBatch) -ne '.json') { throw 'Invalid selection data.' }
        $request = [IO.File]::ReadAllText($fullBatch) | ConvertFrom-Json
        if ($request.Action -notin @('Diff','Unlock','Markdown','Replace')) { throw 'Unknown action.' }
        $Action = [string]$request.Action
        $Paths = @($request.Paths)
        $batch = [string]$request.Batch
        [IO.File]::Delete($fullBatch)
    }
    $selectionInputs = @($Paths)
    $expanded = New-Object Collections.Generic.List[string]
    $seen = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($target in $selectionInputs) {
        $full = [IO.Path]::GetFullPath($target)
        if ([IO.Directory]::Exists($full)) {
            $pending = New-Object 'Collections.Generic.Stack[string]'
            $pending.Push($full)
            while ($pending.Count -gt 0) {
                $folder = $pending.Pop()
                $files = [IO.Directory]::GetFiles($folder)
                [Array]::Sort($files, [StringComparer]::OrdinalIgnoreCase)
                foreach ($file in $files) {
                    if ([IO.Path]::GetExtension($file).ToLowerInvariant() -in @('.xlsm','.xlam','.xls') -and -not [IO.Path]::GetFileName($file).StartsWith('~$') -and $seen.Add($file)) { $expanded.Add($file) }
                }
                foreach ($sub in [IO.Directory]::GetDirectories($folder)) {
                    if (([IO.File]::GetAttributes($sub) -band [IO.FileAttributes]::ReparsePoint) -eq 0) { $pending.Push($sub) }
                }
            }
        } elseif ($seen.Add($full)) { $expanded.Add($full) }
    }
    $Paths = $expanded.ToArray()
    $count = $Paths.Count
    Write-Host ''
    Write-Host ('  VBA Tools  /  {0}' -f $Action) -ForegroundColor Cyan
    Write-Host ('  ─────────────────────────────────────────────') -ForegroundColor DarkGray
    Write-Host ('  Selected: {0}  |  Output: next to each source workbook' -f $count) -ForegroundColor DarkGray
    Write-Host ''
    if ($Action -eq 'Diff' -and $count -eq 0) { throw 'VBA Diff: found 0 VBA workbook(s). Select exactly two files.' }
    if ($count -eq 0) { throw 'Select one or more files.' }
    if ($Action -eq 'Replace') {
        if (-not $ReplaceOptionsProvided) {
            if ($env:VBATOOLS_HEADLESS -eq '1' -or [Console]::IsInputRedirected) { throw 'Provide the replacement options for noninteractive use.' }
            [Console]::Write('  Find: '); $FindText = [Console]::ReadLine()
            [Console]::Write('  Replace with (empty to delete): '); $ReplaceText = [Console]::ReadLine()
            [Console]::Write('  Match  1: Whole word  2: Partial [2]: ')
            $choice = [Console]::ReadLine()
            if ($choice -eq '1') { $MatchKind = 'Word' }
            elseif ($choice -eq '2' -or $choice -eq '') { $MatchKind = 'Partial' }
            else { throw 'Enter 1 or 2 for the match type.' }
            Write-Host ''
        }
        if ([string]::IsNullOrEmpty($FindText)) { throw 'The search text cannot be empty.' }
    }
    if ($Action -in @('Diff','Replace','Markdown')) {
        $pythonInfo = Get-VbaPython
        $python = $pythonInfo.Path
        $pythonLauncher = $pythonInfo.Launcher
    }
    # Validate external file values once at the entry; all feature scripts use these paths.
    $validated = New-Object Collections.Generic.List[string]
    $supported = if ($Action -eq 'Replace') { @('.xls','.xlsm','.xlam','.xlsb') } else { @('.xls','.xla','.xlt','.xlsm','.xlam','.xlsb','.xltm','.xlsx','.xltx') }
    for ($index = 0; $index -lt $count; $index++) {
        $path = [IO.Path]::GetFullPath($Paths[$index])
        if (-not [IO.File]::Exists($path) -or [IO.Path]::GetExtension($path).ToLowerInvariant() -notin $supported) {
            $message = 'Supported Excel file not found: ' + $path
            if ($Action -eq 'Diff') { throw $message }
            Write-VbaScreen 'Failed' ([IO.Path]::GetFileName($path)) $message Red
            $results.Add([pscustomobject]@{ Input = $path; Status = 'Error'; Output = ''; Detail = $message })
            $exitCode = 1
        } else { $validated.Add($path) }
    }
    if ($validated.Count -gt 0 -and $Action -eq 'Unlock') {
        Write-VbaScreen 'Loading' 'Loading the VBA readers.'
        Add-Type -Path (Join-Path (Split-Path $PSScriptRoot -Parent) 'lib\VbaToolkitNative.cs')
    }
    $lib = Join-Path (Split-Path $PSScriptRoot -Parent) 'lib'
    if ($Action -eq 'Diff') {
        Write-VbaScreen 'Inspecting' 'Counting VBA workbooks before source extraction.'
        $inventory = & (Join-Path $lib 'Replace.ps1') -Job Inventory -Paths $validated.ToArray() -PythonPath $python -PythonLauncher:$pythonLauncher
        $withVba = New-Object Collections.Generic.List[string]
        foreach ($path in $inventory.paths) { $withVba.Add([string]$path) }
        if ($withVba.Count -ne 2) { throw ('VBA Diff: found {0} VBA workbook(s). Select exactly two files.' -f $withVba.Count) }
        $validated = $withVba
        $projects = @{}
        $index = 0
        foreach ($path in $validated) {
            $index++
            Write-VbaScreen ($index.ToString() + ' / 2') ([IO.Path]::GetFileName($path)) 'Extracting VBA.' Cyan
            $projects[$path] = & (Join-Path $lib 'Replace.ps1') -Job Modules -FilePath $path -PythonPath $python -PythonLauncher:$pythonLauncher
        }
        $report = New-VbaSiblingPath $validated[0] '_vba-diff' '.txt' @('.html','.json')
        $prefix = [IO.Path]::Combine([IO.Path]::GetDirectoryName($report), [IO.Path]::GetFileNameWithoutExtension($report))
        & (Join-Path $lib 'Diff.ps1') -FileA $validated[0] -FileB $validated[1] -OutputPrefix $prefix -Projects $projects
        $results.Add([pscustomobject]@{ Input = $validated.ToArray(); Status = 'Success'; Output = $report; Html = ($prefix + '.html'); Json = ($prefix + '.json'); Detail = 'Saved the VBA diff.' })
        Write-VbaScreen 'OK' 'VBA Diff' $report Cyan
    } else {
        $index = 0
        foreach ($path in $validated) {
            $index++
            Write-VbaScreen ($index.ToString() + ' / ' + $validated.Count) ([IO.Path]::GetFileName($path)) 'Processing.' Cyan
            try {
                if ($Action -eq 'Unlock') {
                    $output = New-VbaSiblingPath $path '_unlocked' ([IO.Path]::GetExtension($path))
                    $result = & (Join-Path $lib 'Unlock.ps1') -FilePath $path -OutputPath $output
                } elseif ($Action -eq 'Markdown') {
                    $output = New-VbaSiblingPath $path '_vba' '.md'
                    $result = & (Join-Path $lib 'Replace.ps1') -Job Markdown -FilePath $path -OutputPath $output -PythonPath $python -PythonLauncher:$pythonLauncher
                } else {
                    $output = New-VbaSiblingPath $path '_replaced' ([IO.Path]::GetExtension($path))
                    $result = & (Join-Path $lib 'Replace.ps1') -FilePath $path -OutputPath $output -FindText $FindText -ReplaceText $ReplaceText -MatchKind $MatchKind -PythonPath $python -PythonLauncher:$pythonLauncher
                }
                $item = [ordered]@{ Input = $path; Status = $result.Status; Output = $result.Output; Detail = $result.Detail }
                if ($Action -eq 'Replace') {
                    $item['ChangedModules'] = $result.ChangedModules
                    $item['Occurrences'] = $result.Occurrences
                    $item['Verified'] = $result.Verified
                }
                $results.Add([pscustomobject]$item)
                $label = if ($result.Status -eq 'NoPassword') { 'No marker' } elseif ($result.Status -eq 'NoChanges') { 'Unchanged' } else { 'OK' }
                Write-VbaScreen $label $result.Detail $result.Output Cyan
            } catch {
                $exitCode = 1
                $results.Add([pscustomobject]@{ Input = $path; Status = 'Error'; Output = ''; Detail = $_.Exception.Message })
                Write-VbaScreen 'Failed' ([IO.Path]::GetFileName($path)) $_.Exception.Message Red
            }
        }
    }
} catch {
    $exitCode = 1
    $results.Add([pscustomobject]@{ Input = @($Paths); Status = 'Error'; Output = ''; Detail = $_.Exception.Message })
    Write-VbaScreen 'Notice' $_.Exception.Message '' Red
} finally {
    $record = [pscustomobject]@{ Run = $runId; Batch = $batch; Action = $Action; Pid = $PID; Started = $started.ToString('o'); Finished = (Get-Date).ToString('o'); Targets = $selectionInputs; Inputs = @($Paths); ExitCode = $exitCode; Results = @($results.ToArray()) }
    $logPath = Join-Path $cache ('results\' + $runId + '.json')
    [IO.File]::WriteAllText($logPath, ($record | ConvertTo-Json -Depth 8), (New-Object Text.UTF8Encoding($true)))
    if ($ResultFile) { Write-VbaNewTextFile ([IO.Path]::GetFullPath($ResultFile)) ($record | ConvertTo-Json -Depth 8) }
    $success = @($results | Where-Object { $_.Status -eq 'Success' }).Count
    $skipped = @($results | Where-Object { $_.Status -eq 'NoPassword' }).Count
    $unchanged = @($results | Where-Object { $_.Status -eq 'NoChanges' }).Count
    $failed = @($results | Where-Object { $_.Status -eq 'Error' }).Count
    Write-Host ''
    Write-Host ('  ─────────────────────────────────────────────') -ForegroundColor DarkGray
    Write-Host ('  Done  {0} succeeded  /  {1} no marker  /  {2} unchanged  /  {3} failed' -f $success, $skipped, $unchanged, $failed)
    if (-not $NoPause -and $env:VBATOOLS_HEADLESS -ne '1' -and -not [Console]::IsInputRedirected) {
        Write-Host '  Press any key to close.' -ForegroundColor DarkGray
        [void][Console]::ReadKey($true)
    }
}
exit $exitCode
