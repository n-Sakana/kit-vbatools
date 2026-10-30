param(
    [string]$FilePath = '',
    [string]$OutputPath = '',
    [AllowEmptyString()][string]$FindText = '',
    [AllowEmptyString()][string]$ReplaceText = '',
    [string]$MatchKind = 'Partial',
    [Parameter(Mandatory)][string]$PythonPath,
    [switch]$PythonLauncher,
    [ValidateSet('Replace','Markdown','Inventory','Modules')][string]$Job = 'Replace',
    [string[]]$Paths = @()
)
$ErrorActionPreference = 'Stop'
$folder = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) ('VBATools\replace\' + [Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($folder)
$process = $null
try {
    $requestPath = Join-Path $folder 'request.json'
    $request = [pscustomobject]@{ source = $FilePath; output = $OutputPath; find = $FindText; replacement = $ReplaceText; match = $MatchKind; job = $Job }
    if ($Job -eq 'Inventory') { $request = [pscustomobject]@{ job = $Job; paths = @($Paths) } }
    [IO.File]::WriteAllText($requestPath, ($request | ConvertTo-Json -Depth 3), (New-Object Text.UTF8Encoding($true)))
    $entry = Join-Path $PSScriptRoot 'replace.py'
    $arguments = '-B -X utf8 "{0}" "{1}"' -f $entry, $requestPath
    if ($PythonLauncher) { $arguments = '-3 ' + $arguments }
    $info = New-Object Diagnostics.ProcessStartInfo($PythonPath, $arguments)
    $info.UseShellExecute = $false
    $info.CreateNoWindow = $true
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardError = $true
    $info.StandardOutputEncoding = [Text.Encoding]::UTF8
    $info.StandardErrorEncoding = [Text.Encoding]::UTF8
    $process = [Diagnostics.Process]::Start($info)
    [IO.File]::WriteAllText((Join-Path $folder 'python.pid'), [string]$process.Id)
    $stdout = $process.StandardOutput.ReadToEndAsync()
    $stderr = $process.StandardError.ReadToEndAsync()
    if (-not $process.WaitForExit(120000)) { $process.Kill(); $process.WaitForExit(); throw 'The Python workbook process timed out.' }
    $text = $stdout.Result
    $errorText = $stderr.Result
    if ($process.ExitCode -ne 0) { throw ('Workbook processing failed: ' + $errorText.Trim()) }
    $record = $text | ConvertFrom-Json
    if ($Job -in @('Inventory','Modules')) { return $record }
    $status = if ($Job -eq 'Replace' -and $record.occurrences -eq 0) { 'NoChanges' } else { 'Success' }
    $detail = 'Replaced {1} occurrence(s) in {0} module(s).' -f $record.changed_modules, $record.occurrences
    if ($status -eq 'NoChanges') { $detail = '0 replacements. No copy was created.' }
    if ($Job -eq 'Markdown') { $detail = '{0} module(s) exported.' -f $record.changed_modules }
    [pscustomobject]@{ Status = $status; Output = [string]$record.output; Detail = $detail; ChangedModules = $record.changed_modules; Occurrences = $record.occurrences; Verified = $record.verified }
} finally {
    if ($null -ne $process) { $process.Dispose() }
    [IO.Directory]::Delete($folder, $true)
}
