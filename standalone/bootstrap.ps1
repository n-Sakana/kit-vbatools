# Source for the generated single-file BAT tools. No sibling files are needed at runtime.
$ErrorActionPreference = 'Stop'
[Threading.Thread]::CurrentThread.CurrentUICulture = [Globalization.CultureInfo]::GetCultureInfo('en-US')
[Console]::OutputEncoding = New-Object Text.UTF8Encoding($false)
$action = $env:VBATOOLS_ACTION
$paths = New-Object Collections.Generic.List[string]
for ($index = 1; $index -le [int]$env:VBATOOLS_ARG_COUNT; $index++) {
    $paths.Add([Environment]::GetEnvironmentVariable('VBATOOLS_ARG_' + $index))
}
if ($paths.Count -eq 0) {
    if ($env:VBATOOLS_HEADLESS -eq '1') { [Console]::Error.WriteLine('Provide one or more files or folders.'); exit 2 }
    Add-Type -AssemblyName System.Windows.Forms
    $picker = New-Object Windows.Forms.OpenFileDialog
    try {
        $picker.Title = if ($action -eq 'Diff') { 'VBA Diff: select exactly two workbooks' } else { 'VBA Tools: select workbooks' }
        $picker.Filter = 'Excel VBA workbooks (*.xlsm;*.xlam;*.xls)|*.xlsm;*.xlam;*.xls|All files (*.*)|*.*'
        $picker.Multiselect = $true
        if ($picker.ShowDialog() -ne [Windows.Forms.DialogResult]::OK) { exit 0 }
        foreach ($path in $picker.FileNames) { $paths.Add($path) }
    } finally { $picker.Dispose() }
}
$selfText = [IO.File]::ReadAllText($env:VBATOOLS_SELF)
$payload = [Convert]::FromBase64String([regex]::Match($selfText, '(?m)^# VBATOOLS_PAYLOAD=([^\r\n]+)').Groups[1].Value)
$hasher = [Security.Cryptography.SHA256]::Create()
try { $hash = [BitConverter]::ToString($hasher.ComputeHash($payload)).Replace('-', '') } finally { $hasher.Dispose() }
$cache = [IO.Path]::Combine([Environment]::GetFolderPath('LocalApplicationData'), 'VBATools', 'standalone', $action + '-' + $hash)
$mutex = New-Object Threading.Mutex($false, ('Local\VBAToolsStandalone-' + $hash))
$held = $false
try {
    try { $held = $mutex.WaitOne(30000) } catch [Threading.AbandonedMutexException] { $held = $true }
    if (-not $held) { throw 'Timed out while preparing the standalone tool.' }
    if (-not [IO.File]::Exists([IO.Path]::Combine($cache, 'complete'))) {
        [void][IO.Directory]::CreateDirectory($cache)
        Add-Type -AssemblyName System.IO.Compression
        $memory = New-Object IO.MemoryStream(,$payload)
        $archive = New-Object IO.Compression.ZipArchive($memory, [IO.Compression.ZipArchiveMode]::Read)
        try {
            foreach ($entry in $archive.Entries) {
                $target = [IO.Path]::Combine($cache, $entry.FullName.Replace('/', [IO.Path]::DirectorySeparatorChar))
                if (-not [IO.Path]::GetFullPath($target).StartsWith($cache + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) { throw 'Invalid embedded file path.' }
                [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($target))
                $inputStream = $entry.Open()
                $outputStream = [IO.File]::Create($target)
                try { $inputStream.CopyTo($outputStream) } finally { $outputStream.Dispose(); $inputStream.Dispose() }
            }
        } finally { $archive.Dispose(); $memory.Dispose() }
        [IO.File]::WriteAllText([IO.Path]::Combine($cache, 'complete'), $hash)
    }
} finally {
    if ($held) { $mutex.ReleaseMutex() }
    $mutex.Dispose()
}
$parameters = @{ Action = $action; Paths = $paths.ToArray() }
if ($action -eq 'Replace' -and $env:VBATOOLS_REPLACE_OPTIONS) {
    $options = [IO.File]::ReadAllText($env:VBATOOLS_REPLACE_OPTIONS) | ConvertFrom-Json
    $parameters['ReplaceOptionsProvided'] = $true
    $parameters['FindText'] = [string]$options.Find
    $parameters['ReplaceText'] = [string]$options.Replacement
    $parameters['MatchKind'] = [string]$options.Match
}
if ($env:VBATOOLS_TEST_RESULT_FILE) { $parameters['ResultFile'] = $env:VBATOOLS_TEST_RESULT_FILE }
& ([IO.Path]::Combine($cache, 'src', 'app', 'main.ps1')) @parameters
exit $LASTEXITCODE
