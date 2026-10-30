# Binary fixture support. App source extraction uses the shared Python reader.
function Import-VbaReaders {
    if ('Macro2Chat.BookIO' -as [type]) { return }
    $vendor = Join-Path (Split-Path $PSScriptRoot -Parent) 'vendor\Macro2Chat'
    $sources = @('00_OutputFiles.cs','05_Ole2.cs','06_VbaCompression.cs','07_VbaProject.cs','08_BookIO.cs') | ForEach-Object { Join-Path $vendor $_ }
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    # In memory only: no custom EXE or cached DLL is installed.
    Add-Type -Path $sources -ReferencedAssemblies @('System.dll','System.Core.dll','System.IO.Compression.dll','System.IO.Compression.FileSystem.dll')
    if (-not ('VbaToolkitNative' -as [type])) { Add-Type -Path (Join-Path $PSScriptRoot 'VbaToolkitNative.cs') }
}

function New-VbaSiblingPath([string]$InputPath, [string]$Suffix, [string]$Extension, [string[]]$AlsoExtensions = @()) {
    $folder = [IO.Path]::GetDirectoryName($InputPath)
    $stem = [IO.Path]::GetFileNameWithoutExtension($InputPath) + $Suffix
    $candidate = Join-Path $folder ($stem + $Extension)
    $number = 2
    while ($true) {
        $occupied = [IO.File]::Exists($candidate) -or [IO.Directory]::Exists($candidate)
        foreach ($other in $AlsoExtensions) { $occupied = $occupied -or [IO.File]::Exists([IO.Path]::ChangeExtension($candidate, $other)) }
        if (-not $occupied) { break }
        $candidate = Join-Path $folder ($stem + '_' + $number + $Extension)
        $number++
    }
    return $candidate
}

function Write-VbaNewTextFile([string]$Path, [string]$Text) {
    $stream = New-Object IO.FileStream($Path, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write)
    $writer = New-Object IO.StreamWriter($stream, (New-Object Text.UTF8Encoding($true)))
    try { $writer.Write($Text) } finally { $writer.Dispose() }
}

function Get-VbaPython {
    $python = [Environment]::GetEnvironmentVariable('VBATOOLS_PYTHON')
    if ($python) {
        $python = [IO.Path]::GetFullPath($python)
        if (-not [IO.File]::Exists($python)) { throw 'The Python executable in VBATOOLS_PYTHON was not found.' }
        return [pscustomobject]@{ Path = $python; Launcher = $false }
    }
    foreach ($name in @('python.exe','py.exe')) {
        foreach ($folder in [Environment]::GetEnvironmentVariable('PATH').Split([IO.Path]::PathSeparator)) {
            $directory = $folder.Trim().Trim('"').TrimEnd([char]92)
            if (-not $directory -or $directory.EndsWith('\Microsoft\WindowsApps', [StringComparison]::OrdinalIgnoreCase)) { continue }
            $candidate = [IO.Path]::Combine($directory, $name)
            if ([IO.File]::Exists($candidate)) { return [pscustomobject]@{ Path = $candidate; Launcher = ($name -eq 'py.exe') } }
        }
    }
    throw 'This action requires Python 3.10 or later. Set VBATOOLS_PYTHON to the path of python.exe.'
}

function Write-VbaScreen([string]$Status, [string]$Name, [string]$Detail = '', [ConsoleColor]$Color = 'Gray') {
    $width = 0
    foreach ($character in $Status.ToCharArray()) { $width += $(if ([int]$character -gt 255) { 2 } else { 1 }) }
    $line = '  ' + $Status + (' ' * [Math]::Max(1, 12 - $width)) + $Name
    Write-Host $line -ForegroundColor $Color
    if ($Detail) { Write-Host ('              ' + $Detail) -ForegroundColor DarkGray }
}
