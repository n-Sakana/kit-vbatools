# Adapted from toolrack vba-devkit/lib/VBAToolkit.psm1. See THIRD_PARTY.md.
$ErrorActionPreference = 'Stop'

function Get-VbaLogPath {
    $folder = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'VBATools'
    [void][IO.Directory]::CreateDirectory($folder)
    return (Join-Path $folder 'operations.log')
}



# ============================================================================
# OLE2 Compound Document Parser
# ============================================================================

function Read-SectorChain([byte[]]$bytes, [int]$startSector, [int]$sectorSize, [int[]]$fat) {
    return ,[VbaToolkitNative]::ReadSectorChain($bytes, $startSector, $sectorSize, $fat)
}

function Read-MiniStream([byte[]]$miniStreamData, [int]$startSector, [int]$size, [int]$miniSectorSize, [int[]]$miniFat) {
    return ,[VbaToolkitNative]::ReadMiniStream($miniStreamData, $startSector, $size, $miniSectorSize, $miniFat)
}

function Read-Ole2([byte[]]$bytes) {
    $sectorPow = [BitConverter]::ToUInt16($bytes, 30)
    $sectorSize = [int][Math]::Pow(2, $sectorPow)
    $miniSectorPow = [BitConverter]::ToUInt16($bytes, 32)
    $miniSectorSize = [int][Math]::Pow(2, $miniSectorPow)
    $firstDirSector = [BitConverter]::ToInt32($bytes, 48)
    $miniStreamCutoff = [BitConverter]::ToUInt32($bytes, 56)
    $firstMiniFatSector = [BitConverter]::ToInt32($bytes, 60)
    $firstDifatSector = [BitConverter]::ToInt32($bytes, 68)

    $difat = [System.Collections.ArrayList]::new()
    for ($i = 0; $i -lt 109; $i++) {
        $v = [BitConverter]::ToInt32($bytes, 76 + $i * 4)
        if ($v -ge 0) { [void]$difat.Add($v) }
    }
    $nextDifat = $firstDifatSector
    while ($nextDifat -ge 0 -and $nextDifat -ne -2) {
        $off = ($nextDifat + 1) * $sectorSize
        $ePS = $sectorSize / 4 - 1
        for ($i = 0; $i -lt $ePS; $i++) {
            $v = [BitConverter]::ToInt32($bytes, $off + $i * 4)
            if ($v -ge 0) { [void]$difat.Add($v) }
        }
        $nextDifat = [BitConverter]::ToInt32($bytes, $off + $ePS * 4)
    }

    $fatEntries = [int]($bytes.Length / $sectorSize)
    [int[]]$fat = New-Object int[] $fatEntries
    for ($i = 0; $i -lt $fatEntries; $i++) { $fat[$i] = -1 }
    $idx = 0
    foreach ($ds in $difat) {
        $off = ($ds + 1) * $sectorSize
        for ($i = 0; $i -lt ($sectorSize / 4) -and $idx -lt $fatEntries; $i++) {
            $fat[$idx++] = [BitConverter]::ToInt32($bytes, $off + $i * 4)
        }
    }

    $dirData = Read-SectorChain $bytes $firstDirSector $sectorSize $fat
    $entries = [System.Collections.ArrayList]::new()
    for ($i = 0; $i -lt [int]($dirData.Length / 128); $i++) {
        $eOff = $i * 128
        $nameLen = [BitConverter]::ToUInt16($dirData, $eOff + 64)
        $name = ''
        if ($nameLen -gt 2) {
            $name = [System.Text.Encoding]::Unicode.GetString($dirData, $eOff, $nameLen - 2)
        }
        [void]$entries.Add([PSCustomObject]@{
            Name = $name; ObjType = $dirData[$eOff + 66]
            Start = [BitConverter]::ToInt32($dirData, $eOff + 116)
            Size = [BitConverter]::ToUInt32($dirData, $eOff + 120)
            DirOffset = $eOff
        })
    }

    [int[]]$miniFat = @()
    if ($firstMiniFatSector -ge 0 -and $firstMiniFatSector -ne -2) {
        $mfData = Read-SectorChain $bytes $firstMiniFatSector $sectorSize $fat
        $miniFat = New-Object int[] ([int]($mfData.Length / 4))
        for ($i = 0; $i -lt $miniFat.Length; $i++) {
            $miniFat[$i] = [BitConverter]::ToInt32($mfData, $i * 4)
        }
    }

    $rootEntry = $entries | Where-Object { $_.ObjType -eq 5 } | Select-Object -First 1
    [byte[]]$miniStreamData = @()
    if ($rootEntry -and $rootEntry.Start -ge 0) {
        $miniStreamData = Read-SectorChain $bytes $rootEntry.Start $sectorSize $fat
    }

    return @{
        Entries = $entries; Bytes = $bytes; SectorSize = $sectorSize
        MiniSectorSize = $miniSectorSize; MiniStreamCutoff = $miniStreamCutoff
        Fat = $fat; MiniFat = $miniFat; MiniStreamData = $miniStreamData
        FirstDirSector = $firstDirSector
    }
}

function Read-Ole2Stream($ole2, $entry) {
    if ($entry.Size -lt $ole2.MiniStreamCutoff -and $ole2.MiniStreamData.Length -gt 0) {
        return Read-MiniStream $ole2.MiniStreamData $entry.Start $entry.Size $ole2.MiniSectorSize $ole2.MiniFat
    } else {
        $raw = Read-SectorChain $ole2.Bytes $entry.Start $ole2.SectorSize $ole2.Fat
        if ($raw.Length -gt $entry.Size) {
            $trimmed = New-Object byte[] $entry.Size
            [Array]::Copy($raw, $trimmed, $entry.Size)
            return ,$trimmed
        }
        return ,$raw
    }
}

function Write-Ole2Stream([byte[]]$ole2Bytes, $ole2, $entry, [byte[]]$newData) {
    $sectorSize = $ole2.SectorSize
    $fat = $ole2.Fat

    if ($entry.Size -lt $ole2.MiniStreamCutoff) {
        $miniSectorSize = $ole2.MiniSectorSize
        $miniFat = $ole2.MiniFat
        $s = $entry.Start; $written = 0
        while ($s -ge 0 -and $s -ne -2 -and $written -lt $newData.Length) {
            $off = $s * $miniSectorSize
            $toWrite = [Math]::Min($miniSectorSize, $newData.Length - $written)
            [Array]::Copy($newData, $written, $ole2.MiniStreamData, $off, $toWrite)
            if ($toWrite -lt $miniSectorSize) {
                for ($p = $toWrite; $p -lt $miniSectorSize; $p++) { $ole2.MiniStreamData[$off + $p] = 0 }
            }
            $written += $miniSectorSize
            if ($s -lt $miniFat.Length) { $s = $miniFat[$s] } else { break }
        }
        $rootEntry = $ole2.Entries | Where-Object { $_.ObjType -eq 5 } | Select-Object -First 1
        $s2 = $rootEntry.Start; $written2 = 0; $visited = @{}
        while ($s2 -ge 0 -and $s2 -ne -2 -and -not $visited.ContainsKey($s2) -and $written2 -lt $ole2.MiniStreamData.Length) {
            $visited[$s2] = $true
            $off2 = ($s2 + 1) * $sectorSize
            [Array]::Copy($ole2.MiniStreamData, $written2, $ole2Bytes, $off2, [Math]::Min($sectorSize, $ole2.MiniStreamData.Length - $written2))
            $written2 += $sectorSize; $s2 = $fat[$s2]
        }
        # Validate: count sectors actually used
        $chainLen = 0; $sv = $entry.Start; $seen = @{}
        while ($sv -ge 0 -and $sv -ne -2 -and -not $seen.ContainsKey($sv)) { $seen[$sv] = $true; $chainLen++; $sv = if ($sv -lt $miniFat.Length) { $miniFat[$sv] } else { -1 } }
        $chainCapacity = $chainLen * $miniSectorSize
        if ($newData.Length -gt $chainCapacity) {
            throw "Write-Ole2Stream: data truncated (mini stream). Data=$($newData.Length) bytes, chain capacity=$chainCapacity bytes."
        }
    } else {
        $s = $entry.Start; $written = 0; $visited = @{}
        while ($s -ge 0 -and $s -ne -2 -and -not $visited.ContainsKey($s) -and $written -lt $newData.Length) {
            $visited[$s] = $true
            $off = ($s + 1) * $sectorSize
            $toWrite = [Math]::Min($sectorSize, $newData.Length - $written)
            [Array]::Copy($newData, $written, $ole2Bytes, $off, $toWrite)
            if ($toWrite -lt $sectorSize) {
                for ($p = $toWrite; $p -lt $sectorSize; $p++) { $ole2Bytes[$off + $p] = 0 }
            }
            $written += $sectorSize; $s = $fat[$s]
        }
        # Validate: count sectors actually used
        $chainLen = 0; $sv = $entry.Start; $visitedV = @{}
        while ($sv -ge 0 -and $sv -ne -2 -and -not $visitedV.ContainsKey($sv)) { $visitedV[$sv]=$true; $chainLen++; $sv = $fat[$sv] }
        $chainCapacity = $chainLen * $sectorSize
        if ($newData.Length -gt $chainCapacity) {
            throw "Write-Ole2Stream: data truncated. Data=$($newData.Length) bytes, chain capacity=$chainCapacity bytes."
        }
    }

    # Update size in directory
    $dirSectorData = Read-SectorChain $ole2Bytes $ole2.FirstDirSector $sectorSize $fat
    [Array]::Copy([BitConverter]::GetBytes([uint32]$newData.Length), 0, $dirSectorData, $entry.DirOffset + 120, 4)
    $s3 = $ole2.FirstDirSector; $written3 = 0; $visited3 = @{}
    while ($s3 -ge 0 -and $s3 -ne -2 -and -not $visited3.ContainsKey($s3)) {
        $visited3[$s3] = $true
        [Array]::Copy($dirSectorData, $written3, $ole2Bytes, ($s3 + 1) * $sectorSize, [Math]::Min($sectorSize, $dirSectorData.Length - $written3))
        $written3 += $sectorSize; $s3 = $fat[$s3]
    }
}


function Get-VbaProjectBytes([string]$filePath) {
    $ext = [IO.Path]::GetExtension($filePath).ToLower()
    if ($ext -in @('.xls','.xla','.xlt')) {
        return @{ Bytes = [IO.File]::ReadAllBytes($filePath); IsZip = $false }
    }
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $block = [ScriptBlock]::Create('
        param($path)
        $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
        try {
            $entry = $zip.Entries | Where-Object { $_.Name -eq "vbaProject.bin" } | Select-Object -First 1
            if (-not $entry) { return $null }
            $s = $entry.Open(); $ms = New-Object IO.MemoryStream; $s.CopyTo($ms); $s.Close()
            return ,$ms.ToArray()
        } finally { $zip.Dispose() }
    ')
    $bytes = & $block $filePath
    return @{ Bytes = $bytes; IsZip = $true }
}

function Save-VbaProjectBytes([string]$filePath, [byte[]]$ole2Bytes, [bool]$isZip) {
    if (-not $ole2Bytes -or $ole2Bytes.Length -eq 0) {
        throw "Save-VbaProjectBytes: ole2Bytes is null or empty"
    }
    if ($isZip) {
        Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
        $zip = [System.IO.Compression.ZipFile]::Open("$filePath", [System.IO.Compression.ZipArchiveMode]::Update)
        try {
            $entry = $zip.Entries | Where-Object { $_.Name -eq 'vbaProject.bin' } | Select-Object -First 1
            if (-not $entry) { throw "Save-VbaProjectBytes: vbaProject.bin not found in ZIP" }
            $stream = $entry.Open()
            $stream.SetLength(0)
            $stream.Write($ole2Bytes, 0, $ole2Bytes.Length)
            $stream.Flush()
            $stream.Close()
        } finally {
            $zip.Dispose()
        }
    } else {
        [IO.File]::WriteAllBytes("$filePath", $ole2Bytes)
    }
    # Post-write check is an integrity check, not a second input validation.
    $savedSize = (New-Object IO.FileInfo($filePath)).Length
    if ($savedSize -eq 0) {
        throw "Save-VbaProjectBytes: file is 0 bytes after write: $filePath"
    }
}


function Write-VbaLog {
    param([string]$ToolName, [string]$InputFile, [string]$Message, [string]$Level = 'INFO')
    $logPath = Get-VbaLogPath
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $fileName = [IO.Path]::GetFileName($InputFile)
    $entry = "[$timestamp] [$Level] [$ToolName] $fileName | $Message"
    [IO.File]::AppendAllText($logPath, "$entry`r`n", [System.Text.Encoding]::UTF8)
}

function Write-VbaStatus {
    param([string]$ToolName, [string]$FileName, [string]$Message)
    Write-Host "  $Message" -ForegroundColor Gray
}

function Write-VbaResult {
    param([string]$ToolName, [string]$FileName, [string]$Summary, [string]$OutputDir, [double]$ElapsedSec)
    Write-Host "  $Summary" -ForegroundColor Green
    if ($OutputDir) { Write-Host "  Output: $OutputDir" -ForegroundColor Gray }
    if ($ElapsedSec -gt 0) { Write-Host "  Done ($([Math]::Round($ElapsedSec, 1))s)" -ForegroundColor Gray }
}

function Write-VbaError {
    param([string]$ToolName, [string]$FileName, [string]$Message)
    Write-Host "  ERROR: $Message" -ForegroundColor Red
    Write-VbaLog $ToolName $FileName $Message 'ERROR'
}

function Write-VbaHeader {
    param([string]$ToolName, [string]$FileName)
    Write-Host "[$ToolName] $FileName" -ForegroundColor Cyan
}

