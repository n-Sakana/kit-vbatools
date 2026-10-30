param(
    [Parameter(Mandatory)][string]$FileA,
    [Parameter(Mandatory)][string]$FileB,
    [Parameter(Mandatory)][string]$OutputPrefix,
    [Parameter(Mandatory)][hashtable]$Projects
)

$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot\VBAToolkit.psm1" -Force -DisableNameChecking

$nameA = [IO.Path]::GetFileName($FileA)
$nameB = [IO.Path]::GetFileName($FileB)
$sw = [System.Diagnostics.Stopwatch]::StartNew()

Write-VbaHeader 'Diff' "$nameA vs $nameB"
Write-VbaLog 'Diff' $FileA "Comparing with $nameB"

function Get-AllModules([string]$path) {
    $project = $Projects[$path]
    $result = [ordered]@{}
    foreach ($mod in $project.Modules) {
        $result[$mod.Name] = @{ Code = $mod.Code.Replace("`r`n", "`n"); Ext = $mod.Extension.TrimStart([char]46) }
    }
    return $result
}

# ============================================================================
# LCS diff algorithm
# ============================================================================

# Greedy approximate diff (not full LCS). Searches up to 100 lines ahead for sync points.
function Get-GreedyDiff([string[]]$a, [string[]]$b) {
    $m = $a.Count; $n = $b.Count
    # For large files, use a simpler greedy approach
    $result = [System.Collections.ArrayList]::new()
    $ia = 0; $ib = 0

    while ($ia -lt $m -or $ib -lt $n) {
        if ($ia -lt $m -and $ib -lt $n -and $a[$ia] -ceq $b[$ib]) {
            [void]$result.Add(@{ Type = 'equal'; LineA = $ia; LineB = $ib; TextA = $a[$ia]; TextB = $b[$ib] })
            $ia++; $ib++; continue
        }

        $bestAi = -1; $bestBi = -1; $bestDist = ($m + $n) * 2
        $searchA = [Math]::Min($ia + 100, $m)
        $searchB = [Math]::Min($ib + 100, $n)
        for ($ai = $ia; $ai -lt $searchA; $ai++) {
            for ($bi = $ib; $bi -lt $searchB; $bi++) {
                if ($a[$ai] -ceq $b[$bi]) {
                    $dist = ($ai - $ia) + ($bi - $ib)
                    if ($dist -lt $bestDist) { $bestDist = $dist; $bestAi = $ai; $bestBi = $bi }
                    break
                }
            }
        }
        if ($bestAi -eq -1) { $bestAi = $m; $bestBi = $n }

        # Pair up removed/added lines where possible
        $remCount = $bestAi - $ia; $addCount = $bestBi - $ib
        $pairCount = [Math]::Min($remCount, $addCount)
        for ($p = 0; $p -lt $pairCount; $p++) {
            [void]$result.Add(@{ Type = 'changed'; LineA = $ia; LineB = $ib; TextA = $a[$ia]; TextB = $b[$ib] })
            $ia++; $ib++
        }
        while ($ia -lt $bestAi) {
            [void]$result.Add(@{ Type = 'removed'; LineA = $ia; LineB = -1; TextA = $a[$ia]; TextB = '' })
            $ia++
        }
        while ($ib -lt $bestBi) {
            [void]$result.Add(@{ Type = 'added'; LineA = -1; LineB = $ib; TextA = ''; TextB = $b[$ib] })
            $ib++
        }
        $ia = $bestAi; $ib = $bestBi
    }
    return ,$result
}

# ============================================================================
# Build module diffs
# ============================================================================

$modsA = Get-AllModules $FileA
$modsB = Get-AllModules $FileB

$allNames = [System.Collections.ArrayList]::new()
foreach ($k in $modsA.Keys) { if ($allNames -notcontains $k) { [void]$allNames.Add($k) } }
foreach ($k in $modsB.Keys) { if ($allNames -notcontains $k) { [void]$allNames.Add($k) } }
$allNames = $allNames | Sort-Object

$moduleDiffs = [System.Collections.ArrayList]::new()
$added = 0; $removed = 0; $modified = 0; $unchanged = 0

foreach ($name in $allNames) {
    $inA = $modsA.Contains($name); $inB = $modsB.Contains($name)
    if ($inA -and -not $inB) {
        $removed++
        $ext = $modsA[$name].Ext
        Write-Host "  - $name.$ext (removed)" -ForegroundColor Red
        $linesA = $modsA[$name].Code -split "`n"
        $diff = [System.Collections.ArrayList]::new()
        for ($i = 0; $i -lt $linesA.Count; $i++) {
            [void]$diff.Add(@{ Type = 'removed'; LineA = $i; LineB = -1; TextA = $linesA[$i]; TextB = '' })
        }
        [void]$moduleDiffs.Add(@{ Name = "$name.$ext"; Status = 'removed'; Diff = $diff })
    }
    elseif (-not $inA -and $inB) {
        $added++
        $ext = $modsB[$name].Ext
        Write-Host "  + $name.$ext (added)" -ForegroundColor Green
        $linesB = $modsB[$name].Code -split "`n"
        $diff = [System.Collections.ArrayList]::new()
        for ($i = 0; $i -lt $linesB.Count; $i++) {
            [void]$diff.Add(@{ Type = 'added'; LineA = -1; LineB = $i; TextA = ''; TextB = $linesB[$i] })
        }
        [void]$moduleDiffs.Add(@{ Name = "$name.$ext"; Status = 'added'; Diff = $diff })
    }
    else {
        $codeA = $modsA[$name].Code; $codeB = $modsB[$name].Code
        $ext = $modsA[$name].Ext
        if ($codeA -ceq $codeB) {
            $unchanged++
            $linesA = $codeA -split "`n"
            $diff = [System.Collections.ArrayList]::new()
            for ($i = 0; $i -lt $linesA.Count; $i++) {
                [void]$diff.Add(@{ Type = 'equal'; LineA = $i; LineB = $i; TextA = $linesA[$i]; TextB = $linesA[$i] })
            }
            [void]$moduleDiffs.Add(@{ Name = "$name.$ext"; Status = 'unchanged'; Diff = $diff })
        } else {
            $modified++
            Write-Host "  ~ $name.$ext (modified)" -ForegroundColor Yellow
            $diff = Get-GreedyDiff ($codeA -split "`n") ($codeB -split "`n")
            [void]$moduleDiffs.Add(@{ Name = "$name.$ext"; Status = 'modified'; Diff = $diff })
        }
    }
}

Write-Host ""
$parts = @()
if ($added) { $parts += "$added added" }
if ($removed) { $parts += "$removed removed" }
if ($modified) { $parts += "$modified modified" }
if ($unchanged) { $parts += "$unchanged unchanged" }
Write-Host "Summary: $($parts -join ', ')"

# Versioned report data is independent of the replaceable HTML template.
$modules = New-Object Collections.Generic.List[object]
foreach ($module in $moduleDiffs) {
    $rows = New-Object Collections.Generic.List[object]
    foreach ($row in $module.Diff) {
        if ($row.Type -eq 'equal') {
            $rows.Add([pscustomobject]@{ kind = 'equal'; oldLine = $row.LineA + 1; newLine = $row.LineB + 1; text = $row.TextA })
        } else {
            if ($row.Type -eq 'changed' -or $row.Type -eq 'removed') {
                $rows.Add([pscustomobject]@{ kind = 'remove'; oldLine = $row.LineA + 1; newLine = $null; text = $row.TextA })
            }
            if ($row.Type -eq 'changed' -or $row.Type -eq 'added') {
                $rows.Add([pscustomobject]@{ kind = 'add'; oldLine = $null; newLine = $row.LineB + 1; text = $row.TextB })
            }
        }
    }
    $modules.Add([pscustomobject]@{ name = $module.Name; status = $module.Status; rows = @($rows.ToArray()) })
}
$data = [pscustomobject]@{
    schema = 1
    generated = [DateTime]::UtcNow.ToString('o')
    files = [pscustomobject]@{ a = [pscustomobject]@{ name = $nameA; path = $FileA }; b = [pscustomobject]@{ name = $nameB; path = $FileB } }
    summary = [pscustomobject]@{ added = $added; removed = $removed; modified = $modified; unchanged = $unchanged }
    modules = @($modules.ToArray())
}
$json = $data | ConvertTo-Json -Depth 10
$template = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'report.html'))
# Prevent source text from closing the inert JSON script element.
$safeJson = $json.Replace('&', '\u0026').Replace('<', '\u003c').Replace('>', '\u003e').Replace([string][char]0x2028, '\u2028').Replace([string][char]0x2029, '\u2029')
$title = [Net.WebUtility]::HtmlEncode('VBA Diff')
$subtitle = [Net.WebUtility]::HtmlEncode($nameA + ' / ' + $nameB)
$substitutions = @{ '{{TITLE}}' = $title; '{{SUBTITLE}}' = $subtitle; '{{DIFF_JSON}}' = $safeJson }
$html = [regex]::Replace($template, '\{\{(?:TITLE|SUBTITLE|DIFF_JSON)\}\}', [Text.RegularExpressions.MatchEvaluator]{ param($match) [string]$substitutions[$match.Value] })
Write-VbaNewTextFile ($OutputPrefix + '.json') $json
Write-VbaNewTextFile ($OutputPrefix + '.html') $html

# Text report
$textReport = [System.Text.StringBuilder]::new()
[void]$textReport.AppendLine("# VBA Diff Report")
[void]$textReport.AppendLine("# A: $nameA")
[void]$textReport.AppendLine("# B: $nameB")
[void]$textReport.AppendLine("# Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
[void]$textReport.AppendLine("")
[void]$textReport.AppendLine("Added: $added, Removed: $removed, Modified: $modified, Unchanged: $unchanged")
[void]$textReport.AppendLine("")

$addedMods = @($moduleDiffs | Where-Object { $_.Status -eq 'added' })
$removedMods = @($moduleDiffs | Where-Object { $_.Status -eq 'removed' })
$modifiedMods = @($moduleDiffs | Where-Object { $_.Status -eq 'modified' })
$unchangedMods = @($moduleDiffs | Where-Object { $_.Status -eq 'unchanged' })

if ($addedMods) {
    [void]$textReport.AppendLine("Added modules:")
    foreach ($m in $addedMods) { [void]$textReport.AppendLine("  + $($m.Name)") }
    [void]$textReport.AppendLine("")
}
if ($removedMods) {
    [void]$textReport.AppendLine("Removed modules:")
    foreach ($m in $removedMods) { [void]$textReport.AppendLine("  - $($m.Name)") }
    [void]$textReport.AppendLine("")
}
if ($modifiedMods) {
    [void]$textReport.AppendLine("Modified modules:")
    foreach ($m in $modifiedMods) { [void]$textReport.AppendLine("  ~ $($m.Name)") }
    [void]$textReport.AppendLine("")
}
if ($unchangedMods) {
    [void]$textReport.AppendLine("Unchanged: $($unchangedMods.Count) module(s)")
    [void]$textReport.AppendLine("")
}

foreach ($module in $moduleDiffs) {
    [void]$textReport.AppendLine(('=== {0} [{1}] ===' -f $module.Name, $module.Status))
    if ($module.Status -ne 'unchanged') {
        Write-Host ''
        Write-Host ('  ' + $module.Name) -ForegroundColor Cyan
    }
    foreach ($row in $module.Diff) {
        if ($row.Type -eq 'changed' -or $row.Type -eq 'removed') {
            [void]$textReport.AppendLine('- ' + $row.TextA)
            Write-Host ('  - {0,5}  {1}' -f ($row.LineA + 1), $row.TextA) -ForegroundColor DarkGray
        }
        if ($row.Type -eq 'changed' -or $row.Type -eq 'added') {
            [void]$textReport.AppendLine('+ ' + $row.TextB)
            Write-Host ('  + {0,5}  {1}' -f ($row.LineB + 1), $row.TextB) -ForegroundColor Cyan
        }
        if ($row.Type -eq 'equal') { [void]$textReport.AppendLine('  ' + $row.TextA) }
    }
    [void]$textReport.AppendLine('')
}
Write-VbaNewTextFile ($OutputPrefix + '.txt') $textReport.ToString()



$sw.Stop()
Write-VbaResult 'Diff' "$nameA vs $nameB" "$($parts -join ', ')" $OutputPrefix $sw.Elapsed.TotalSeconds
Write-VbaLog 'Diff' $FileA "vs $nameB | $($parts -join ', ') | -> $OutputPrefix"
