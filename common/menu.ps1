# Adapted from KeyLoop/common/menu.ps1: current-user registry only.
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$script:StandaloneRoot = [IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent))
$script:App = [IO.File]::ReadAllText((Join-Path $script:StandaloneRoot 'app.json')) | ConvertFrom-Json
if ($script:App.Id -ne 'VBATools' -or $script:App.MenuKey -ne 'kit-vbatools') { throw 'Unknown application.' }

function Get-MenuPlan {
    $powershell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $entry = Join-Path $script:StandaloneRoot 'src\app\collect.ps1'
    $base = '"{0}" -NoLogo -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "{1}"' -f $powershell, $entry
    foreach ($context in @('Software\Classes\*\shell', 'Software\Classes\Directory\shell')) {
    [pscustomobject]@{
        Path = $context + '\' + $script:App.MenuKey
        Label = $script:App.Name
        Children = @(
            [pscustomobject]@{ Id = '01-diff'; Label = 'VBA Diff'; Command = $base + ' -Action Diff -Target "%1"' },
            [pscustomobject]@{ Id = '02-unlock'; Label = 'Unlock'; Command = $base + ' -Action Unlock -Target "%1"' },
            [pscustomobject]@{ Id = '03-markdown'; Label = 'Export Markdown'; Command = $base + ' -Action Markdown -Target "%1"' },
            [pscustomobject]@{ Id = '04-replace'; Label = 'Replace Code'; Command = $base + ' -Action Replace -Target "%1"' }
        )
    }
    }
}

function Send-MenuChanged {
    if (-not ('VBAToolsMenuNotify' -as [type])) {
        Add-Type 'using System; using System.Runtime.InteropServices; public static class VBAToolsMenuNotify { [DllImport("shell32.dll")] public static extern void SHChangeNotify(uint e, uint f, IntPtr a, IntPtr b); }'
    }
    [VBAToolsMenuNotify]::SHChangeNotify(0x08000000, 0, [IntPtr]::Zero, [IntPtr]::Zero)
}

function Install-StandaloneMenu {
    if ($env:OS -ne 'Windows_NT') { throw 'Windows is required.' }
    foreach ($file in @('src\app\collect.ps1', 'src\app\main.ps1', 'src\lib\Diff.ps1', 'src\lib\Unlock.ps1', 'src\lib\Replace.ps1', 'src\vendor\Macro2Chat\macro2chat.py')) {
        if (-not [IO.File]::Exists((Join-Path $script:StandaloneRoot $file))) { throw ('Missing file: ' + $file) }
    }
    foreach ($item in @(Get-MenuPlan)) {
        [Microsoft.Win32.Registry]::CurrentUser.DeleteSubKeyTree($item.Path, $false)
        $key = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($item.Path)
        try {
            $key.SetValue('MUIVerb', $item.Label)
            $key.SetValue('InstallRoot', $script:StandaloneRoot)
            $key.SetValue('SubCommands', '')
            $key.SetValue('MultiSelectModel', 'Player')
            foreach ($child in $item.Children) {
                $sub = $key.CreateSubKey('shell\' + $child.Id)
                try {
                    $sub.SetValue('MUIVerb', $child.Label)
                    $sub.SetValue('MultiSelectModel', 'Player')
                    $command = $sub.CreateSubKey('command')
                    try { $command.SetValue('', $child.Command) } finally { $command.Dispose() }
                } finally { $sub.Dispose() }
            }
        } finally { $key.Dispose() }
    }
    Send-MenuChanged
}
