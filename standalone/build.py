"""Build four single-file BAT tools from the shared, checked-in app sources."""
from pathlib import Path
import base64
import hashlib
import io
import json
import zipfile

root = Path(__file__).resolve().parent.parent
common = ['src/app/main.ps1', 'src/lib/Runtime.ps1']
files = {
    'Diff': common + ['src/lib/Diff.ps1', 'src/lib/VBAToolkit.psm1', 'src/lib/Replace.ps1', 'src/lib/replace.py', 'src/vendor/Macro2Chat/macro2chat.py', 'src/lib/report.html'],
    'Unlock': common + ['src/lib/Unlock.ps1', 'src/lib/VBAToolkit.psm1', 'src/lib/VbaToolkitNative.cs'],
    'Markdown': common + ['src/lib/Replace.ps1', 'src/lib/replace.py', 'src/vendor/Macro2Chat/macro2chat.py'],
    'Replace': common + ['src/lib/Replace.ps1', 'src/lib/replace.py', 'src/vendor/Macro2Chat/macro2chat.py'],
}
bootstrap = (root / 'standalone/bootstrap.ps1').read_text(encoding='utf-8-sig')
manifest = {}
for action, paths in files.items():
    memory = io.BytesIO()
    with zipfile.ZipFile(memory, 'w', compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for relative in paths:
            entry = zipfile.ZipInfo(relative, (2026, 10, 30, 0, 0, 0))
            entry.compress_type = zipfile.ZIP_DEFLATED
            archive.writestr(entry, (root / relative).read_bytes())
    header = '\n'.join([
        '@echo off >nul 2>&1', '@echo off', 'setlocal DisableDelayedExpansion',
        'set "VBATOOLS_SELF=%~f0"', f'set "VBATOOLS_ACTION={action}"', 'set "VBATOOLS_ARG_COUNT=0"',
        ':collect_arguments', 'if "%~1"=="" goto run_powershell', 'set /a VBATOOLS_ARG_COUNT+=1 >nul',
        'set "VBATOOLS_ARG_%VBATOOLS_ARG_COUNT%=%~1"', 'shift', 'goto collect_arguments', ':run_powershell',
        '"%SystemRoot%\\System32\\WindowsPowerShell\\v1.0\\powershell.exe" -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -Command "$s=[IO.File]::ReadAllText($env:VBATOOLS_SELF);$m=\'#__VBATOOLS_SCRIPT__\'; & ([scriptblock]::Create($s.Substring($s.LastIndexOf($m)+$m.Length)))"',
        'exit /b %ERRORLEVEL%', '#__VBATOOLS_SCRIPT__',
    ])
    text = header + '\n' + bootstrap + '\n# VBATOOLS_PAYLOAD=' + base64.b64encode(memory.getvalue()).decode('ascii') + '\n'
    target = root / 'standalone' / (action + '.bat')
    target.write_bytes(b'\xef\xbb\xbf' + text.replace('\r\n', '\n').replace('\n', '\r\n').encode('utf-8'))
    manifest[action] = {'bytes': target.stat().st_size, 'sha256': hashlib.sha256(target.read_bytes()).hexdigest(), 'sources': {name: hashlib.sha256((root / name).read_bytes()).hexdigest() for name in paths}}
(root / 'standalone/manifest.json').write_text(json.dumps(manifest, indent=2) + '\n', encoding='utf-8')
print(json.dumps({action: item['bytes'] for action, item in manifest.items()}))
