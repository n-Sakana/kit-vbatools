# Standalone BAT tools

Copy any one of `Diff.bat`, `Unlock.bat`, `Markdown.bat` or `Replace.bat` to use that action. Each BAT contains its own PowerShell and compressed source payload. No installation, context-menu registration or sibling files are needed.

Drop files or folders onto a BAT. Selected folders are searched recursively for `.xlsm`, `.xlam` and `.xls`; overlapping selections are deduplicated. Outputs are placed next to the source workbooks without overwriting them. Diff requires exactly two VBA workbooks after expansion.

Double-click a BAT without arguments to open the standard Windows file picker. It supports multiple files and filters Excel VBA workbooks. Cancel exits without processing. Folder selection is supported by dropping folders, rather than through this picker.

All tools require Windows PowerShell 5.1. **Diff, Markdown and Replace require an existing Python 3.10 or later**; the embedded Macro2Chat Python script is not a Python interpreter. Set `VBATOOLS_PYTHON` to an absolute `python.exe` path if needed. Unlock requires no Python.

The same app entry, processing modules and TUI are embedded. Diff, Markdown and Replace share Macro2Chat's Python extraction path in both the standalone tools and installed kit, including legacy XLS projects under `_VBA_PROJECT_CUR`. Diff counts VBA projects from container metadata before extracting source, and proceeds only for exactly two books. Replace always uses the Python rebuild with complete execution-cache invalidation. No custom EXE or DLL is included or generated as a persistent app artifact.

For noninteractive replacement, `VBATOOLS_REPLACE_OPTIONS` can point to a UTF-8 JSON file containing `Find`, `Replacement` and `Match` (`Word` or `Partial`). Without that setting, Replace asks for the three values in the TUI.

The payload is extracted under `%LOCALAPPDATA%\VBATools\standalone`, identified by its SHA-256. These are source caches only. `bootstrap.ps1`, `build.py` and `manifest.json` are build materials; they are not required when running a copied BAT. Rebuild all four BAT files after changing shared code or the report template:

```text
python standalone/build.py
```
