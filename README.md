# VBATools

Four VBA tools in the Windows file context menu. The menu title is **VBA Tools** and the registration key is `kit-vbatools`.

Windows 11 and Windows PowerShell 5.1 are required. Unlock uses Windows standard components. Diff, Markdown and Replace Code use an existing Python 3.10 or later. No administrator rights, custom EXE, DLL or shell extension are required. The app never starts Excel, a browser or a dialog.

## Install and remove

Run `install.bat`, select files, then right-click > **Show more options** > **VBA Tools**. Shift + right-click also opens the classic menu. Keep this folder in place; run `install.bat` again after moving or renaming it. Run `uninstall.bat` to remove the menu.

Only `HKCU\Software\Classes\*\shell\kit-vbatools` and `HKCU\Software\Classes\Directory\shell\kit-vbatools` are registered. Uninstall backs them up under `%LOCALAPPDATA%\VBATools\menu-backups`. It leaves workbooks, reports and logs in place. `install.bat -PlanOnly` and `uninstall.bat -PlanOnly` display the planned changes.

## Actions

| Menu | Selection | Output next to the source workbook |
| --- | --- | --- |
| VBA Diff | Two different workbooks | `name_vba-diff.txt`, `.html` and `.json`, next to A |
| Unlock | One or more workbooks | `name_unlocked.xlsm`, retaining the source extension |
| Export Markdown | One or more workbooks | `name_vba.md` |
| Replace Code | One or more workbooks | `name_replaced.xlsm`, retaining the source extension |

Sources are never overwritten. Existing output names receive `_2`, `_3`, etc. A and B follow the order in which Explorer delivers the selected files.

Files and folders can be mixed. Selected folders are searched recursively for `.xlsm`, `.xlam` and `.xls`; overlapping selections are deduplicated. Temporary `~$` files and directory junctions/reparse points are skipped. Diff counts containers with a VBA project before decompressing any VBA source and requires exactly two. The other actions process the full inventory once, with outputs next to each book.

**VBA Diff** first counts VBA-bearing workbooks using only container metadata. A count other than two displays `VBA Diff: found N VBA workbook(s). Select exactly two files.` before extracting module source. With exactly two, Diff uses the same Macro2Chat Python `extract` reader as Markdown and Replace Code, including nested legacy XLS projects. It compares module additions, removals and changes, including case changes. Changed lines appear in the TUI and the text report. The HTML and versioned JSON reports are saved alongside it; no browser opens.

**Unlock** uses the supplied `DPB=` to `DPx=` patch on the PROJECT stream of a separate copy. ZIP workbooks retain their format; there is no SaveAs conversion. A missing marker is reported and skipped without creating a copy. This does not remove a password required to open a workbook, or worksheet protection. Excel/VBE may ask about the invalidated password marker when the copy is opened. To clear the protection settings permanently, clear the password fields in VBE and save the copy.

**Export Markdown** uses the same bundled Macro2Chat Python `extract` reader as Replace Code, in both the kit and standalone tool. It reads all standard, class, document and form code modules into one UTF-8 BOM Markdown file per workbook. This includes legacy `.xls` projects nested under `_VBA_PROJECT_CUR`. Empty modules remain in the index and headings. Form designs and images are excluded. Code pages are respected. Partial or damaged source is reported as a failure.

**Replace Code** asks once for the search text, replacement text and match type, then applies them to all selected workbooks. Matching is literal and case sensitive. An empty replacement deletes matches. Whole word matches require both neighbors to be outside `A-Z`, `a-z`, `0-9` and `_`; partial matches also replace inside other text. Code comments and string literals are included. Module names and preserved attribute headers are excluded.

The result shows changed module and replacement counts. Zero matches create no copy. All writes use the bundled Macro2Chat Python `extract` then `rebuild`: every module cache is removed, MODULEOFFSET becomes zero, SRP caches are removed, and the project execution cache is invalidated. Rebuild reopens and verifies its output. Existing VBA signatures are removed. Check the resulting VBA syntax and behavior in the output copy.

Replace Code supports `.xlsm`, `.xlam`, `.xlsb` and `.xls`. Markdown reads both OLE2 and ZIP-based VBA workbooks. Both actions find Python as `python.exe` or `py.exe -3` on PATH, excluding Microsoft Store aliases. Set `VBATOOLS_PYTHON` to an absolute `python.exe` path to choose it explicitly.

The other actions accept `.xls`, `.xla`, `.xlt`, `.xlsm`, `.xlam`, `.xlsb`, `.xltm`, `.xlsx` and `.xltx`. Unlock skips macro-free books; Diff and Markdown report that VBA could not be read. Whole-file encryption is unsupported. One failed workbook does not prevent the remaining selected books from being processed.

## TUI and selection collection

The console shows the current file, item count, progress, result, output path and final totals with a small color palette. It stays open until a key is pressed. No dialogs or GUI applications are opened.

Ordinary Windows verbs start once per file. `collect.ps1` groups those calls using a per-user, per-action mutex and `%LOCALAPPDATA%\VBATools\queue`. It waits at least 2.5 seconds, then finishes after 1.5 seconds without another arrival, with a 15-second maximum per burst. Repeated invocations of the same action within this interval may merge. Very large selections arriving over more than 15 seconds can form separate groups; use the direct array entry below for those cases.

Only queues, temporary extraction data and logs use `%LOCALAPPDATA%\VBATools`. No DLL is cached. Results are also available in `results\*.json`.

## Direct PowerShell use

```powershell
& '.\src\app\main.ps1' -Action Diff -Paths @('C:\work\before.xlsm', 'C:\work\after.xlsm')
& '.\src\app\main.ps1' -Action Unlock -Paths @('C:\work\one.xlsm', 'C:\work\two.xls')
& '.\src\app\main.ps1' -Action Markdown -Paths @('C:\work\one.xlsm', 'C:\work\two.xls')
& '.\src\app\main.ps1' -Action Replace -Paths @('C:\work\one.xlsm', 'C:\work\two.xls')
& '.\src\app\main.ps1' -Action Replace -Paths @('C:\work\one.xlsm') -ReplaceOptionsProvided -FindText 'oldName' -ReplaceText 'newName' -MatchKind Word -NoPause
```

Use `-NoPause` for automation. Success, no marker and no changes return exit code 0; any failure returns 1. `-ReplaceOptionsProvided` supplies replacement settings without interactive prompts.

## HTML template contract

Replace `src/lib/report.html` to change the HTML presentation. The app replaces three literal tokens: `{{TITLE}}` and `{{SUBTITLE}}` contain HTML-escaped text; `{{DIFF_JSON}}` contains the full JSON with HTML-sensitive characters escaped. Keep the JSON in an inert `<script type="application/json" id="vba-diff-data">` element and parse its `textContent`. Render source through `textContent`, never `innerHTML`. The shipped template has no CSS.

JSON schema 1 has `generated`, `files.a/b` (`name`, `path`), `summary` (`added`, `removed`, `modified`, `unchanged`) and `modules[]`. Each module has `name` including extension, `status` (`added`, `removed`, `modified`, `unchanged`) and `rows[]`. Each row has `kind` (`add`, `remove`, `equal`), `oldLine`, `newLine` and `text`. Line numbers are one based; an absent side is null. A changed line is a remove row followed by an add row. The adjacent JSON file contains the same unescaped data.

See [THIRD_PARTY.md](THIRD_PARTY.md) for reused source code.
