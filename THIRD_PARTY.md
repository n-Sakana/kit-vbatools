# Reused source code

The supplied repositories were read without modification.

- **kit/KeyLoop**: installation wrappers, current-user menu registration, uninstall backup, Shell notification, classic conhost launch and LOCALAPPDATA conventions.
- **vba-devkit/lib/Diff.ps1**: module comparison, the greedy diff algorithm with a 100-line lookahead, and text reporting. Extraction uses Macro2Chat; comparison also detects case changes. Report data and HTML templates are separated.
- **vba-devkit/lib/Unlock.ps1**: marker search, the DPB-to-DPx patch and non-destructive copy behavior. Patching is restricted to the PROJECT stream; ZIP formats remain unchanged.
- **vba-devkit/lib/VBAToolkit.psm1**: OLE2 streams, ZIP part reading/writing and logging. The C# byte helper is a separate source file; no heredoc is used.
- **Macro2Chat/src/00_OutputFiles.cs and 05_Ole2.cs through 08_BookIO.cs**: included unchanged under `src/vendor/Macro2Chat`. The first file is a compilation dependency. These provide OLE2, compression, module and code-page reading in memory. The C# write-back path is not used.
- **Macro2Chat/python/macro2chat.py**: included unchanged in the same vendor folder. Diff, Markdown and Replace Code share its `extract` reader, including nested legacy XLS projects. Diff's earlier inventory pass uses its OLE container reader without VBA decompression, or ZIP member metadata. Replace Code also calls its `rebuild` function, including all-module cache removal, zero MODULEOFFSET, SRP removal, project-cache invalidation and output verification.

Source locations and SHA-256 hashes at import are recorded in the task's external work log.
