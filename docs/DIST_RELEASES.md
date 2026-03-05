# Dist Releases

All issue builds go to `dist` root.

## Create a new revision

Run from repo root:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\publish-dist.ps1
```

Optional runtime/config:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\publish-dist.ps1 -Runtime win-x64 -Configuration Release
```

## Output structure

- Latest (always overwritten in root):
  - `dist/DwgAutoPrinter.App.exe`
  - `dist/smart-revision-update.lsp`
- Revision snapshots (no folders):
  - `dist/DwgAutoPrinter.App.rev-0001.exe`
  - `dist/smart-revision-update.rev-0001.lsp`
- `dist/index.json`

Each run creates new revisioned files in root (`.rev-XXXX`) so every revision is kept without revision folders.
