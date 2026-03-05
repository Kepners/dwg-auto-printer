# Dist Releases

All issue builds go to `dist` and are never overwritten.

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

- `dist/rev-0001_YYYYMMDD-HHMMSS/`
  - `DwgAutoPrinter.App.exe`
  - `smart-revision-update.lsp`
- `dist/index.json`

Each run creates a new `rev-XXXX_...` folder so every revision is kept.
