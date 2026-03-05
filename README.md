# DWGAutoPrinter

> Batch auto-printer for DWG/CAD drawing files. Configure once, print all.

---

## The Problem
Printing 50 DWG files in AutoCAD = 50x: open file → plot → pick printer → pick paper size → OK. DWGAutoPrinter does this in one click.

## Planned Features
- Batch select DWG files or entire folders
- Configure plot settings once (printer, paper size, scale)
- Auto-detect drawing paper size from title blocks
- Print to physical printer + PDF simultaneously
- Progress queue with status per file

## Status
🚧 **Planning phase** — architecture being designed

---

## Design
Deep Navy `#0A1628` + Engineering Blue `#1565C0` + AutoCAD Red `#E2231A`

---

*Target: Windows desktop utility for engineers*

---

## Dist Releases
Issue builds are created in `dist` and each run keeps a new revision folder.

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\publish-dist.ps1
```

Release index:
- `dist/index.json`

Latest issued UI app:
- `dist/rev-XXXX/DwgAutoPrinter.App.exe`
