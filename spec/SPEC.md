# DWG Auto Printer Specification

## Document Control
- Version: 1.0 (MVP draft)
- Date: 2026-03-05
- Source: popup intake + reviewed answers
- Target platform: Windows

## 1. Goal
Build a portable EXE that can be dropped into an architecture job folder, open AutoCAD, identify the latest revision from paper space title block data, update all paper space layouts, and print to PDF using AutoCAD default plotting behavior.

## 2. Confirmed Defaults (from intake)
- AutoCAD versions: 2022 to 2026
- Scan behavior: root folder only (not recursive)
- Revision source: title block in paper space
- Filename revision pattern (secondary only): `_REV##`
- Layout policy: all paper space layouts, model space excluded
- Output mode: PDF
- PDF output folder: `./PDF`
- PDF filename format: `<dwg number><rev>.pdf`
- Plot configuration baseline: `DWG To PDF` (AutoCAD standard)
- AutoCAD instance mode: reuse running instance when available
- Failure policy: retry 2 times, continue queue, show user friendly log list
- Skip existing PDF: no (always regenerate)
- Optional Excel revision register: not used for MVP

## 3. Scope
### In scope (MVP)
- Portable desktop app (single EXE)
- Scan DWG files in selected root folder
- Open drawings in configurable batches (e.g. 3 or 5)
- Read title block attributes from paper space layouts
- Group sheets by drawing number and select latest revision
- Open selected DWGs in AutoCAD via automation
- Trigger no-prompt LISP command(s) from app to update revisions
- Update each paper space layout before plotting
- Plot each layout to PDF
- Show progress and errors in a simple one-window UI

### Out of scope (MVP)
- Recursive folder watch mode
- Cloud sync/email
- CAD standard normalization
- Excel-based revision source

## 4. Functional Workflow
1. User drops EXE in job root and launches app.
2. App default root path = EXE directory (editable).
3. App scans `*.dwg` in root only.
4. App reads paper space title block attributes to extract:
   - drawing number
   - revision
   - layout name
5. App resolves latest revision per drawing number.
6. App builds queue from latest DWGs.
7. For each DWG:
   - open in AutoCAD
   - force paper space context before revision read/write
   - iterate all paper space layouts
   - update layout state (regen/update fields)
   - plot layout to PDF using AutoCAD defaults
8. App writes per-item status in UI log list and summary file.

## 5. Configuration Variables
| Key | Type | Default | Notes |
|---|---|---|---|
| `root_path` | string | EXE directory | User editable path |
| `scan_recursive` | bool | `false` | Root only |
| `dwg_pattern` | string | `*.dwg` | Input file filter |
| `open_batch_size` | int | `5` | AutoCAD open/processing batch size (3/5/custom) |
| `revision_source` | enum | `title_block` | Primary revision logic |
| `revision_pattern_fallback` | string | `_REV##` | Secondary fallback only |
| `layout_scope` | enum | `paperspace_all` | Process all paper space layouts |
| `include_model_space` | bool | `false` | Explicitly disabled |
| `output_mode` | enum | `pdf` | MVP PDF output |
| `pdf_output_dir` | string | `./PDF` | Created if missing |
| `pdf_name_pattern` | string | `<dwg number><rev>.pdf` | Sanitized at runtime |
| `plot_device` | string | `DWG To PDF.pc3` | Uses AutoCAD standard device |
| `plot_style_policy` | enum | `layout_current` | Respect current layout/page setup |
| `autocad_launch_mode` | enum | `reuse_running_instance` | Attach then fallback launch |
| `max_retry_count` | int | `2` | Per DWG/layout plot failure |
| `continue_on_error` | bool | `true` | Queue continues after retries |
| `skip_if_pdf_exists` | bool | `false` | Always re-plot |
| `target_autocad_versions` | string | `2022-2026` | Compatibility target |

## 6. Proposed App Structure
```
DWGAutoPrinter/
  spec/
    SPEC.md
    intake.reviewed.answers.json
  src/
    DwgAutoPrinter.App/
      Program.cs
      MainWindow.cs
    DwgAutoPrinter.Core/
      Models/
      Services/
      Interfaces/
    DwgAutoPrinter.Autocad/
      AutocadSession.cs
      TitleBlockReader.cs
      PlotService.cs
    DwgAutoPrinter.Infrastructure/
      ConfigStore.cs
      FileScanner.cs
      JobLogger.cs
  output/
    logs/
```

## 7. Core Modules and Responsibilities
- `FileScanner`: find DWG files in root path.
- `TitleBlockReader`: extract drawing number/revision from paper space title block attributes.
- `RevisionResolver`: choose latest revision per drawing number.
- `AutocadSession`: connect to running AutoCAD or launch one.
- `LayoutUpdater`: apply update steps per layout before plot (regen + field update).
- `LispAutomationBridge`: call parameterized no-prompt LISP commands via COM `SendCommand`.
- `PlotService`: send plot command to PDF device.
- `PrintQueueRunner`: orchestrate queue, retries, cancellation.
- `JobLogger`: write UI log list + text/json log file.
- `ConfigStore`: persist last used options.

## 8. Function-Level Contract (MVP)
```text
RunJob(rootPath, options) -> JobResult
ScanDwgs(rootPath, recursive=false) -> List<DwgFile>
ReadPaperSpaceMetadata(dwgPath) -> List<SheetMetadata>
ResolveLatestSheets(sheetList) -> List<LatestSheetSet>
OpenDwgInAutocad(dwgPath) -> AutocadDocumentHandle
UpdateLayout(doc, layoutName) -> void
PlotLayoutToPdf(doc, layoutName, outputPdfPath) -> PlotResult
ExecuteWithRetry(action, maxRetry=2) -> Result
BuildPdfPath(outputDir, drawingNumber, revision) -> string
LogEvent(level, message, context) -> void
InvokeLispNoPrompt(cmdOrExpr) -> Result
```

## 9. UI Specification (MVP)
- Single window.
- Controls:
  - root folder picker
  - scan button
  - file/revision grid preview
  - start/stop printing
  - live log list (user friendly messages)
  - summary counters (queued, success, failed)
- No advanced settings panel in MVP; only essential options.

## 10. Error Handling Rules
- Retry failed plot operation up to 2 times.
- If still failing, mark item `Failed` and continue.
- Keep AutoCAD session alive unless fatal COM failure.
- On fatal AutoCAD failure:
  - stop queue
  - show error banner
  - write crash-safe log entry

## 11. Packaging
- Build as portable Windows EXE.
- No installer required.
- Include runtime dependencies in publish output.
- MVP recommendation: .NET 8 single-file x64 publish with COM interop.

## 12. Validation Checklist
- Can attach to AutoCAD 2022 to 2026.
- Correctly ignores model space.
- Correctly processes all paper space layouts.
- Correctly resolves latest revision from title block.
- Produces PDF in `./PDF` with expected naming.
- Retry logic and log list behave as specified.

## 13. Open Technical Items
- Final title block attribute tag map needs sample DWGs.
- If some projects use custom LISP for revision extraction, define integration point in `TitleBlockReader` after MVP baseline.

## 14. Confirmed LISP Baseline
- Existing production baseline identified:
  - `C:/Users/kepne/OneDrive/Documents/#Architecture/ARCHER_CAD_TEMPLATE/DRAWING SHEET - Standard/bellway-revision-update.lsp`
- Repo upgraded baseline added:
  - `lisp/smart-revision-update.lsp`
- This upgraded script is the reference behavior for:
  - robust title block revision detection
  - safer layout-level validation
  - dry-run verification before writes
