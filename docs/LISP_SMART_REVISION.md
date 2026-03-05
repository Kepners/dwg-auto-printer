# Smart Revision LISP (EXE-Controlled)

## File
- Script: [`lisp/smart-revision-update.lsp`](/c:/Users/kepne/projects/DWGAutoPrinter/lisp/smart-revision-update.lsp)

## Behavior
- On load, the script does nothing except define functions.
- Manual command entry points are intentionally removed.
- The script runs only when the EXE sends the API call.

## EXE API Call
```lisp
(dap:exe-run "C:/JobFolder" 5 "NEXT" "" "Construction Issue" T)
```

Parameters:
- `folder`: root DWG folder (root only)
- `batch-size`: 3/5/custom (used by EXE batching)
- `rev-mode`: `NEXT` or `EXACT`
- `exact-rev`: value used only when `EXACT`
- `target-desc`: revision description text
- `close-after`: `T` or `NIL` (EXE orchestration flag)

## Safety Note
- Multi-document open/close from inside LISP is disabled to reduce AutoCAD crash risk.
- Batch opening/closing is handled by the EXE controller.
