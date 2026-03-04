# DWGAutoPrinter — Architecture

## Planned Architecture

```
DWGAutoPrinter/
├── src/
│   ├── app.py / main.js     # Entry point + UI
│   ├── dwg_reader.py        # DWG metadata extraction (title block, paper size)
│   ├── print_queue.py       # Batch print queue manager
│   ├── plot_settings.py     # Printer/paper/scale config (saved per printer)
│   └── autocad_bridge.py   # COM automation bridge (Option A)
├── assets/
│   └── logo.png / .ico
├── docs/
│   └── SPEC.md
└── CLAUDE.md
```

## DWG Processing Pipeline

```
DWG Files (input)
      │
      ├─→ dwg_reader ──→ extract title block (paper size, drawing no., rev)
      │
      ├─→ plot_settings ──→ apply printer config (remembered per paper size)
      │
      └─→ print_queue ──→ spool to printer(s) → progress UI → log
```

## Technology Decision Matrix

| Approach | AutoCAD Required | Quality | Complexity | Best For |
|----------|-----------------|---------|------------|----------|
| COM Automation | Yes | ★★★★★ | Medium | Power users with AutoCAD |
| ODA + PDF print | No | ★★★★☆ | Low | Broad distribution |
| ezdxf headless | No | ★★★☆☆ | Low | Metadata only |

## UI Stack Options

| Stack | Pros | Cons |
|-------|------|------|
| Python + customtkinter | Proven (like PDFRR) | Limited styling |
| Electron + React | Modern UI, web tech | Heavier footprint |
| Python + PyQt6 | Native widgets | GPL license |

**Recommendation**: Start with Python + customtkinter (same stack as PDFRR, proven workflow). Migrate to Electron if UI complexity grows.

---

## Related Projects
- **PDFRR** (`../PDFRR`) — PDF split/merge, shares the engineering tool DNA
  - Both target engineering/CAD workflows
  - Could share title block extraction logic

---

*Last Updated: 2026-03-04*
