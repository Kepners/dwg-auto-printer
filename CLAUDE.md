# DWGAutoPrinter

## Project Overview

**DWGAutoPrinter** — Batch auto-printer for DWG/CAD drawing files. Automates the tedious process of printing multiple AutoCAD DWG files to physical printers or PDF, with smart plot settings detection.

| Item | Value |
|------|-------|
| Type | Desktop App |
| Stack | TBD (Python / Electron / C#) |
| Repo | github.com/Kepners/dwg-auto-printer |
| Distribution | Standalone EXE |
| Monetization | TBD |

---

## Key Documentation

| Doc | Purpose |
|-----|---------|
| [docs/SPEC.md](docs/SPEC.md) | Feature spec and roadmap |
| [ARCHITECTURE.md](ARCHITECTURE.md) | System design |

---

## Design System: Engineering Blue + AutoCAD Red

| Name | Hex | Usage |
|------|-----|-------|
| Deep Navy | `#0A1628` | App background |
| Panel | `#0F1E35` | Panel surfaces |
| Card | `#152440` | Card backgrounds |
| Engineering Blue | `#1565C0` | Primary accent, CTAs |
| Blue Hover | `#1976D2` | Button hover state |
| AutoCAD Red | `#E2231A` | Secondary accent, warnings, active state |
| Red Hover | `#C41E3A` | Red hover |
| Text | `#E8EFF8` | Near-white body text |
| Subtext | `#5A7A9A` | Secondary text |
| Border | `#1E3050` | Borders / separators |

```css
:root {
  --bg: #0A1628;
  --blue: #1565C0;
  --red: #E2231A;
  --text: #E8EFF8;
  --border: #1E3050;
}
```

---

## Problem Statement
Printing multiple DWG files in AutoCAD is painful:
- Must open each file manually
- Plot settings reset per file
- No batch queue
- Hard to print to PDF + physical printer simultaneously

DWGAutoPrinter solves this: drop DWG files in, configure once, print all.

---

## MCP Note
> Per user preference (2026-03-04): configure all MCP servers using Docker going forward.

---

*Created: 2026-03-04*
