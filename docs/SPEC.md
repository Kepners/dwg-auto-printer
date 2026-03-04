# DWGAutoPrinter — Specification

## Overview
Batch auto-printer for DWG/CAD drawing files. Select files, configure plot settings once, print all. Zero manual clicking per file.

## Status
**Phase 0** — Planning / Research

---

## Core Problem
AutoCAD has no native batch print that remembers settings across files. Engineers waste time:
1. Opening each DWG file manually
2. Re-selecting printer and paper size per file
3. Switching between PDF export and physical print
4. Managing drawing revisions and re-prints

---

## MVP Features
- [ ] Batch select DWG files (multi-select, folder scan)
- [ ] Configure plot settings once (printer, paper size, scale, orientation)
- [ ] Print queue with progress indicator
- [ ] Print to physical printer + PDF simultaneously
- [ ] Auto-detect drawing paper size from title block (A1/A2/A3/A4)
- [ ] Remember last-used settings per printer

## Phase 2 Features
- [ ] Watch folder mode (auto-print new files as they appear)
- [ ] Drawing revision detection (skip already-printed revisions)
- [ ] Email confirmation on batch complete
- [ ] Log / audit trail of printed drawings
- [ ] Stamp drawings with print date/time

---

## Technical Approach Options

### Option A: AutoCAD COM Automation (Windows only)
- Uses AutoCAD's COM/ActiveX API
- Requires AutoCAD to be installed
- Most reliable for complex DWG features
- Language: Python (win32com) or C#

### Option B: ODA File Converter + Print
- ODA (Open Design Alliance) free DWG converter
- Can convert DWG → PDF without AutoCAD
- Then print PDF via system printer
- Language: Python + subprocess

### Option C: LibreDWG / ezdxf
- Open source DWG/DXF library
- Good for metadata extraction (title blocks)
- Limited rendering quality
- Language: Python

**Recommended**: Option B for no-AutoCAD-required distribution, Option A for power users with AutoCAD installed.

---

## Target Users
- Structural/civil/mechanical engineers
- Drafters managing large drawing sets
- Print room operators at construction firms
- Anyone printing 10+ DWG files regularly

---

## Design System
Engineering Blue `#1565C0` + AutoCAD Red `#E2231A` on Deep Navy `#0A1628`

---

*Status: DRAFT — Research phase*
*Last Updated: 2026-03-04*
