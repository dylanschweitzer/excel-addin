# Excel Add-in v1.1 - Build Log

**Last Updated:** 2026-01-21

## Overview

Version 1.1 is the Excel-DNA edition of the CopyForLLM add-in. It replaces the VSTO version with a modern, easy-to-build architecture.

**Features:**
- Copy Values, Copy Formulas, and Copy Both buttons for LLM-friendly cell copying
- Prepare to Share button for resetting workbook view state before sharing
- Settings button with configurable Ctrl+; keyboard shortcut for blue/black font toggle
- Check for Updates button (GitHub Releases API)

---

## Quick Start

### Prerequisites
- .NET 8 SDK or later: https://dotnet.microsoft.com/download
- Windows (Excel-DNA only works with Windows desktop Excel)

### Build
```bash
cd /path/to/excel-addin-v1.1
dotnet restore
dotnet build
```

### Output
After building, find the add-in at:
```
bin/Debug/net8.0-windows/CopyForLLM-AddIn64.xll
```
(Use `CopyForLLM-AddIn.xll` for 32-bit Excel)

### Install/Test
1. Open Excel
2. File → Options → Add-ins
3. Manage: Excel Add-ins → Go
4. Browse → select the `.xll` file
5. The "Copy for LLM" tab appears in the ribbon

---

## Features

### Copy Values Button
Copies selected cells with their calculated values:
```
A1: 42
A2: Hello
B1: 100
```

### Copy Formulas Button
Copies selected cells with their formulas:
```
A1: =SUM(B1:B10)
A2: Hello
B1: =A1*2
```
- If a cell contains a plain value (no formula), the value is shown
- Empty cells show `[empty]`

### Copy Both Button
Copies selected cells with formulas and their calculated values:
```
A1: =SUM(B1:B10) → 150
A2: Hello
B1: =A1*2 → 300
```
- Cells with formulas show: `formula → value`
- Cells without formulas show just the value

### Prepare to Share Button
Opens a dialog with options to clean up the workbook before sharing:
- **Set active cell to A1 on each sheet** - Resets cursor position on all worksheets
- **Set zoom to 100% on each sheet** - Resets zoom level on all worksheets

This ensures recipients see a clean, consistent view when they open the file.

### Settings Button
Opens a dialog to configure add-in options:
- **Enable Ctrl+; to toggle font color (blue/black)** - When enabled, pressing Ctrl+; toggles the selected cells' font color between blue and black (disabled by default)

### Keyboard Shortcuts
| Shortcut | Action | Default | Configurable |
|----------|--------|---------|--------------|
| Ctrl+; | Toggle font color blue/black | Off | Yes (in Settings) |

### Check for Updates Button
- Queries GitHub Releases API
- Compares current version to latest release
- Offers to open download page if update available

---

## Project Structure

| File | Purpose |
|------|---------|
| `CopyForLLM.csproj` | Project file with Excel-DNA packages |
| `Ribbon1.cs` | Ribbon UI + button handlers + copy logic |
| `VersionChecker.cs` | GitHub API update checker |
| `BUILD_LOG.md` | This file |
| `COMPILE_INSTRUCTIONS.md` | Step-by-step Windows build guide |

---

## Session Log: 2026-01-21

### Completed
- [x] Excel-DNA migration from VSTO (no Visual Studio required)
- [x] Copy Values, Copy Formulas, Copy Both buttons
- [x] Prepare to Share button with A1 reset and zoom reset
- [x] Settings button with Ctrl+; shortcut toggle
- [x] Blue-black font toggle (disabled by default)
- [x] Fixed Prepare to Share error with hidden sheets
- [x] Build tested on Windows
- [x] Create GitHub release v1.1.0
- [x] Update main README to focus on v1.1

---

## Version History

- **v1.1.0** - Excel-DNA edition with all features
- **v1.0.0** - Original VSTO version
