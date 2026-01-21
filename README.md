# Copy for LLM - Excel Add-in

Copy selected Excel cells with their references, formatted for pasting into ChatGPT, Claude, or other LLMs.

**Example output:**
```
A2: Sales Data
B2: 1500
C2: =A2*0.1 → 150
```

---

## Download

**[Download CopyForLLM-AddIn64.xll](https://github.com/dylanschweitzer/excel-addin/releases/latest/download/CopyForLLM-AddIn64.xll)** (for 64-bit Excel - most common)

[Download CopyForLLM-AddIn.xll](https://github.com/dylanschweitzer/excel-addin/releases/latest/download/CopyForLLM-AddIn.xll) (for 32-bit Excel)

---

## Installation

1. Download the `.xll` file above
2. Right-click the file → Properties → check **Unblock** → OK (if present)
3. Open Excel
4. Go to **File** → **Options** → **Add-ins**
5. At the bottom: Manage: **Excel Add-ins** → click **Go...**
6. Click **Browse...** and select the downloaded `.xll` file
7. Click **OK**
8. The **"Copy for LLM"** tab appears in your ribbon

### Requirements
- Windows with Excel desktop
- [.NET 8 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/8.0) (you may already have it)

---

## Features

### Copy Values
Copies selected cells with their calculated values:
```
A1: 50
A2: Hello
B1: 300
```

### Copy Formulas
Copies selected cells with their formulas:
```
A1: 50
A2: Hello
B1: =A1*6
```

### Copy Both
Copies formulas with their calculated values:
```
A1: 50
A2: Hello
B1: =A1*6 → 300
```

### Prepare to Share
Clean up your workbook before sharing:
- Reset active cell to A1 on all sheets
- Set zoom to 100% on all sheets

### Settings
Configure keyboard shortcuts:
- **Ctrl+;** toggles font color between blue and black (disabled by default)

### Check for Updates
Checks GitHub for newer versions of the add-in.

---

## Usage

1. Select any cells (e.g., A1:D10)
2. Click the **Copy for LLM** tab in the ribbon
3. Click **Copy Values**, **Copy Formulas**, or **Copy Both**
4. Paste (Ctrl+V) into your LLM chat

---

## Building from Source

The add-in uses Excel-DNA and can be built with the .NET CLI (no Visual Studio required).

```bash
cd excel-dna
dotnet restore
dotnet build
```

Output: `bin/Debug/net8.0-windows/CopyForLLM-AddIn64.xll`

See [excel-dna/COMPILE_INSTRUCTIONS.md](excel-dna/COMPILE_INSTRUCTIONS.md) for detailed steps.

---

## Troubleshooting

**Add-in doesn't load**
- Make sure you downloaded the correct version (64-bit vs 32-bit)
- Right-click the .xll file → Properties → Unblock
- Install [.NET 8 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/8.0)

**Ribbon tab doesn't appear**
- Close and reopen Excel
- Check File → Options → Add-ins → Excel Add-ins → ensure CopyForLLM is checked

**"Prepare to Share" shows an error**
- Hidden sheets are now skipped automatically (fixed in v1.1)

---

## Project Structure

```
excel-addin/
├── excel-dna/              # v1.1 Excel-DNA source (recommended)
│   ├── Ribbon1.cs
│   ├── VersionChecker.cs
│   ├── CopyForLLM.csproj
│   └── release/            # Pre-built .xll files
├── com-addin/              # v1.0 VSTO source (archived)
├── manifest.xml            # Web add-in (not recommended)
└── README.md
```

---

## Previous Versions

**v1.0 (VSTO/COM Add-in)** - Requires Visual Studio to build and uses an installer. Still available in the `com-addin/` folder and [v1.0.0 release](https://github.com/dylanschweitzer/excel-addin/releases/tag/v1.0.0) but no longer recommended.

---

## License

MIT
