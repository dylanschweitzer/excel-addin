# Excel Add-in: Copy Cell for LLM - Build Log

**Last Updated:** 2026-01-21

## Project Overview

Excel add-in that adds a ribbon button to copy selected cell(s) reference and contents in a format suitable for pasting into LLMs.

**Output format:**
- Single cell: `A2: value`
- Multiple cells: Each cell on its own line
- Empty cells: `A2: [empty]`

## Two Implementations

This project contains two implementations:

1. **Web Add-in** (manifest.xml) - JavaScript-based, hosted on GitHub Pages
   - Located in root folder
   - Requires sideloading which is blocked on many Excel versions
   - **Not recommended** due to installation difficulties

2. **COM Add-in** (VSTO) - C#-based, compiled installer
   - Located in `com-addin/` folder
   - Easy installation via setup.exe
   - **Recommended approach**

## Completed

### Initial Web Add-in (2026-01-19)
- [x] Initialize git repository
- [x] Create BUILD_LOG.md
- [x] Create manifest.xml with ribbon button configuration
- [x] Create commands.html entry point
- [x] Create commands.js with copy functionality
- [x] Create icon assets (16x16, 32x32, 64x64, 80x80 PNG)
- [x] Restructure for GitHub Pages static hosting
- [x] Create README.md with user and admin instructions
- [x] Push to GitHub (https://github.com/dylanschweitzer/excel-addin)
- [x] Enable GitHub Pages
- [x] Replace REPLACE_WITH_YOUR_URL in manifest.xml with actual GitHub Pages URL

### COM Add-in Development (2026-01-19)
- [x] Discover web add-in sideloading is blocked on consumer Excel
- [x] Research alternatives (VBA .xlam, COM add-in)
- [x] Create VSTO COM add-in project structure
- [x] Write ThisAddIn.cs with copy functionality
- [x] Write Ribbon1.cs for ribbon extensibility
- [x] Write Ribbon1.xml for ribbon UI definition
- [x] Build and test in Visual Studio
- [x] Create ClickOnce installer (setup.exe)
- [x] Document Visual Studio build process

### GitHub Release (2026-01-20)
- [x] Package installer files (setup.exe, CopyForLLM.vsto, Application Files/) into zip
- [x] Create GitHub Release v1.0.0
- [x] Upload installer zip as release asset
- [x] Update README.md with direct download link

### Check for Updates Feature (2026-01-20)
- [x] Add System.Net.Http and System.Web.Extensions references to csproj
- [x] Create VersionChecker.cs service class (GitHub API integration)
- [x] Add "Check for Updates" button to ribbon in new "About" group
- [x] Implement OnCheckForUpdates async handler in Ribbon1.cs

### Excel-DNA v1.1 Release (2026-01-21)
- [x] Create Excel-DNA version (no Visual Studio required, builds with dotnet CLI)
- [x] Copy Values, Copy Formulas, Copy Both buttons
- [x] Prepare to Share button (reset A1, zoom 100% on all sheets)
- [x] Settings button with Ctrl+; blue/black font toggle (disabled by default)
- [x] Fixed Prepare to Share error with hidden sheets (use Goto instead of Select)
- [x] Build and test on Windows
- [x] Push excel-dna/ folder to GitHub
- [x] Create GitHub Release v1.1.0 with .xll files
- [x] Update README to focus on v1.1 Excel-DNA version
- [x] Archive v1.0 VSTO version as "Previous Versions"

## Project Structure

```
excel-addin/
├── BUILD_LOG.md
├── README.md
├── VISUAL_STUDIO_GUIDE.md
├── .gitignore
│
├── excel-dna/                # v1.1 Excel-DNA (recommended)
│   ├── CopyForLLM.csproj
│   ├── Ribbon1.cs
│   ├── VersionChecker.cs
│   ├── BUILD_LOG.md
│   ├── COMPILE_INSTRUCTIONS.md
│   └── release/              # Pre-built .xll files
│       ├── CopyForLLM-AddIn64.xll
│       └── CopyForLLM-AddIn.xll
│
├── com-addin/                # v1.0 VSTO (archived)
│   ├── CopyForLLM.sln
│   └── CopyForLLM/
│       ├── CopyForLLM.csproj
│       ├── ThisAddIn.cs
│       ├── Ribbon1.cs
│       ├── Ribbon1.xml
│       └── VersionChecker.cs
│
├── # Web Add-in (not recommended)
├── manifest.xml
├── commands.html
├── commands.js
└── assets/
```

## How It Works

1. User selects cell(s) in Excel
2. Clicks "Copy for LLM" button in the "LLM Copy" ribbon tab
3. Add-in reads selected range addresses and values
4. Formats as `CellRef: value` (one per line for multiple cells)
5. Copies to clipboard

## Deployment

### COM Add-in (Recommended)
1. Open `com-addin/CopyForLLM.sln` in Visual Studio
2. Build → Publish to create setup.exe
3. Distribute setup.exe to users
4. Users run setup.exe to install

### Web Add-in (Limited Use)
- Hosted at: https://dylanschweitzer.github.io/excel-addin
- Manifest: https://dylanschweitzer.github.io/excel-addin/manifest.xml
- Only works if Excel allows sideloading (enterprise/developer setups)

## Notes

### Why COM Add-in?
The original web add-in approach failed because:
- Modern consumer Excel versions block sideloading of web add-ins
- "Upload My Add-in" option not available without enterprise/developer settings
- Trusted Add-in Catalogs method is complex and unreliable

The COM add-in approach works because:
- Installs via standard Windows installer (setup.exe)
- No special Excel settings required
- Ribbon button appears automatically after installation

### External Dependencies
- **Web add-in**: Loads Microsoft's office.js from https://appsforoffice.microsoft.com
- **COM add-in**: Requires VSTO Runtime (included in installer)
