# Excel Add-in: Copy Cell for LLM - Build Log

**Last Updated:** 2026-01-20

## Project Overview

Excel add-in that adds a ribbon button to copy selected cell(s) reference and contents in a format suitable for pasting into LLMs.

**Output format:**
- Single cell: `A2: value`
- Multiple cells: Each cell on its own line

## Completed

- [x] Initialize git repository
- [x] Create BUILD_LOG.md
- [x] Create manifest.xml with ribbon button configuration
- [x] Create commands.html entry point
- [x] Create commands.js with copy functionality
- [x] Create icon assets (16x16, 32x32, 64x64, 80x80 PNG)
- [x] Restructure for GitHub Pages static hosting
- [x] Create README.md with user and admin instructions
- [x] Remove Node.js dev server dependency

## In Progress

(none)

## Not Started

- [ ] Push to GitHub and enable GitHub Pages
- [ ] Replace REPLACE_WITH_YOUR_URL in manifest.xml with actual GitHub Pages URL (7 places)
- [ ] Test in Excel Desktop
- [ ] Test in Excel Online
- [ ] Consider adding keyboard shortcut

## Resume Notes

All code is complete. Next steps:
1. Create GitHub repo (public)
2. Push this folder to GitHub
3. Enable GitHub Pages (Settings → Pages → main branch → root)
4. Wait for GitHub Pages URL (https://USERNAME.github.io/REPONAME)
5. Edit manifest.xml: find/replace `REPLACE_WITH_YOUR_URL` with your URL
6. Commit and push the updated manifest
7. Test by loading manifest.xml in Excel

## Project Structure

```
excel-addin/
├── manifest.xml      # Add-in manifest (edit REPLACE_WITH_YOUR_URL)
├── commands.html     # Entry point for ribbon commands
├── commands.js       # Copy functionality
├── README.md         # User and admin instructions
├── BUILD_LOG.md
├── .gitignore
├── icon.svg          # Source icon
└── assets/
    ├── icon-16.png
    ├── icon-32.png
    ├── icon-64.png
    └── icon-80.png
```

## Deployment

This add-in uses GitHub Pages for hosting (no server required).

1. Push to GitHub
2. Enable GitHub Pages (Settings → Pages → main branch)
3. Replace `REPLACE_WITH_YOUR_URL` in manifest.xml with your GitHub Pages URL
4. Distribute manifest.xml to users

## Notes

### How It Works
1. User selects cell(s) in Excel
2. Clicks "Copy for LLM" button in the "LLM Copy" ribbon tab
3. Add-in reads selected range address and values
4. Formats as `CellRef: value` (one per line for multiple cells)
5. Copies to clipboard via `navigator.clipboard.writeText()`

### User Experience
- Users only need to: download manifest.xml → load in Excel → use button
- No Node.js, no command line, no local server
