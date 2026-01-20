# Copy Cell for LLM - Excel Add-in

Copy selected Excel cells with their references, formatted for pasting into ChatGPT, Claude, or other LLMs.

**Example output:**
```
A2: Sales Data
B2: 1500
C2: 2024-01-15
```

---

## Installation (Recommended: COM Add-in)

### For End Users

1. **[Download the installer](https://github.com/dylanschweitzer/excel-addin/releases/latest/download/CopyForLLM-v1.0.0-installer.zip)** (or visit [Releases](https://github.com/dylanschweitzer/excel-addin/releases))
2. Extract the zip file
3. Run **setup.exe**
4. Click **Install** when prompted
5. Open Excel - the **LLM Copy** tab appears in the ribbon

That's it!

### To Uninstall

Windows Settings → Apps → search "CopyForLLM" → Uninstall

---

## Usage

1. Select any cell or range (e.g., A1:D10)
2. Click the **LLM Copy** tab in the ribbon
3. Click **Copy for LLM**
4. Paste (Ctrl+V) into your LLM chat

---

## For Developers / Building from Source

See [VISUAL_STUDIO_GUIDE.md](VISUAL_STUDIO_GUIDE.md) for detailed instructions on building the COM add-in with Visual Studio.

### Quick Summary

1. Install Visual Studio with **Office/SharePoint development** workload
2. Open `com-addin/CopyForLLM.sln`
3. Build → Publish CopyForLLM
4. Distribute the generated setup.exe

---

## Output Format

| Selection | Output |
|-----------|--------|
| Single cell A2 | `A2: Hello` |
| Empty cell | `A2: [empty]` |
| Range A2:B3 | `A2: Hello`<br>`B2: World`<br>`A3: Foo`<br>`B3: Bar` |

---

## Alternative: Web Add-in (Limited)

This repo also contains a web-based add-in (manifest.xml), but it's **not recommended** because:
- Most consumer Excel versions block sideloading
- Requires "Upload My Add-in" option which isn't available in standard Excel

If you have an enterprise or developer Excel setup that allows sideloading:
- Manifest URL: https://dylanschweitzer.github.io/excel-addin/manifest.xml

---

## Troubleshooting

**Ribbon tab doesn't appear after install**
- Close and reopen Excel
- Check File → Options → Add-ins → COM Add-ins → Go → ensure CopyForLLM is checked

**Install fails**
- You may need the VSTO Runtime: https://aka.ms/VSTORuntime
- Install it, then run setup.exe again

**Button doesn't copy anything**
- Make sure cells are selected before clicking
- Check that clipboard access isn't blocked by security software

---

## Project Structure

```
excel-addin/
├── com-addin/           # COM Add-in source (recommended)
│   └── CopyForLLM/
├── manifest.xml         # Web add-in (limited use)
├── commands.js
├── commands.html
└── assets/
```

---

## License

MIT
