# How to Compile CopyForLLM Excel Add-in on Windows

Step-by-step instructions to build the `.xll` file from source.

---

## Step 1: Install .NET 8 SDK

1. Open your browser and go to: https://dotnet.microsoft.com/download/dotnet/8.0
2. Under ".NET SDK", click the **Windows x64** installer
3. Run the downloaded installer and follow the prompts
4. When finished, open a new Command Prompt or PowerShell window
5. Verify installation by typing:
   ```
   dotnet --version
   ```
   You should see `8.x.x` or higher

---

## Step 2: Open Command Prompt in Project Folder

1. Open File Explorer
2. Navigate to the `excel-addin-v1.1` folder
3. Click in the address bar and type `cmd` then press Enter
   - This opens Command Prompt in the current folder
   - Alternatively: hold Shift + right-click in the folder → "Open PowerShell window here"

---

## Step 3: Restore Dependencies

In the Command Prompt, type:
```
dotnet restore
```

Wait for it to complete. You should see "Restore completed" with no errors.

---

## Step 4: Build the Add-in

For a **Debug build** (for testing):
```
dotnet build
```

For a **Release build** (for distribution):
```
dotnet build -c Release
```

Wait for the build to complete. You should see "Build succeeded" with 0 errors.

---

## Step 5: Locate the .xll File

After building, find your add-in file at:

**Debug build:**
```
bin\Debug\net8.0-windows\CopyForLLM-AddIn64.xll
```

**Release build:**
```
bin\Release\net8.0-windows\CopyForLLM-AddIn64.xll
```

> Note: Use `CopyForLLM-AddIn.xll` (without "64") if you have 32-bit Excel.

---

## Step 6: Install the Add-in in Excel

1. Open Excel
2. Go to **File** → **Options**
3. Click **Add-ins** in the left sidebar
4. At the bottom, next to "Manage:", select **Excel Add-ins** and click **Go...**
5. Click **Browse...**
6. Navigate to the `.xll` file from Step 5 and select it
7. Click **OK**
8. The "Copy for LLM" tab should now appear in your Excel ribbon

---

## Troubleshooting

### "dotnet" is not recognized
- Make sure you installed the .NET 8 SDK (not just the Runtime)
- Close and reopen Command Prompt after installing
- Try restarting your computer

### Build errors about missing packages
- Run `dotnet restore` again
- Make sure you have internet access (NuGet packages need to download)

### Add-in doesn't load in Excel
- Make sure you're using the correct bitness (64-bit .xll for 64-bit Excel)
- Right-click the .xll file → Properties → check "Unblock" if present → OK
- Check Excel Trust Center settings: File → Options → Trust Center → Trust Center Settings → Add-ins

### "Could not load file or assembly" error
- Make sure .NET 8 Desktop Runtime is installed
- Download from: https://dotnet.microsoft.com/download/dotnet/8.0 (look for ".NET Desktop Runtime")

---

## Quick Reference Commands

```
# Restore packages
dotnet restore

# Build debug version
dotnet build

# Build release version
dotnet build -c Release

# Clean and rebuild
dotnet clean && dotnet build
```
