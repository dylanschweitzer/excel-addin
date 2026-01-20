# Visual Studio Build Guide: Copy Cell for LLM

Step-by-step instructions for building the COM add-in from scratch using Visual Studio.

---

## Prerequisites

1. **Windows PC**
2. **Visual Studio 2022** (Community edition is free)
   - Download: https://visualstudio.microsoft.com/
3. **Microsoft Excel** installed

---

## Step 1: Install Visual Studio with Office Development

1. Run the Visual Studio installer
2. Click **Modify** (if already installed) or continue with fresh install
3. In the **Workloads** tab, check:
   - **Office/SharePoint development**
   - **.NET desktop development**
4. Click **Install** or **Modify**
5. Wait for installation to complete

---

## Step 2: Create a New VSTO Project

1. Open Visual Studio
2. Click **Create a new project**
3. In the search box, type **Excel VSTO**
4. Select **Excel VSTO Add-in** (make sure it says C#)
5. Click **Next**
6. Configure:
   - Project name: `CopyForLLM`
   - Location: wherever you want (e.g., `C:\Users\YourName\source\repos\`)
   - Solution name: `CopyForLLM`
7. Click **Create**
8. If asked for target framework, select **.NET Framework 4.8** (or 4.7.2)

Visual Studio creates the project with default template files.

---

## Step 3: Edit ThisAddIn.cs

1. In **Solution Explorer** (right panel), double-click **ThisAddIn.cs**
2. Select all the code (Ctrl+A)
3. Delete it
4. Paste this code:

```csharp
using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace CopyForLLM
{
    public partial class ThisAddIn
    {
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        public void CopySelectionForLLM()
        {
            try
            {
                Excel.Range selection = Application.Selection as Excel.Range;
                if (selection == null)
                {
                    MessageBox.Show("Please select some cells first.", "Copy for LLM",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var lines = new System.Text.StringBuilder();

                foreach (Excel.Range cell in selection.Cells)
                {
                    string address = cell.Address[false, false];
                    object val = cell.Value2;
                    string value = val == null ? "[empty]" : val.ToString();

                    if (lines.Length > 0) lines.AppendLine();
                    lines.Append($"{address}: {value}");
                }

                if (lines.Length > 0)
                {
                    Clipboard.SetText(lines.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error copying cells: " + ex.Message, "Copy for LLM",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
```

5. Save (Ctrl+S)

---

## Step 4: Add a Ribbon

1. In **Solution Explorer**, right-click the **CopyForLLM** project (not the solution)
2. Click **Add** → **New Item...**
3. In the search box, type **Ribbon**
4. Select **Ribbon (XML)**
5. Name it `Ribbon1.cs`
6. Click **Add**

Visual Studio creates two files: Ribbon1.cs and Ribbon1.xml

---

## Step 5: Edit Ribbon1.cs

1. Double-click **Ribbon1.cs** in Solution Explorer
2. Select all (Ctrl+A) and delete
3. Paste this code:

```csharp
using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace CopyForLLM
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon1()
        {
        }

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("CopyForLLM.Ribbon1.xml");
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnCopyForLLM(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.CopySelectionForLLM();
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            foreach (string name in resourceNames)
            {
                if (string.Compare(resourceName, name, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(name)))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }
    }
}
```

4. Save (Ctrl+S)

---

## Step 6: Edit Ribbon1.xml

1. In Solution Explorer, expand the arrow next to **Ribbon1.cs**
2. Double-click **Ribbon1.xml**
3. Select all (Ctrl+A) and delete
4. Paste this XML:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="Ribbon_Load">
  <ribbon>
    <tabs>
      <tab id="LLMCopyTab" label="LLM Copy">
        <group id="CopyGroup" label="Cell Tools">
          <button id="CopyForLLMButton"
                  label="Copy for LLM"
                  size="large"
                  onAction="OnCopyForLLM"
                  imageMso="Copy"
                  screentip="Copy for LLM"
                  supertip="Copy selected cell(s) reference and contents to clipboard in a format suitable for pasting into LLMs like ChatGPT or Claude."/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
```

5. Save (Ctrl+S)

---

## Step 7: Verify Ribbon1.xml Build Action

1. Click on **Ribbon1.xml** in Solution Explorer
2. Press **F4** to open Properties panel
3. Find **Build Action**
4. Make sure it's set to **Embedded Resource**
   - If not, click the dropdown and select **Embedded Resource**

---

## Step 8: Build the Project

1. Press **Ctrl+Shift+B** (or Build → Build Solution)
2. Look at the **Output** window at the bottom
3. It should say **Build succeeded**

If you get errors:
- Make sure all the code was copied correctly
- Check that namespaces say `CopyForLLM` (not something else)
- Verify Ribbon1.xml has the XML content (not C# code)

---

## Step 9: Test in Debug Mode

1. Press **F5** (or Debug → Start Debugging)
2. Excel opens automatically with the add-in loaded
3. Look for the **LLM Copy** tab in the ribbon
4. Select some cells
5. Click **Copy for LLM**
6. Open Notepad and paste (Ctrl+V) to verify

To stop debugging, close Excel or click the red Stop button in Visual Studio.

---

## Step 10: Create the Installer

1. In the toolbar, change **Debug** to **Release** (dropdown near the top)
2. Click **Build** → **Publish CopyForLLM**
3. Choose a folder location (e.g., `C:\Users\YourName\Desktop\CopyForLLM-Installer`)
4. Click **Next**
5. Select **From a CD-ROM or DVD-ROM**
6. Click **Next**
7. Select **The application will not check for updates**
8. Click **Next** then **Finish** or **Publish**

Visual Studio builds the installer.

---

## Step 11: Distribute

The installer folder contains:
- `setup.exe` - the installer
- `CopyForLLM.vsto` - the add-in package
- `Application Files/` - supporting files

**To distribute:**
1. Zip the entire folder
2. Send to users
3. Users unzip and run `setup.exe`

---

## Managing the Add-in After Installation

### Enable/Disable
1. Excel → File → Options → Add-ins
2. At bottom: Manage: **COM Add-ins** → Go
3. Check/uncheck **CopyForLLM**

### Uninstall
Windows Settings → Apps → search "CopyForLLM" → Uninstall

---

## Common Issues

### "Ribbon tab doesn't appear"
- Restart Excel completely
- Check COM Add-ins is enabled (see above)
- Rebuild and try again

### "Build errors about missing references"
- Make sure Visual Studio has the Office development workload installed
- Try: right-click solution → Restore NuGet Packages

### "The type or namespace 'Ribbon1' could not be found"
- Make sure the namespace in Ribbon1.cs is `CopyForLLM` (matching the project)
- Check that you saved all files

### "Invalid expression term '<'"
- You probably pasted XML into a .cs file or vice versa
- Ribbon1.cs should have C# code (starts with `using`)
- Ribbon1.xml should have XML (starts with `<?xml`)

---

## Files Reference

After completing all steps, your project should have:

| File | Contains |
|------|----------|
| ThisAddIn.cs | Main add-in code with `CopySelectionForLLM()` method |
| Ribbon1.cs | Ribbon handler with `OnCopyForLLM()` callback |
| Ribbon1.xml | Ribbon UI definition (embedded resource) |
