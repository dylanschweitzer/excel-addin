using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
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

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("CopyForLLM.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnCopyForLLM(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.CopySelectionForLLM();
        }

        public async void OnCheckForUpdates(Office.IRibbonControl control)
        {
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                var result = await VersionChecker.CheckForUpdatesAsync();

                if (!result.Success)
                {
                    MessageBox.Show(
                        $"Could not check for updates.\n\n{result.ErrorMessage}",
                        "Update Check Failed",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                if (result.UpdateAvailable)
                {
                    var dialogResult = MessageBox.Show(
                        $"A newer version is available!\n\n" +
                        $"Current version: {result.CurrentVersion}\n" +
                        $"Latest version: {result.LatestVersion}\n\n" +
                        $"Would you like to open the download page?",
                        "Update Available",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Information);

                    if (dialogResult == DialogResult.Yes && !string.IsNullOrEmpty(result.ReleaseUrl))
                    {
                        Process.Start(result.ReleaseUrl);
                    }
                }
                else
                {
                    MessageBox.Show(
                        $"You are running the latest version ({result.CurrentVersion}).",
                        "No Updates Available",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
        }

        #endregion

        #region Helpers

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

        #endregion
    }
}
