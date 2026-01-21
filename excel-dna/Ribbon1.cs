using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CopyForLLM
{
    // Static settings manager
    public static class AddinSettings
    {
        public static bool BlueBlackToggleEnabled { get; set; } = false;

        public static void RegisterShortcuts()
        {
            if (BlueBlackToggleEnabled)
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                // Ctrl+; is "^;" in Excel OnKey syntax
                app.OnKey("^;", "ToggleBlueBlack");
            }
        }

        public static void UnregisterShortcuts()
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                // Passing empty string removes the shortcut binding
                app.OnKey("^;", "");
            }
            catch { }
        }

        public static void UpdateShortcutRegistration()
        {
            if (BlueBlackToggleEnabled)
            {
                RegisterShortcuts();
            }
            else
            {
                UnregisterShortcuts();
            }
        }
    }

    // Excel-DNA command for the blue-black toggle (called via OnKey)
    public static class Commands
    {
        [ExcelCommand(Name = "ToggleBlueBlack")]
        public static void ToggleBlueBlack()
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                var selection = app.Selection as Excel.Range;

                if (selection == null)
                    return;

                // Get the current font color of the first cell
                var currentColor = selection.Font.Color;

                // Define blue and black colors
                int blueColor = ColorTranslator.ToOle(Color.Blue);
                int blackColor = ColorTranslator.ToOle(Color.Black);

                // Toggle: if currently blue, make black; otherwise make blue
                if (currentColor != null && Convert.ToInt32(currentColor) == blueColor)
                {
                    selection.Font.Color = blackColor;
                }
                else
                {
                    selection.Font.Color = blueColor;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error toggling color: " + ex.Message, "Blue-Black Toggle",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    [ComVisible(true)]
    public class Ribbon1 : ExcelRibbon
    {
        public override string GetCustomUI(string ribbonID)
        {
            return @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='Ribbon_Load'>
  <ribbon>
    <tabs>
      <tab id='CopyForLLMTab' label='Copy for LLM'>
        <group id='CellToolsGroup' label='Cell Tools'>
          <button id='CopyValuesButton'
                  label='Copy Values'
                  size='large'
                  onAction='OnCopyValues'
                  imageMso='Copy'
                  screentip='Copy Cell Values for LLM'
                  supertip='Copies selected cells with their values in a format suitable for pasting into an LLM chat (e.g., A1: 42)' />
          <button id='CopyFormulasButton'
                  label='Copy Formulas'
                  size='large'
                  onAction='OnCopyFormulas'
                  imageMso='ShowFormulas'
                  screentip='Copy Cell Formulas for LLM'
                  supertip='Copies selected cells with their formulas in a format suitable for pasting into an LLM chat (e.g., A1: =SUM(B1:B10))' />
          <button id='CopyBothButton'
                  label='Copy Both'
                  size='large'
                  onAction='OnCopyBoth'
                  imageMso='CopyToFolder'
                  screentip='Copy Formulas and Values for LLM'
                  supertip='Copies selected cells with both formulas and their calculated values (e.g., A1: =SUM(B1:B10) → 150)' />
        </group>
        <group id='WorkbookToolsGroup' label='Workbook Tools'>
          <button id='PrepareToShareButton'
                  label='Prepare to Share'
                  size='large'
                  onAction='OnPrepareToShare'
                  imageMso='FileSendAsAttachment'
                  screentip='Prepare Workbook for Sharing'
                  supertip='Reset all sheets to A1 and/or set zoom to 100% before sharing' />
        </group>
        <group id='SettingsGroup' label='Settings'>
          <button id='SettingsButton'
                  label='Settings'
                  size='large'
                  onAction='OnSettings'
                  imageMso='ControlsGallery'
                  screentip='Add-in Settings'
                  supertip='Configure keyboard shortcuts and other options' />
          <button id='CheckForUpdatesButton'
                  label='Check for Updates'
                  size='large'
                  onAction='OnCheckForUpdates'
                  imageMso='WebPagePreview'
                  screentip='Check for Updates'
                  supertip='Check GitHub for newer versions of this add-in' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        private IRibbonUI ribbon;

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            // Register keyboard shortcuts based on settings
            AddinSettings.RegisterShortcuts();
        }

        public void OnCopyValues(IRibbonControl control)
        {
            CopySelectionForLLM(copyFormulas: false);
        }

        public void OnCopyFormulas(IRibbonControl control)
        {
            CopySelectionForLLM(copyFormulas: true);
        }

        public void OnCopyBoth(IRibbonControl control)
        {
            CopySelectionWithBoth();
        }

        public void OnPrepareToShare(IRibbonControl control)
        {
            using (var dialog = new PrepareToShareDialog())
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    PrepareWorkbookForSharing(dialog.SetCellToA1, dialog.SetZoomTo100);
                }
            }
        }

        public void OnSettings(IRibbonControl control)
        {
            using (var dialog = new SettingsDialog())
            {
                dialog.ShowDialog();
            }
        }

        public async void OnCheckForUpdates(IRibbonControl control)
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
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = result.ReleaseUrl,
                            UseShellExecute = true
                        });
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

        private void CopySelectionForLLM(bool copyFormulas)
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                var selection = app.Selection as Excel.Range;

                if (selection == null)
                {
                    MessageBox.Show("Please select some cells first.", "Copy for LLM",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var lines = new StringBuilder();

                foreach (Excel.Range cell in selection.Cells)
                {
                    string address = cell.Address[false, false];
                    string content;

                    if (copyFormulas)
                    {
                        // Get the formula - if the cell has no formula, Formula returns the value as string
                        object formula = cell.Formula;
                        if (formula == null)
                        {
                            content = "[empty]";
                        }
                        else
                        {
                            string formulaStr = formula.ToString();
                            // If it doesn't start with '=', it's just a value, not a formula
                            content = string.IsNullOrEmpty(formulaStr) ? "[empty]" : formulaStr;
                        }
                    }
                    else
                    {
                        // Get the calculated value
                        object val = cell.Value2;
                        content = val == null ? "[empty]" : val.ToString();
                    }

                    if (lines.Length > 0) lines.AppendLine();
                    lines.Append($"{address}: {content}");
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

        private void CopySelectionWithBoth()
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                var selection = app.Selection as Excel.Range;

                if (selection == null)
                {
                    MessageBox.Show("Please select some cells first.", "Copy for LLM",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var lines = new StringBuilder();

                foreach (Excel.Range cell in selection.Cells)
                {
                    string address = cell.Address[false, false];

                    object formula = cell.Formula;
                    object val = cell.Value2;

                    string formulaStr = formula?.ToString() ?? "";
                    string valueStr = val == null ? "[empty]" : val.ToString();

                    string content;

                    // Check if cell has a formula (starts with '=')
                    if (!string.IsNullOrEmpty(formulaStr) && formulaStr.StartsWith("="))
                    {
                        // Show both: formula → value
                        content = $"{formulaStr} → {valueStr}";
                    }
                    else
                    {
                        // No formula, just show the value
                        content = valueStr;
                    }

                    if (lines.Length > 0) lines.AppendLine();
                    lines.Append($"{address}: {content}");
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

        private void PrepareWorkbookForSharing(bool setCellToA1, bool setZoomTo100)
        {
            if (!setCellToA1 && !setZoomTo100)
                return;

            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                var workbook = app.ActiveWorkbook;

                if (workbook == null)
                {
                    MessageBox.Show("No workbook is open.", "Prepare to Share",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Remember the currently active sheet to restore it later
                var originalSheet = app.ActiveSheet as Excel.Worksheet;

                // Disable screen updating for performance
                app.ScreenUpdating = false;

                try
                {
                    foreach (Excel.Worksheet sheet in workbook.Worksheets)
                    {
                        // Skip hidden sheets
                        if (sheet.Visible != Excel.XlSheetVisibility.xlSheetVisible)
                            continue;

                        sheet.Activate();

                        if (setCellToA1)
                        {
                            // Use Goto instead of Select - more reliable
                            app.Goto(sheet.Range["A1"], true);
                        }

                        if (setZoomTo100)
                        {
                            app.ActiveWindow.Zoom = 100;
                        }
                    }

                    // Restore the original active sheet
                    if (originalSheet != null && originalSheet.Visible == Excel.XlSheetVisibility.xlSheetVisible)
                    {
                        originalSheet.Activate();
                        if (setCellToA1)
                        {
                            app.Goto(originalSheet.Range["A1"], true);
                        }
                    }
                }
                finally
                {
                    app.ScreenUpdating = true;
                }

                MessageBox.Show("Workbook prepared for sharing.", "Prepare to Share",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error preparing workbook: " + ex.Message, "Prepare to Share",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    public class PrepareToShareDialog : Form
    {
        public bool SetCellToA1 { get; private set; }
        public bool SetZoomTo100 { get; private set; }

        private CheckBox chkSetCellToA1;
        private CheckBox chkSetZoomTo100;

        public PrepareToShareDialog()
        {
            Text = "Prepare to Share";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            Width = 300;
            Height = 180;

            chkSetCellToA1 = new CheckBox
            {
                Text = "Set active cell to A1 on each sheet",
                Location = new System.Drawing.Point(20, 20),
                Width = 250,
                Checked = true
            };

            chkSetZoomTo100 = new CheckBox
            {
                Text = "Set zoom to 100% on each sheet",
                Location = new System.Drawing.Point(20, 50),
                Width = 250,
                Checked = true
            };

            var btnOK = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Location = new System.Drawing.Point(100, 100),
                Width = 80
            };
            btnOK.Click += (s, e) =>
            {
                SetCellToA1 = chkSetCellToA1.Checked;
                SetZoomTo100 = chkSetZoomTo100.Checked;
            };

            var btnCancel = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Location = new System.Drawing.Point(190, 100),
                Width = 80
            };

            Controls.Add(chkSetCellToA1);
            Controls.Add(chkSetZoomTo100);
            Controls.Add(btnOK);
            Controls.Add(btnCancel);

            AcceptButton = btnOK;
            CancelButton = btnCancel;
        }
    }

    public class SettingsDialog : Form
    {
        private CheckBox chkBlueBlackToggle;

        public SettingsDialog()
        {
            Text = "Settings";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            Width = 350;
            Height = 180;

            var lblShortcuts = new Label
            {
                Text = "Keyboard Shortcuts",
                Location = new Point(20, 15),
                Width = 300,
                Font = new Font(Font, FontStyle.Bold)
            };

            chkBlueBlackToggle = new CheckBox
            {
                Text = "Enable Ctrl+; to toggle font color (blue/black)",
                Location = new Point(20, 45),
                Width = 300,
                Checked = AddinSettings.BlueBlackToggleEnabled
            };

            var btnOK = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Location = new Point(150, 100),
                Width = 80
            };
            btnOK.Click += (s, e) =>
            {
                AddinSettings.BlueBlackToggleEnabled = chkBlueBlackToggle.Checked;
                AddinSettings.UpdateShortcutRegistration();
            };

            var btnCancel = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Location = new Point(240, 100),
                Width = 80
            };

            Controls.Add(lblShortcuts);
            Controls.Add(chkBlueBlackToggle);
            Controls.Add(btnOK);
            Controls.Add(btnCancel);

            AcceptButton = btnOK;
            CancelButton = btnCancel;
        }
    }
}
