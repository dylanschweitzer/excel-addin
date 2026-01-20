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
