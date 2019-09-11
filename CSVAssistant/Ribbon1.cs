using Microsoft.Office.Tools.Ribbon;

namespace CSVAssistant
{
    public partial class Ribbon1
    {

        private void SaveButtonAction(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.SaveAsUnicodeCSV(true, false);
        }

        private void SaveAsButtonAction(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.SaveAsUnicodeCSV(true, true);
        }

        private void ExpandAction(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExpandColumn();
        }

        private void CollapseAction(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.CollapseColumn();
        }

        private void FormatAction(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.FormatCell();
        }

        private void IdCheckAction(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.CheckId(true);
        }

        private void CSVCheckAction(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.CSVChecker(false);
        }

        private void CSVCheckAllAction(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.CSVChecker(true);
        }

        private void CSVOpenCheckAction(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.OpenCheck();
        }

        private void OpenImageAction(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.OpenLocalResource();
        }

        private void OpenIOSImageAction(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.OpenI18nResource();
        }

        private void button13_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.SVNCommit();
        }

        private void button14_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.SVNRevert();
        }

        private void button15_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.SVNDiff();
        }

        private void button16_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.SVNLog();
        }

        private void dropDown1_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.SVNRegionDiff(dropDown1.SelectedItem.Tag.ToString());
        }

        private void Button17_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.LoadCellFormat();
        }

        private void Button18_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ClearCellFormat();
        }
    }
}
