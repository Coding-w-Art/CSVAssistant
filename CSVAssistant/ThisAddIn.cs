using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.IO;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace CSVAssistant
{
    public partial class ThisAddIn
    {
        private Excel.Application app;
        private List<string> unicodeFiles; //a list of opened Unicode CSV files. We populate this list on WorkBookOpen event to avoid checking for CSV files on every Save event.
        private bool sFlag = false;

        //Unicode file byte order marks.
        private const string UTF_16BE_BOM = "FEFF";
        private const string UTF_16LE_BOM = "FFFE";
        private const string UTF_8_BOM = "EFBBBF";
        private string GLOBAL_DIRECTORY = "";
        private ToolTip toolTip = new ToolTip();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            app = Application;
            unicodeFiles = new List<string>();
            app.WorkbookOpen += App_WorkbookOpen;
            app.WorkbookBeforeClose += App_WorkbookBeforeClose;
            app.WorkbookBeforeSave += App_WorkbookBeforeSave;
            app.WindowActivate += App_WindowActivate;
        }

        void App_WindowActivate(Excel.Workbook Wb, Excel.Window wn)
        {
            SetGlobalDirectory();
            SetDiffRegionMenu();
        }

        public bool GetGlobalDirectory(out string directory)
        {
            string name = app.ActiveWorkbook.Name;
            app.StatusBar = name;
            if (!name.ToLower().EndsWith(".csv") ||
                !name.StartsWith("CSV"))
            {
                directory = "";
                return false;
            }

            string fullName = app.ActiveWorkbook.FullName;
            app.StatusBar = fullName;
            string[] split = fullName.Split('\\');
            if (split.Length < 3)
            {
                directory = "";
                return false;
            }

            directory = split[split.Length - 3];
            app.StatusBar = directory;
            if (string.IsNullOrEmpty(directory)) return false;
            return true;
        }

        void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            app = null;
            unicodeFiles = null;
        }

        void App_WorkbookBeforeSave(Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            //Override Save behaviour for Unicode CSV files.
            if (!SaveAsUI && !sFlag && unicodeFiles.Contains(Wb.FullName))
            {
                Cancel = true;
                SaveAsUnicodeCSV(false, false);
            }
            sFlag = false;
        }

        //This is required to show our custom Ribbon
        //protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        //{
        //    return new Ribbon1();
        //}

        void App_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel)
        {
            unicodeFiles.Remove(Wb.FullName);
            app.StatusBar = "就绪";
        }

        void App_WorkbookOpen(Excel.Workbook Wb)
        {
            //Check to see if the opened document is a Unicode CSV files, so we can override Excel's Save method
            if (Wb.FullName.ToLower().EndsWith(".csv"))
            {
                if (isFileUnicode(Wb.FullName) && !unicodeFiles.Contains(Wb.FullName))
                {
                    unicodeFiles.Add(Wb.FullName);
                }

                FrozenTrailing();
                //FormatColor();
                SetDiffRegionMenu();
            }
            else
            {
                app.StatusBar = "就绪";
            }
        }

        /// <summary>
        /// This method check whether Excel is in Cell Editing mode or not
        /// There are few ways to check this (eg. check to see if a standard menu item is disabled etc.)
        /// I know in cell editing mode app.DisplayAlerts throws an Exception, so here I'm relying on that behaviour
        /// </summary>
        /// <returns>true if Excel is in cell editing mode</returns>
        private bool isInCellEditingMode()
        {
            bool flag = false;
            try
            {
                app.DisplayAlerts = false; //This will throw an Exception if Excel is in Cell Editing Mode
            }
            catch (Exception)
            {
                flag = true;
            }
            return flag;
        }

        /// <summary>
        /// This will create a temporary file in Unicode text (*.txt) format, overwrite the current loaded file by replaing all tabs with a comma and reload the file.
        /// </summary>
        /// <param name="force">To force save the current file as a Unicode CSV.
        /// When called from the Ribbon items Save/SaveAs, <i>force</i> will be true
        /// If this parameter is true and the file name extention is not .csv, then a SaveAs dialog will be displayed to choose a .csv file</param>
        /// <param name="newFile">To show a SaveAs dialog box to select a new file name
        /// This will be set to true when called from the Ribbon item SaveAs</param>
        public void SaveAsUnicodeCSV(bool force, bool newFile)
        {
            if (!CheckId(false)) return;
            //if (!CheckContent()) return;

            bool currDispAlert = app.DisplayAlerts;
            bool flag = true;
            int i;
            string filename = app.ActiveWorkbook.FullName;

            if (force) //then make sure a csv file name is selected.
            {
                if (newFile || !filename.ToLower().EndsWith(".csv"))
                {
                    Office.FileDialog d = app.get_FileDialog(Office.MsoFileDialogType.msoFileDialogSaveAs);
                    i = app.ActiveWorkbook.Name.LastIndexOf(".");
                    if (i >= 0)
                    {
                        d.InitialFileName = app.ActiveWorkbook.Name.Substring(0, i);
                    }
                    else
                    {
                        d.InitialFileName = app.ActiveWorkbook.Name;
                    }
                    d.AllowMultiSelect = false;
                    Office.FileDialogFilters f = d.Filters;
                    for (i = 1; i <= f.Count; i++)
                    {
                        if ("*.csv".Equals(f.Item(i).Extensions))
                        {
                            d.FilterIndex = i;
                            break;
                        }
                    }
                    if (d.Show() == 0) //User cancelled the dialog
                    {
                        flag = false;
                    }
                    else
                    {
                        filename = d.SelectedItems.Item(1);
                    }
                }
                if (flag && !filename.ToLower().EndsWith(".csv"))
                {
                    MessageBox.Show("请先选择一个 CSV 文件来保存。");
                    flag = false;
                }
            }

            if (flag && filename.ToLower().EndsWith(".csv") && (force || unicodeFiles.Contains(filename)))
            {
                if (isInCellEditingMode())
                {
                    MessageBox.Show("请先完成当前单元格的编辑，并选择到一个空白的单元格。");
                }
                else
                {
                    try
                    {
                        //Getting current selection to restore the current cell selection
                        Excel.Range rng = (Excel.Range)app.ActiveCell;
                        int row = rng.Row;
                        int col = rng.Column;

                        string tempFile = System.IO.Path.GetTempFileName();

                        try
                        {
                            sFlag = true; //This is to prevent this method getting called again from app_WorkbookBeforeSave event caused by the next SaveAs call
                            app.ActiveWorkbook.SaveAs(tempFile, Excel.XlFileFormat.xlUnicodeText);
                            app.ActiveWorkbook.Close();

                            if (new FileInfo(tempFile).Length <= (1024 * 1024)) //If its less than 1MB, load the whole data to memory for character replacement
                            {
                                File.WriteAllText(filename, File.ReadAllText(tempFile, UnicodeEncoding.UTF8).Replace("\t", ","), UnicodeEncoding.UTF8);
                            }
                            else //otherwise read chunks for data (in 10KB chunks) into memory
                            {
                                using (StreamReader sr = new StreamReader(tempFile, true))
                                using (StreamWriter sw = new StreamWriter(filename, false, sr.CurrentEncoding))
                                {
                                    char[] buffer = new char[10 * 1024]; //10KB Chunks
                                    while (!sr.EndOfStream)
                                    {
                                        int cnt = sr.ReadBlock(buffer, 0, buffer.Length);
                                        for (i = 0; i < cnt; i++)
                                        {
                                            if (buffer[i] == '\t')
                                            {
                                                buffer[i] = ',';
                                            }
                                        }
                                        sw.Write(buffer, 0, cnt);
                                    }
                                }
                            }
                        }
                        finally
                        {
                            File.Delete(tempFile);
                        }

                        app.Workbooks.Open(filename, Type.Missing, Type.Missing, Excel.XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing, ",");
                        Excel.Worksheet ws = app.ActiveWorkbook.ActiveSheet;
                        ws.Cells[row, col].Select();
                        app.StatusBar = "已保存为 UTF-8 编码的 CSV 文件。";
                        FrozenTrailing();
                        if (!unicodeFiles.Contains(filename))
                        {
                            unicodeFiles.Add(filename);
                        }
                        app.ActiveWorkbook.Saved = true;
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show("保存为 UTF-8 编码的 CSV 文件时发生错误：" + e.Message);
                    }
                    finally
                    {
                        app.DisplayAlerts = currDispAlert;
                    }
                }
            }
        }

        /// <summary>
        /// This method will try and read the first few bytes to see if it contains a Unicode BOM
        /// </summary>
        /// <param name="filename">File to check for including full path</param>
        /// <returns>true if its a Unicode file</returns>
        private bool isFileUnicode(string filename)
        {
            bool ret = false;
            try
            {
                byte[] buff = new byte[3];
                using (FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    fs.Read(buff, 0, 3);
                }

                string hx = "";
                foreach (byte letter in buff)
                {
                    hx += string.Format("{0:X2}", Convert.ToInt32(letter));
                    //Checking to see the first bytes matches with any of the defined Unicode BOM
                    //We only check for UTF8 and UTF16 here.
                    ret = UTF_8_BOM.Equals(hx) || UTF_16BE_BOM.Equals(hx) || UTF_16LE_BOM.Equals(hx);
                    if (ret)
                    {
                        break;
                    }
                }
            }
            catch (IOException)
            {
                //ignore any exception
            }
            return ret;
        }

        public void ExpandColumn()
        {
            Excel.Worksheet workSheet = app.ActiveWorkbook.ActiveSheet;
            Excel.Range range = workSheet.UsedRange;
            Excel.Range newRange = workSheet.Range[range.Cells[2, 1], range.Cells[range.Rows.Count, range.Columns.Count]];
            newRange.Columns.ShrinkToFit = false;
            newRange.Columns.AutoFit();
        }

        public void CollapseColumn()
        {
            Excel.Worksheet workSheet = app.ActiveWorkbook.ActiveSheet;
            Excel.Range range = workSheet.UsedRange;
            range.EntireColumn.ShrinkToFit = false;
            range.EntireColumn.UseStandardWidth = true;
        }

        public void FormatCell()
        {
            Excel.Worksheet workSheet = app.ActiveWorkbook.ActiveSheet;
            Excel.Range range = workSheet.UsedRange;
            for (int i = 4; i <= range.Rows.Count; i++)
            {
                Excel.Range cell2 = range.Cells[i, 2];
                if (cell2.Value2 == null)
                    return;

                Excel.Range cell1 = range.Cells[i, 1];
                cell1.Value2 = (i - 3).ToString();
            }
        }

        public bool CheckId(bool showMessagebox)
        {
            string name = app.ActiveWorkbook.FullName;
            if (!name.ToLower().EndsWith(".csv")) return true;

            Excel.Worksheet workSheet = app.ActiveWorkbook.ActiveSheet;
            Excel.Range range = workSheet.UsedRange;

            string title = range.Cells[2, 1].Value2.ToString();
            if (title != "id" && title != "Name")
            {
                return true;
            }

            Dictionary<string, int> dict = new Dictionary<string, int>();
            for (int i = 4; 1 <= range.Rows.Count; i++)
            {
                Excel.Range cell = range.Cells[i, 1];
                if (cell.Value2 == null)
                    break;

                //int value = 0;
                string value = cell.Value2.ToString();
                if (string.IsNullOrEmpty(value))
                {
                    MessageBox.Show("第" + i + "行序号错误。", "序号错误", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return false;
                }
                else
                {
                    if (!dict.ContainsKey(value))
                    {
                        dict[value] = i;
                    }
                    else
                    {
                        cell.Activate();
                        MessageBox.Show("第" + i + "行序号[" + value + "]与第" + dict[value] + "行序号重复。", "序号重复", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return false;
                    }
                }
            }
            if (showMessagebox)
                MessageBox.Show("序号检查完成，未发现重复序号。", "检查完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return true;
        }

        public bool CheckContent()
        {
            string name = app.ActiveWorkbook.FullName;
            if (!name.Contains("CSVUTF8Strings")) return true;

            Excel.Worksheet workSheet = app.ActiveWorkbook.ActiveSheet;
            Excel.Range range = workSheet.UsedRange;
            for (int i = 4; 1 <= range.Rows.Count; i++)
            {
                Excel.Range cell = range.Cells[i, 2];
                if (string.IsNullOrEmpty(cell.Value2) || string.IsNullOrWhiteSpace(cell.Value2))
                {
                    cell.Activate();
                    MessageBox.Show("第" + i + "行内容错误。", "内容错误", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return false;
                }
            }
            return true;
        }

        public void CSVChecker(bool checkAll)
        {
            if (String.IsNullOrEmpty(GLOBAL_DIRECTORY))
                return;

            System.Diagnostics.Process process = new System.Diagnostics.Process();
            string filePath = Path.GetDirectoryName(app.ActiveWorkbook.FullName) + "/../../../../Tools/CsvChecker/CSVCheckerDev.bat";
            if (!File.Exists(filePath))
            {
                MessageBox.Show("无法检查当前表格，文件路径有误", "文件路径错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            process.StartInfo.WorkingDirectory = Path.GetDirectoryName(filePath);
            process.StartInfo.FileName = filePath;
            process.StartInfo.UseShellExecute = true;
            process.StartInfo.CreateNoWindow = false;
            if (checkAll)
            {
                process.StartInfo.Arguments = GLOBAL_DIRECTORY;
            }
            else
            {
                int index = app.ActiveWorkbook.Name.LastIndexOf(".");
                process.StartInfo.Arguments = GLOBAL_DIRECTORY + "," + app.ActiveWorkbook.Name.Substring(0, index);
            }
            process.Start();
        }

        public void OpenCheck()
        {
            string name = app.ActiveWorkbook.Name;
            if (!name.StartsWith("CSV")) return;

            name = name.Replace("CSV", "");
            string filePath = Path.GetDirectoryName(app.ActiveWorkbook.FullName) + "/../../../../Tools/CsvChecker/checks/Check" + name;
            //FileInfo fileInfo = new FileInfo(filePath);
            if (File.Exists(filePath))
            {
                //fileInfo.IsReadOnly = false;
                //Excel.Application excel = new Excel.Application();
                //Excel.Workbook workbook = excel.Application.Workbooks.Add(fileInfo.FullName);
                //workbook.ChangeFileAccess(Excel.XlFileAccess.xlReadWrite);
                //excel.Visible = true;
                System.Diagnostics.Process.Start(filePath);
            }
            else
            {
                MessageBox.Show("没有找到对应的检查表", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        public void OpenLocalResource()
        {
            string file = GetCellContent();
            if (string.IsNullOrEmpty(file)) return;

            if (!OpenLocalImage(file))
            {
                MessageBox.Show("没有找到填写的资源。\n" + file, "预览资源", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OpenI18nResource()
        {
            string file = GetCellContent();
            if (string.IsNullOrEmpty(file)) return;

            if (!OpenI18nImage(file))
            {
                MessageBox.Show("没有找到填写的资源。\n" + file, "预览资源", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public bool OpenLocalImage(string file)
        {
            string filePath = Path.GetDirectoryName(app.ActiveWorkbook.FullName) + "/../../../../Client/Assets/XResources/" + file;
            return OpenImage(filePath);
        }

        public bool OpenI18nImage(string file)
        {
            string filePath = Path.GetDirectoryName(app.ActiveWorkbook.FullName) + string.Format("/../../../../Client/Assets/Region/{0}{1}", GLOBAL_DIRECTORY, file);
            return OpenImage(filePath);
        }

        public bool OpenImage(string filePath)
        {
            if (File.Exists(filePath + ".png"))
            {
                ShowImage(filePath + ".png");
                return true;
            }
            else if (File.Exists(filePath + ".jpg"))
            {
                ShowImage(filePath + ".jpg");
                return true;
            }

            if (filePath.Contains("UI/Texture/Character/Card"))
            {
                DirectoryInfo directory = new DirectoryInfo(Path.GetDirectoryName(app.ActiveWorkbook.FullName) + "/../../../../Client/Assets/XResources/UI/Texture/Character/Card");
                if (directory.Exists)
                {
                    FileInfo[] files = directory.GetFiles(filePath.Substring(filePath.LastIndexOf('/') + 1) + ".png", SearchOption.AllDirectories);
                    foreach (FileInfo fi in files)
                    {
                        if (fi.Exists && fi.Extension == ".png")
                        {
                            ShowImage(fi.FullName);
                            return true;
                        }
                    }
                }
            }

            if (imageForm != null)
            {
                imageForm.Hide();
            }
            return false;
        }

        public string GetCellContent()
        {
            string name = app.ActiveWorkbook.Name;
            if (!name.StartsWith("CSV") || !name.EndsWith(".csv")) return null;

            object select = Globals.ThisAddIn.Application.Selection;
            if ((select as Excel.Range) == null) return null;

            Excel.Range range = select as Excel.Range;
            Excel.Range cell = range.Cells[1, 1];
            if (cell.Value2 == null) return null;

            return cell.Value2.ToString();
        }

        ImageForm imageForm;
        public void ShowImage(string path)
        {
            if (imageForm != null)
            {
                imageForm.SetImage(path);
                imageForm.Show();
            }
            else
            {
                imageForm = new ImageForm(path);
                imageForm.Show();
            }
        }

        public void FrozenTrailing()
        {
            string name = app.ActiveWorkbook.FullName;
            if (!name.ToLower().EndsWith(".csv")) return;

            Excel.Worksheet workSheet = app.ActiveWorkbook.ActiveSheet;
            Excel.Range range = workSheet.UsedRange;

            Excel.Range cell = range.Cells[4, GetForzenColumn(name)];
            cell.Activate();
            cell.Application.ActiveWindow.FreezePanes = true;
        }

        public void FormatColor()
        {
            string name = app.ActiveWorkbook.FullName;
            if (!name.StartsWith("CSV")) return;
            if (!name.EndsWith(".csv")) return;

            Excel.Worksheet workSheet = app.ActiveWorkbook.ActiveSheet;
            Excel.Range range = workSheet.UsedRange;

            Excel.Range topRange = range.Range[range.Cells[1, 1], range.Cells[3, range.Columns.Count]];
            topRange.Interior.ColorIndex = 35;

            Excel.Range idRange = range.Range[range.Cells[4, 1], range.Cells[range.Rows.Count, 1]];
            idRange.Interior.ColorIndex = 34;
        }

        int GetForzenColumn(string name)
        {
            if (name.Contains("CSVWarriors.csv"))
                return 5;
            if (name.Contains("CSVSkillCast.csv"))
                return 8;
            if (name.Contains("CSVSkillBuff.csv"))
                return 9;
            if (name.Contains("CSVWarriorGroup.csv"))
                return 3;

            return 2;
        }

        public void SetGlobalDirectory()
        {
            if (GetGlobalDirectory(out GLOBAL_DIRECTORY))
            {
                Globals.Ribbons.Ribbon1.group3.Label = "表格检查 [" + GLOBAL_DIRECTORY + "]";
            }

            if (string.IsNullOrEmpty(GLOBAL_DIRECTORY))
            {
                Globals.Ribbons.Ribbon1.group3.Visible = false;
                Globals.Ribbons.Ribbon1.group4.Visible = false;
                Globals.Ribbons.Ribbon1.group5.Visible = false;
                return;
            }

            Globals.Ribbons.Ribbon1.group3.Visible = true;
            Globals.Ribbons.Ribbon1.group4.Visible = true;
            Globals.Ribbons.Ribbon1.group5.Visible = true;

            if (GLOBAL_DIRECTORY == "_Dev")
            {
                Globals.Ribbons.Ribbon1.button11.Enabled = false;
                Globals.Ribbons.Ribbon1.button12.Enabled = false;
            }
            else
            {
                Globals.Ribbons.Ribbon1.button11.Enabled = true;
                Globals.Ribbons.Ribbon1.button12.Enabled = true;
            }
        }

        public void SetDiffRegionMenu()
        {
            if (string.IsNullOrEmpty(GLOBAL_DIRECTORY))
            {
                Globals.Ribbons.Ribbon1.dropDown1.Enabled = false;
                return;
            }

            Globals.Ribbons.Ribbon1.dropDown1.Enabled = true;

            if (Globals.Ribbons.Ribbon1.dropDown1.SelectedItem.Label == GLOBAL_DIRECTORY)
                return;

            foreach (Microsoft.Office.Tools.Ribbon.RibbonDropDownItem item in Globals.Ribbons.Ribbon1.dropDown1.Items)
            {
                if (item.Label == GLOBAL_DIRECTORY)
                {
                    Globals.Ribbons.Ribbon1.dropDown1.SelectedItem = item;
                    break;
                }
            }
        }

        public void FormatStyle()
        {
            Excel.Worksheet workSheet = app.ActiveWorkbook.ActiveSheet;
            Excel.Range range = workSheet.UsedRange;
            range.ClearFormats();
            range.NumberFormatLocal = "@";

            int i = 0, j = 0;
            string tmp = "";
            try
            {
                string[,] data = new string[range.Rows.Count, range.Columns.Count];
                for (i = 0; i <= range.Rows.Count; i++)
                {
                    for (j = 0; j <= range.Columns.Count; i++)
                    {
                        Excel.Range cell = range.Cells[i + 1, j + 1];
                        //cell.NumberFormatLocal = "@";
                        data[i, j] = cell.Value2.ToString();
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Error" + "[" + i + "," + j + "]" + tmp + " "  + e.Message);
            }

            //for (int i = 0; i <= range.Rows.Count; i++)
            //{
            //    for (int j = 0; j <= range.Columns.Count; i++)
            //    {
            //        Excel.Range cell = range.Cells[i + 1, j + 1];
            //        cell.Value = data[i, j];
            //        cell.NumberFormat = "@";
            //    }
            //}
        }

        void InvokeSVNProcess(string command)
        {
            System.Diagnostics.Process.Start("TortoiseProc.exe", @"/command:" + command + " /closeonend:0");
        }

        internal void SVNCommit()
        {
            InvokeSVNProcess("commit /path:" + app.ActiveWorkbook.FullName);
        }

        internal void SVNRevert()
        {
            InvokeSVNProcess("revert /path:" + app.ActiveWorkbook.FullName);
        }

        internal void SVNDiff()
        {
            InvokeSVNProcess("diff /path:" + app.ActiveWorkbook.FullName);
        }

        internal void SVNLog()
        {
            InvokeSVNProcess("log /path:" + app.ActiveWorkbook.FullName);
        }

        internal void SVNRegionDiff(string region)
        {
            string path2 = app.ActiveWorkbook.FullName;
            string path1 = path2.Replace(GLOBAL_DIRECTORY, region);
            InvokeSVNProcess("diff /path:" + path1 + " /path2:" + path2);
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
