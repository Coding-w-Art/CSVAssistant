namespace CSVAssistant
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl6 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl7 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl8 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl9 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl10 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl11 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl12 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl13 = this.Factory.CreateRibbonDropDownItem();
            this.CSVAssistant = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.separator4 = this.Factory.CreateRibbonSeparator();
            this.label2 = this.Factory.CreateRibbonLabel();
            this.dropDown1 = this.Factory.CreateRibbonDropDown();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button17 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.button4 = this.Factory.CreateRibbonButton();
            this.button5 = this.Factory.CreateRibbonButton();
            this.button6 = this.Factory.CreateRibbonButton();
            this.button7 = this.Factory.CreateRibbonButton();
            this.button8 = this.Factory.CreateRibbonButton();
            this.button9 = this.Factory.CreateRibbonButton();
            this.button10 = this.Factory.CreateRibbonButton();
            this.button11 = this.Factory.CreateRibbonButton();
            this.button12 = this.Factory.CreateRibbonButton();
            this.button13 = this.Factory.CreateRibbonButton();
            this.button16 = this.Factory.CreateRibbonButton();
            this.button14 = this.Factory.CreateRibbonButton();
            this.button15 = this.Factory.CreateRibbonButton();
            this.group6 = this.Factory.CreateRibbonGroup();
            this.button18 = this.Factory.CreateRibbonButton();
            this.separator5 = this.Factory.CreateRibbonSeparator();
            this.CSVAssistant.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.group5.SuspendLayout();
            this.group6.SuspendLayout();
            this.SuspendLayout();
            // 
            // CSVAssistant
            // 
            this.CSVAssistant.Groups.Add(this.group1);
            this.CSVAssistant.Groups.Add(this.group2);
            this.CSVAssistant.Groups.Add(this.group3);
            this.CSVAssistant.Groups.Add(this.group4);
            this.CSVAssistant.Groups.Add(this.group5);
            this.CSVAssistant.Groups.Add(this.group6);
            this.CSVAssistant.Label = "CSV 表格助手";
            this.CSVAssistant.Name = "CSVAssistant";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button1);
            this.group1.Items.Add(this.separator1);
            this.group1.Items.Add(this.button2);
            this.group1.Label = "保存";
            this.group1.Name = "group1";
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // group2
            // 
            this.group2.Items.Add(this.button3);
            this.group2.Items.Add(this.button4);
            this.group2.Items.Add(this.separator2);
            this.group2.Items.Add(this.button5);
            this.group2.Items.Add(this.button6);
            this.group2.Label = "快捷工具";
            this.group2.Name = "group2";
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // group3
            // 
            this.group3.Items.Add(this.button7);
            this.group3.Items.Add(this.button8);
            this.group3.Items.Add(this.button9);
            this.group3.Label = "表格检查";
            this.group3.Name = "group3";
            // 
            // group4
            // 
            this.group4.Items.Add(this.button10);
            this.group4.Items.Add(this.separator3);
            this.group4.Items.Add(this.button11);
            this.group4.Items.Add(this.button12);
            this.group4.Label = "资源预览";
            this.group4.Name = "group4";
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // group5
            // 
            this.group5.Items.Add(this.button13);
            this.group5.Items.Add(this.button16);
            this.group5.Items.Add(this.button14);
            this.group5.Items.Add(this.separator4);
            this.group5.Items.Add(this.button15);
            this.group5.Items.Add(this.label2);
            this.group5.Items.Add(this.dropDown1);
            this.group5.Label = "SVN 工具";
            this.group5.Name = "group5";
            // 
            // separator4
            // 
            this.separator4.Name = "separator4";
            // 
            // label2
            // 
            this.label2.Label = "和其他大区比较：";
            this.label2.Name = "label2";
            // 
            // dropDown1
            // 
            ribbonDropDownItemImpl1.Label = "_Dev";
            ribbonDropDownItemImpl1.OfficeImageId = "ReviewCompareAndMerge";
            ribbonDropDownItemImpl1.Tag = "_Dev";
            ribbonDropDownItemImpl2.Label = "CN";
            ribbonDropDownItemImpl2.OfficeImageId = "ReviewCompareAndMerge";
            ribbonDropDownItemImpl2.Tag = "CN";
            ribbonDropDownItemImpl3.Label = "CN_APP";
            ribbonDropDownItemImpl3.OfficeImageId = "ReviewCompareAndMerge";
            ribbonDropDownItemImpl3.Tag = "CN_APP";
            ribbonDropDownItemImpl4.Label = "HMT";
            ribbonDropDownItemImpl4.OfficeImageId = "ReviewCompareAndMerge";
            ribbonDropDownItemImpl4.Tag = "HMT";
            ribbonDropDownItemImpl5.Label = "CN_APP_HRG";
            ribbonDropDownItemImpl5.OfficeImageId = "ReviewCompareAndMerge";
            ribbonDropDownItemImpl5.Tag = "CN_APP_HRG";
            ribbonDropDownItemImpl6.Label = "CN_Mailiang";
            ribbonDropDownItemImpl6.OfficeImageId = "ReviewCompareAndMerge";
            ribbonDropDownItemImpl6.Tag = "CN_Mailiang";
            ribbonDropDownItemImpl7.Label = "CN_Xinji";
            ribbonDropDownItemImpl7.OfficeImageId = "ReviewCompareAndMerge";
            ribbonDropDownItemImpl7.Tag = "CN_Xinji";
            ribbonDropDownItemImpl8.Label = "SM";
            ribbonDropDownItemImpl8.OfficeImageId = "ReviewCompareAndMerge";
            ribbonDropDownItemImpl8.Tag = "SM";
            ribbonDropDownItemImpl9.Label = "SM_EN";
            ribbonDropDownItemImpl9.OfficeImageId = "ReviewCompareAndMerge";
            ribbonDropDownItemImpl9.Tag = "SM_EN";
            ribbonDropDownItemImpl10.Label = "SM_EN_zh_TW";
            ribbonDropDownItemImpl10.OfficeImageId = "ReviewCompareAndMerge";
            ribbonDropDownItemImpl10.Tag = "SM_EN_zh_TW";
            ribbonDropDownItemImpl11.Label = "TH";
            ribbonDropDownItemImpl11.OfficeImageId = "ReviewCompareAndMerge";
            ribbonDropDownItemImpl11.Tag = "TH";
            ribbonDropDownItemImpl12.Label = "JP";
            ribbonDropDownItemImpl12.OfficeImageId = "ReviewCompareAndMerge";
            ribbonDropDownItemImpl12.Tag = "JP";
            ribbonDropDownItemImpl13.Label = "SEA";
            ribbonDropDownItemImpl13.OfficeImageId = "ReviewCompareAndMerge";
            ribbonDropDownItemImpl13.Tag = "SEA";
            this.dropDown1.Items.Add(ribbonDropDownItemImpl1);
            this.dropDown1.Items.Add(ribbonDropDownItemImpl2);
            this.dropDown1.Items.Add(ribbonDropDownItemImpl3);
            this.dropDown1.Items.Add(ribbonDropDownItemImpl4);
            this.dropDown1.Items.Add(ribbonDropDownItemImpl5);
            this.dropDown1.Items.Add(ribbonDropDownItemImpl6);
            this.dropDown1.Items.Add(ribbonDropDownItemImpl7);
            this.dropDown1.Items.Add(ribbonDropDownItemImpl8);
            this.dropDown1.Items.Add(ribbonDropDownItemImpl9);
            this.dropDown1.Items.Add(ribbonDropDownItemImpl10);
            this.dropDown1.Items.Add(ribbonDropDownItemImpl11);
            this.dropDown1.Items.Add(ribbonDropDownItemImpl12);
            this.dropDown1.Items.Add(ribbonDropDownItemImpl13);
            this.dropDown1.Label = "和其他大区比较";
            this.dropDown1.Name = "dropDown1";
            this.dropDown1.OfficeImageId = "ReviewCompareAndMerge";
            this.dropDown1.ShowImage = true;
            this.dropDown1.ShowLabel = false;
            this.dropDown1.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.dropDown1_SelectionChanged);
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Label = "保存";
            this.button1.Name = "button1";
            this.button1.OfficeImageId = "FileSave";
            this.button1.ScreenTip = "保存当前的 CSV 文件。如果当前文件不是 CSV，则另存为 CSV 文件。";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SaveButtonAction);
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Label = "另存为…";
            this.button2.Name = "button2";
            this.button2.OfficeImageId = "FileSaveAs";
            this.button2.ScreenTip = "另存为 UTF-8 编码的 CSV 文件。";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SaveAsButtonAction);
            // 
            // button17
            // 
            this.button17.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button17.Label = "加载样式";
            this.button17.Name = "button17";
            this.button17.OfficeImageId = "XDRichTextArea";
            this.button17.ShowImage = true;
            this.button17.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button17_Click);
            // 
            // button3
            // 
            this.button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button3.Label = "最佳宽度";
            this.button3.Name = "button3";
            this.button3.OfficeImageId = "SizeToControlWidth";
            this.button3.ScreenTip = "根据内容长度调整列宽度，使内容完整显示出来。";
            this.button3.ShowImage = true;
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ExpandAction);
            // 
            // button4
            // 
            this.button4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button4.Label = "默认宽度";
            this.button4.Name = "button4";
            this.button4.OfficeImageId = "FormatCellsMenu";
            this.button4.ScreenTip = "将所有列的宽度调整为默认宽度（80像素）。";
            this.button4.ShowImage = true;
            this.button4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CollapseAction);
            // 
            // button5
            // 
            this.button5.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button5.Label = "生成序号";
            this.button5.Name = "button5";
            this.button5.OfficeImageId = "FormatNumberDefault";
            this.button5.ScreenTip = "将序号列由1开始递增填充。";
            this.button5.ShowImage = true;
            this.button5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.FormatAction);
            // 
            // button6
            // 
            this.button6.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button6.Label = "检查序号";
            this.button6.Name = "button6";
            this.button6.OfficeImageId = "FileViewDigitalSignatures";
            this.button6.ScreenTip = "检查序号列是否有重复的序号。";
            this.button6.ShowImage = true;
            this.button6.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.IdCheckAction);
            // 
            // button7
            // 
            this.button7.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button7.Label = "检查当前表格";
            this.button7.Name = "button7";
            this.button7.OfficeImageId = "FileMarkAsFinal";
            this.button7.ScreenTip = "使用表格检查工具检查当前表格。";
            this.button7.ShowImage = true;
            this.button7.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CSVCheckAction);
            // 
            // button8
            // 
            this.button8.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button8.Label = "检查所有表格";
            this.button8.Name = "button8";
            this.button8.OfficeImageId = "ReviewEndReviewPowerPoint";
            this.button8.ScreenTip = "使用表格检查工具检查所有表格。";
            this.button8.ShowImage = true;
            this.button8.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CSVCheckAllAction);
            // 
            // button9
            // 
            this.button9.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button9.Label = "打开检查表";
            this.button9.Name = "button9";
            this.button9.OfficeImageId = "OpenAttachedMasterPage";
            this.button9.ScreenTip = "打开对应CSV的检查配置表格。";
            this.button9.ShowImage = true;
            this.button9.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CSVOpenCheckAction);
            // 
            // button10
            // 
            this.button10.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button10.Label = "预览默认资源";
            this.button10.Name = "button10";
            this.button10.OfficeImageId = "OmsImageFromFile";
            this.button10.ScreenTip = "预览选中单元格中资源路径对应的默认图片资源。";
            this.button10.ShowImage = true;
            this.button10.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OpenImageAction);
            // 
            // button11
            // 
            this.button11.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button11.Label = "预览国际化资源";
            this.button11.Name = "button11";
            this.button11.OfficeImageId = "OmsImageFromClip";
            this.button11.ScreenTip = "预览选中单元格中国际化资源路径对应的图片资源。";
            this.button11.ShowImage = true;
            this.button11.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OpenIOSImageAction);
            // 
            // button12
            // 
            this.button12.Label = "";
            this.button12.Name = "button12";
            // 
            // button13
            // 
            this.button13.Label = "提交";
            this.button13.Name = "button13";
            this.button13.OfficeImageId = "UpgradeWorkbook";
            this.button13.ShowImage = true;
            this.button13.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button13_Click);
            // 
            // button16
            // 
            this.button16.Label = "查看日志";
            this.button16.Name = "button16";
            this.button16.OfficeImageId = "ContactProperties";
            this.button16.ShowImage = true;
            this.button16.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button16_Click);
            // 
            // button14
            // 
            this.button14.Label = "还原";
            this.button14.Name = "button14";
            this.button14.OfficeImageId = "Refresh";
            this.button14.ShowImage = true;
            this.button14.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button14_Click);
            // 
            // button15
            // 
            this.button15.Label = "和仓库版本比较";
            this.button15.Name = "button15";
            this.button15.OfficeImageId = "ReviewCompareLastVersion";
            this.button15.ShowImage = true;
            this.button15.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button15_Click);
            // 
            // group6
            // 
            this.group6.Items.Add(this.button17);
            this.group6.Items.Add(this.separator5);
            this.group6.Items.Add(this.button18);
            this.group6.Label = "样式工具";
            this.group6.Name = "group6";
            // 
            // button18
            // 
            this.button18.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button18.Label = "清除样式";
            this.button18.Name = "button18";
            this.button18.OfficeImageId = "ClearAllFormatting";
            this.button18.ShowImage = true;
            this.button18.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button18_Click);
            // 
            // separator5
            // 
            this.separator5.Name = "separator5";
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.CSVAssistant);
            this.CSVAssistant.ResumeLayout(false);
            this.CSVAssistant.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();
            this.group6.ResumeLayout(false);
            this.group6.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab CSVAssistant;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button6;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button8;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button9;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button10;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button11;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button12;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button13;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button14;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator4;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDown1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button15;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button16;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button17;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button18;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator5;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
