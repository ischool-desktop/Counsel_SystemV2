namespace Counsel_System2.Forms
{
    partial class StudABCardReportForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnPrint = new DevComponents.DotNetBar.ButtonX();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.groupPanel1 = new DevComponents.DotNetBar.Controls.GroupPanel();
            this.lnkChkMappingField = new System.Windows.Forms.LinkLabel();
            this.lnkUserDefUpload = new System.Windows.Forms.LinkLabel();
            this.lnkUserDef = new System.Windows.Forms.LinkLabel();
            this.lnkDefaultView = new System.Windows.Forms.LinkLabel();
            this.chkUserDef = new DevComponents.DotNetBar.Controls.CheckBoxX();
            this.chkDefault = new DevComponents.DotNetBar.Controls.CheckBoxX();
            this.groupPanel2 = new DevComponents.DotNetBar.Controls.GroupPanel();
            this.chkFileSplitBySNum = new DevComponents.DotNetBar.Controls.CheckBoxX();
            this.chkFileAllInOne = new DevComponents.DotNetBar.Controls.CheckBoxX();
            this.groupPanel1.SuspendLayout();
            this.groupPanel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnPrint
            // 
            this.btnPrint.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnPrint.AutoSize = true;
            this.btnPrint.BackColor = System.Drawing.Color.Transparent;
            this.btnPrint.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnPrint.Location = new System.Drawing.Point(186, 122);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(75, 25);
            this.btnPrint.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnPrint.TabIndex = 0;
            this.btnPrint.Text = "列印";
            this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.AutoSize = true;
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Location = new System.Drawing.Point(267, 122);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 25);
            this.btnExit.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnExit.TabIndex = 1;
            this.btnExit.Text = "離開";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // groupPanel1
            // 
            this.groupPanel1.BackColor = System.Drawing.Color.Transparent;
            this.groupPanel1.CanvasColor = System.Drawing.SystemColors.Control;
            this.groupPanel1.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007;
            this.groupPanel1.Controls.Add(this.lnkChkMappingField);
            this.groupPanel1.Controls.Add(this.lnkUserDefUpload);
            this.groupPanel1.Controls.Add(this.lnkUserDef);
            this.groupPanel1.Controls.Add(this.lnkDefaultView);
            this.groupPanel1.Controls.Add(this.chkUserDef);
            this.groupPanel1.Controls.Add(this.chkDefault);
            this.groupPanel1.Location = new System.Drawing.Point(14, 10);
            this.groupPanel1.Name = "groupPanel1";
            this.groupPanel1.Size = new System.Drawing.Size(327, 94);
            // 
            // 
            // 
            this.groupPanel1.Style.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground2;
            this.groupPanel1.Style.BackColorGradientAngle = 90;
            this.groupPanel1.Style.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground;
            this.groupPanel1.Style.BorderBottom = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel1.Style.BorderBottomWidth = 1;
            this.groupPanel1.Style.BorderColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBorder;
            this.groupPanel1.Style.BorderLeft = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel1.Style.BorderLeftWidth = 1;
            this.groupPanel1.Style.BorderRight = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel1.Style.BorderRightWidth = 1;
            this.groupPanel1.Style.BorderTop = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel1.Style.BorderTopWidth = 1;
            this.groupPanel1.Style.Class = "";
            this.groupPanel1.Style.CornerDiameter = 4;
            this.groupPanel1.Style.CornerType = DevComponents.DotNetBar.eCornerType.Rounded;
            this.groupPanel1.Style.TextColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelText;
            this.groupPanel1.Style.TextLineAlignment = DevComponents.DotNetBar.eStyleTextAlignment.Near;
            // 
            // 
            // 
            this.groupPanel1.StyleMouseDown.Class = "";
            this.groupPanel1.StyleMouseDown.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            // 
            // 
            // 
            this.groupPanel1.StyleMouseOver.Class = "";
            this.groupPanel1.StyleMouseOver.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.groupPanel1.TabIndex = 2;
            this.groupPanel1.Text = "範本";
            // 
            // lnkChkMappingField
            // 
            this.lnkChkMappingField.AutoSize = true;
            this.lnkChkMappingField.Location = new System.Drawing.Point(223, 10);
            this.lnkChkMappingField.Name = "lnkChkMappingField";
            this.lnkChkMappingField.Size = new System.Drawing.Size(86, 17);
            this.lnkChkMappingField.TabIndex = 7;
            this.lnkChkMappingField.TabStop = true;
            this.lnkChkMappingField.Text = "檢視合併欄位";
            this.lnkChkMappingField.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkChkMappingField_LinkClicked);
            // 
            // lnkUserDefUpload
            // 
            this.lnkUserDefUpload.AutoSize = true;
            this.lnkUserDefUpload.Location = new System.Drawing.Point(223, 37);
            this.lnkUserDefUpload.Name = "lnkUserDefUpload";
            this.lnkUserDefUpload.Size = new System.Drawing.Size(34, 17);
            this.lnkUserDefUpload.TabIndex = 6;
            this.lnkUserDefUpload.TabStop = true;
            this.lnkUserDefUpload.Text = "上傳";
            this.lnkUserDefUpload.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkUserDefUpload_LinkClicked);
            // 
            // lnkUserDef
            // 
            this.lnkUserDef.AutoSize = true;
            this.lnkUserDef.Location = new System.Drawing.Point(177, 37);
            this.lnkUserDef.Name = "lnkUserDef";
            this.lnkUserDef.Size = new System.Drawing.Size(34, 17);
            this.lnkUserDef.TabIndex = 5;
            this.lnkUserDef.TabStop = true;
            this.lnkUserDef.Text = "檢視";
            this.lnkUserDef.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkUserDef_LinkClicked);
            // 
            // lnkDefaultView
            // 
            this.lnkDefaultView.AutoSize = true;
            this.lnkDefaultView.Location = new System.Drawing.Point(177, 10);
            this.lnkDefaultView.Name = "lnkDefaultView";
            this.lnkDefaultView.Size = new System.Drawing.Size(34, 17);
            this.lnkDefaultView.TabIndex = 4;
            this.lnkDefaultView.TabStop = true;
            this.lnkDefaultView.Text = "檢視";
            this.lnkDefaultView.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkDefaultView_LinkClicked);
            // 
            // chkUserDef
            // 
            this.chkUserDef.AutoSize = true;
            // 
            // 
            // 
            this.chkUserDef.BackgroundStyle.Class = "";
            this.chkUserDef.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.chkUserDef.CheckBoxStyle = DevComponents.DotNetBar.eCheckBoxStyle.RadioButton;
            this.chkUserDef.Location = new System.Drawing.Point(12, 36);
            this.chkUserDef.Name = "chkUserDef";
            this.chkUserDef.Size = new System.Drawing.Size(147, 21);
            this.chkUserDef.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.chkUserDef.TabIndex = 1;
            this.chkUserDef.Text = "自訂輔導資料紀錄表";
            // 
            // chkDefault
            // 
            this.chkDefault.AutoSize = true;
            // 
            // 
            // 
            this.chkDefault.BackgroundStyle.Class = "";
            this.chkDefault.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.chkDefault.CheckBoxStyle = DevComponents.DotNetBar.eCheckBoxStyle.RadioButton;
            this.chkDefault.Location = new System.Drawing.Point(12, 7);
            this.chkDefault.Name = "chkDefault";
            this.chkDefault.Size = new System.Drawing.Size(147, 21);
            this.chkDefault.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.chkDefault.TabIndex = 0;
            this.chkDefault.Text = "預設輔導資料紀錄表";
            // 
            // groupPanel2
            // 
            this.groupPanel2.BackColor = System.Drawing.Color.Transparent;
            this.groupPanel2.CanvasColor = System.Drawing.SystemColors.Control;
            this.groupPanel2.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007;
            this.groupPanel2.Controls.Add(this.chkFileSplitBySNum);
            this.groupPanel2.Controls.Add(this.chkFileAllInOne);
            this.groupPanel2.Location = new System.Drawing.Point(14, 260);
            this.groupPanel2.Name = "groupPanel2";
            this.groupPanel2.Size = new System.Drawing.Size(291, 66);
            // 
            // 
            // 
            this.groupPanel2.Style.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground2;
            this.groupPanel2.Style.BackColorGradientAngle = 90;
            this.groupPanel2.Style.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground;
            this.groupPanel2.Style.BorderBottom = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel2.Style.BorderBottomWidth = 1;
            this.groupPanel2.Style.BorderColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBorder;
            this.groupPanel2.Style.BorderLeft = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel2.Style.BorderLeftWidth = 1;
            this.groupPanel2.Style.BorderRight = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel2.Style.BorderRightWidth = 1;
            this.groupPanel2.Style.BorderTop = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel2.Style.BorderTopWidth = 1;
            this.groupPanel2.Style.Class = "";
            this.groupPanel2.Style.CornerDiameter = 4;
            this.groupPanel2.Style.CornerType = DevComponents.DotNetBar.eCornerType.Rounded;
            this.groupPanel2.Style.TextColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelText;
            this.groupPanel2.Style.TextLineAlignment = DevComponents.DotNetBar.eStyleTextAlignment.Near;
            // 
            // 
            // 
            this.groupPanel2.StyleMouseDown.Class = "";
            this.groupPanel2.StyleMouseDown.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            // 
            // 
            // 
            this.groupPanel2.StyleMouseOver.Class = "";
            this.groupPanel2.StyleMouseOver.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.groupPanel2.TabIndex = 3;
            this.groupPanel2.Text = "檔案產生方式";
            // 
            // chkFileSplitBySNum
            // 
            this.chkFileSplitBySNum.AutoSize = true;
            // 
            // 
            // 
            this.chkFileSplitBySNum.BackgroundStyle.Class = "";
            this.chkFileSplitBySNum.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.chkFileSplitBySNum.CheckBoxStyle = DevComponents.DotNetBar.eCheckBoxStyle.RadioButton;
            this.chkFileSplitBySNum.Location = new System.Drawing.Point(140, 7);
            this.chkFileSplitBySNum.Name = "chkFileSplitBySNum";
            this.chkFileSplitBySNum.Size = new System.Drawing.Size(107, 21);
            this.chkFileSplitBySNum.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.chkFileSplitBySNum.TabIndex = 1;
            this.chkFileSplitBySNum.Text = "依學號分檔案";
            // 
            // chkFileAllInOne
            // 
            this.chkFileAllInOne.AutoSize = true;
            // 
            // 
            // 
            this.chkFileAllInOne.BackgroundStyle.Class = "";
            this.chkFileAllInOne.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.chkFileAllInOne.CheckBoxStyle = DevComponents.DotNetBar.eCheckBoxStyle.RadioButton;
            this.chkFileAllInOne.Location = new System.Drawing.Point(12, 7);
            this.chkFileAllInOne.Name = "chkFileAllInOne";
            this.chkFileAllInOne.Size = new System.Drawing.Size(107, 21);
            this.chkFileAllInOne.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.chkFileAllInOne.TabIndex = 0;
            this.chkFileAllInOne.Text = "產生單一檔案";
            // 
            // StudABCardReportForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(357, 156);
            this.Controls.Add(this.groupPanel2);
            this.Controls.Add(this.groupPanel1);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnPrint);
            this.DoubleBuffered = true;
            this.Name = "StudABCardReportForm";
            this.Text = "學生綜合資料紀錄表";
            this.Load += new System.EventHandler(this.StudABCardReportForm_Load);
            this.groupPanel1.ResumeLayout(false);
            this.groupPanel1.PerformLayout();
            this.groupPanel2.ResumeLayout(false);
            this.groupPanel2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.ButtonX btnPrint;
        private DevComponents.DotNetBar.ButtonX btnExit;
        private DevComponents.DotNetBar.Controls.GroupPanel groupPanel1;
        private DevComponents.DotNetBar.Controls.CheckBoxX chkDefault;
        private DevComponents.DotNetBar.Controls.CheckBoxX chkUserDef;
        private System.Windows.Forms.LinkLabel lnkDefaultView;
        private System.Windows.Forms.LinkLabel lnkUserDefUpload;
        private System.Windows.Forms.LinkLabel lnkUserDef;
        private DevComponents.DotNetBar.Controls.CheckBoxX chkFileSplitBySNum;
        private DevComponents.DotNetBar.Controls.GroupPanel groupPanel2;
        private DevComponents.DotNetBar.Controls.CheckBoxX chkFileAllInOne;
        private System.Windows.Forms.LinkLabel lnkChkMappingField;
    }
}