namespace Counsel_System2.Forms
{
    partial class ABCardPrintForm
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
            this.ilblDownTemplate = new System.Windows.Forms.LinkLabel();
            this.ilblUploadTemplate = new System.Windows.Forms.LinkLabel();
            this.ilblExportExcelField = new System.Windows.Forms.LinkLabel();
            this.groupPanel1 = new DevComponents.DotNetBar.Controls.GroupPanel();
            this.ilbDefaultTemplate = new System.Windows.Forms.LinkLabel();
            this.lblMsg = new DevComponents.DotNetBar.LabelX();
            this.groupPanel2 = new DevComponents.DotNetBar.Controls.GroupPanel();
            this.chkSplitByStudentNumber = new System.Windows.Forms.CheckBox();
            this.groupPanel1.SuspendLayout();
            this.groupPanel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnPrint
            // 
            this.btnPrint.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnPrint.BackColor = System.Drawing.Color.Transparent;
            this.btnPrint.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnPrint.Location = new System.Drawing.Point(128, 191);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(57, 23);
            this.btnPrint.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnPrint.TabIndex = 0;
            this.btnPrint.Text = "列印";
            this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Location = new System.Drawing.Point(191, 191);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(57, 23);
            this.btnExit.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnExit.TabIndex = 1;
            this.btnExit.Text = "離開";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // ilblDownTemplate
            // 
            this.ilblDownTemplate.AutoSize = true;
            this.ilblDownTemplate.BackColor = System.Drawing.Color.Transparent;
            this.ilblDownTemplate.Location = new System.Drawing.Point(13, 12);
            this.ilblDownTemplate.Name = "ilblDownTemplate";
            this.ilblDownTemplate.Size = new System.Drawing.Size(86, 17);
            this.ilblDownTemplate.TabIndex = 2;
            this.ilblDownTemplate.TabStop = true;
            this.ilblDownTemplate.Text = "下載目前樣板";
            this.ilblDownTemplate.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.ilblDownTemplate_LinkClicked);
            // 
            // ilblUploadTemplate
            // 
            this.ilblUploadTemplate.AutoSize = true;
            this.ilblUploadTemplate.BackColor = System.Drawing.Color.Transparent;
            this.ilblUploadTemplate.Location = new System.Drawing.Point(129, 12);
            this.ilblUploadTemplate.Name = "ilblUploadTemplate";
            this.ilblUploadTemplate.Size = new System.Drawing.Size(86, 17);
            this.ilblUploadTemplate.TabIndex = 3;
            this.ilblUploadTemplate.TabStop = true;
            this.ilblUploadTemplate.Text = "上傳自訂樣板";
            this.ilblUploadTemplate.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.ilblUploadTemplate_LinkClicked);
            // 
            // ilblExportExcelField
            // 
            this.ilblExportExcelField.AutoSize = true;
            this.ilblExportExcelField.BackColor = System.Drawing.Color.Transparent;
            this.ilblExportExcelField.Location = new System.Drawing.Point(13, 46);
            this.ilblExportExcelField.Name = "ilblExportExcelField";
            this.ilblExportExcelField.Size = new System.Drawing.Size(86, 17);
            this.ilblExportExcelField.TabIndex = 4;
            this.ilblExportExcelField.TabStop = true;
            this.ilblExportExcelField.Text = "產生合併欄位";
            this.ilblExportExcelField.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.ilblExportExcelField_LinkClicked);
            // 
            // groupPanel1
            // 
            this.groupPanel1.BackColor = System.Drawing.Color.Transparent;
            this.groupPanel1.CanvasColor = System.Drawing.SystemColors.Control;
            this.groupPanel1.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007;
            this.groupPanel1.Controls.Add(this.ilbDefaultTemplate);
            this.groupPanel1.Controls.Add(this.ilblDownTemplate);
            this.groupPanel1.Controls.Add(this.ilblExportExcelField);
            this.groupPanel1.Controls.Add(this.ilblUploadTemplate);
            this.groupPanel1.Location = new System.Drawing.Point(13, 13);
            this.groupPanel1.Name = "groupPanel1";
            this.groupPanel1.Size = new System.Drawing.Size(236, 100);
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
            this.groupPanel1.TabIndex = 5;
            this.groupPanel1.Text = "樣板";
            this.groupPanel1.Click += new System.EventHandler(this.groupPanel1_Click);
            // 
            // ilbDefaultTemplate
            // 
            this.ilbDefaultTemplate.AutoSize = true;
            this.ilbDefaultTemplate.BackColor = System.Drawing.Color.Transparent;
            this.ilbDefaultTemplate.Location = new System.Drawing.Point(129, 46);
            this.ilbDefaultTemplate.Name = "ilbDefaultTemplate";
            this.ilbDefaultTemplate.Size = new System.Drawing.Size(86, 17);
            this.ilbDefaultTemplate.TabIndex = 5;
            this.ilbDefaultTemplate.TabStop = true;
            this.ilbDefaultTemplate.Text = "使用預設樣版";
            this.ilbDefaultTemplate.Visible = false;
            this.ilbDefaultTemplate.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.ilbDefaultTemplate_LinkClicked);
            // 
            // lblMsg
            // 
            this.lblMsg.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lblMsg.BackgroundStyle.Class = "";
            this.lblMsg.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblMsg.ForeColor = System.Drawing.Color.Red;
            this.lblMsg.Location = new System.Drawing.Point(13, 189);
            this.lblMsg.Name = "lblMsg";
            this.lblMsg.Size = new System.Drawing.Size(89, 23);
            this.lblMsg.TabIndex = 6;
            this.lblMsg.Text = "資料讀取中..";
            // 
            // groupPanel2
            // 
            this.groupPanel2.BackColor = System.Drawing.Color.Transparent;
            this.groupPanel2.CanvasColor = System.Drawing.SystemColors.Control;
            this.groupPanel2.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007;
            this.groupPanel2.Controls.Add(this.chkSplitByStudentNumber);
            this.groupPanel2.Location = new System.Drawing.Point(13, 120);
            this.groupPanel2.Name = "groupPanel2";
            this.groupPanel2.Size = new System.Drawing.Size(235, 63);
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
            this.groupPanel2.TabIndex = 7;
            this.groupPanel2.Text = "產生檔案設定";
            // 
            // chkSplitByStudentNumber
            // 
            this.chkSplitByStudentNumber.AutoSize = true;
            this.chkSplitByStudentNumber.Location = new System.Drawing.Point(16, 6);
            this.chkSplitByStudentNumber.Name = "chkSplitByStudentNumber";
            this.chkSplitByStudentNumber.Size = new System.Drawing.Size(144, 21);
            this.chkSplitByStudentNumber.TabIndex = 0;
            this.chkSplitByStudentNumber.Text = "依學號產生每個檔案";
            this.chkSplitByStudentNumber.UseVisualStyleBackColor = true;
            // 
            // ABCardPrintForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(262, 219);
            this.Controls.Add(this.groupPanel2);
            this.Controls.Add(this.lblMsg);
            this.Controls.Add(this.groupPanel1);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnPrint);
            this.DoubleBuffered = true;
            this.Name = "ABCardPrintForm";
            this.Text = "綜合資料紀錄表";
            this.Load += new System.EventHandler(this.ABCardPrintForm_Load);
            this.groupPanel1.ResumeLayout(false);
            this.groupPanel1.PerformLayout();
            this.groupPanel2.ResumeLayout(false);
            this.groupPanel2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.ButtonX btnPrint;
        private DevComponents.DotNetBar.ButtonX btnExit;
        private System.Windows.Forms.LinkLabel ilblDownTemplate;
        private System.Windows.Forms.LinkLabel ilblUploadTemplate;
        private System.Windows.Forms.LinkLabel ilblExportExcelField;
        private DevComponents.DotNetBar.Controls.GroupPanel groupPanel1;
        private System.Windows.Forms.LinkLabel ilbDefaultTemplate;
        private DevComponents.DotNetBar.LabelX lblMsg;
        private DevComponents.DotNetBar.Controls.GroupPanel groupPanel2;
        private System.Windows.Forms.CheckBox chkSplitByStudentNumber;
    }
}