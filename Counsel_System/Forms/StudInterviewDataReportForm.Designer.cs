namespace Counsel_System2.Forms
{
    partial class StudInterviewDataReportForm
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
            this.lnkDownload = new System.Windows.Forms.LinkLabel();
            this.btnPrint = new DevComponents.DotNetBar.ButtonX();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.linkUpData = new System.Windows.Forms.LinkLabel();
            this.linkViewGeDin = new System.Windows.Forms.LinkLabel();
            this.rbDEF_1 = new System.Windows.Forms.RadioButton();
            this.rbDEF_2 = new System.Windows.Forms.RadioButton();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.SuspendLayout();
            // 
            // lnkDownload
            // 
            this.lnkDownload.AutoSize = true;
            this.lnkDownload.BackColor = System.Drawing.Color.Transparent;
            this.lnkDownload.Location = new System.Drawing.Point(87, 63);
            this.lnkDownload.Name = "lnkDownload";
            this.lnkDownload.Size = new System.Drawing.Size(86, 17);
            this.lnkDownload.TabIndex = 0;
            this.lnkDownload.TabStop = true;
            this.lnkDownload.Text = "檢視預設範本";
            this.lnkDownload.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkDownload_LinkClicked);
            // 
            // btnPrint
            // 
            this.btnPrint.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnPrint.BackColor = System.Drawing.Color.Transparent;
            this.btnPrint.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnPrint.Location = new System.Drawing.Point(135, 144);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(75, 25);
            this.btnPrint.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnPrint.TabIndex = 1;
            this.btnPrint.Text = "列印";
            this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Location = new System.Drawing.Point(216, 144);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 25);
            this.btnExit.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnExit.TabIndex = 2;
            this.btnExit.Text = "離開";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // linkUpData
            // 
            this.linkUpData.AutoSize = true;
            this.linkUpData.BackColor = System.Drawing.Color.Transparent;
            this.linkUpData.Location = new System.Drawing.Point(184, 103);
            this.linkUpData.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.linkUpData.Name = "linkUpData";
            this.linkUpData.Size = new System.Drawing.Size(34, 17);
            this.linkUpData.TabIndex = 3;
            this.linkUpData.TabStop = true;
            this.linkUpData.Text = "上傳";
            this.linkUpData.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkUpData_LinkClicked);
            // 
            // linkViewGeDin
            // 
            this.linkViewGeDin.AutoSize = true;
            this.linkViewGeDin.BackColor = System.Drawing.Color.Transparent;
            this.linkViewGeDin.Location = new System.Drawing.Point(87, 103);
            this.linkViewGeDin.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.linkViewGeDin.Name = "linkViewGeDin";
            this.linkViewGeDin.Size = new System.Drawing.Size(86, 17);
            this.linkViewGeDin.TabIndex = 4;
            this.linkViewGeDin.TabStop = true;
            this.linkViewGeDin.Text = "檢視自訂範本";
            this.linkViewGeDin.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkViewGeDin_LinkClicked);
            // 
            // rbDEF_1
            // 
            this.rbDEF_1.AutoSize = true;
            this.rbDEF_1.BackColor = System.Drawing.Color.Transparent;
            this.rbDEF_1.Checked = true;
            this.rbDEF_1.Location = new System.Drawing.Point(60, 65);
            this.rbDEF_1.Name = "rbDEF_1";
            this.rbDEF_1.Size = new System.Drawing.Size(14, 13);
            this.rbDEF_1.TabIndex = 6;
            this.rbDEF_1.TabStop = true;
            this.rbDEF_1.UseVisualStyleBackColor = false;
            // 
            // rbDEF_2
            // 
            this.rbDEF_2.AutoSize = true;
            this.rbDEF_2.BackColor = System.Drawing.Color.Transparent;
            this.rbDEF_2.CausesValidation = false;
            this.rbDEF_2.Location = new System.Drawing.Point(60, 105);
            this.rbDEF_2.Margin = new System.Windows.Forms.Padding(4);
            this.rbDEF_2.Name = "rbDEF_2";
            this.rbDEF_2.Size = new System.Drawing.Size(14, 13);
            this.rbDEF_2.TabIndex = 5;
            this.rbDEF_2.UseVisualStyleBackColor = false;
            // 
            // labelX1
            // 
            this.labelX1.AutoSize = true;
            this.labelX1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.Class = "";
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.Location = new System.Drawing.Point(22, 21);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(114, 21);
            this.labelX1.TabIndex = 7;
            this.labelX1.Text = "請選擇列印設定：";
            // 
            // StudInterviewDataReportForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(303, 181);
            this.Controls.Add(this.labelX1);
            this.Controls.Add(this.rbDEF_1);
            this.Controls.Add(this.rbDEF_2);
            this.Controls.Add(this.linkUpData);
            this.Controls.Add(this.linkViewGeDin);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.lnkDownload);
            this.Controls.Add(this.btnPrint);
            this.DoubleBuffered = true;
            this.Name = "StudInterviewDataReportForm";
            this.Text = "學生輔導晤談紀錄表";
            this.Load += new System.EventHandler(this.StudInterviewDataReportForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.LinkLabel lnkDownload;
        private DevComponents.DotNetBar.ButtonX btnPrint;
        private DevComponents.DotNetBar.ButtonX btnExit;
        private System.Windows.Forms.LinkLabel linkUpData;
        private System.Windows.Forms.LinkLabel linkViewGeDin;
        private System.Windows.Forms.RadioButton rbDEF_1;
        private System.Windows.Forms.RadioButton rbDEF_2;
        private DevComponents.DotNetBar.LabelX labelX1;
    }
}