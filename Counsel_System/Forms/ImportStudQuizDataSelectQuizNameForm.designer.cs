namespace Counsel_System2.Forms
{
    partial class ImportStudQuizDataSelectQuizNameForm
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
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.cbxQuizName = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.btnRun = new DevComponents.DotNetBar.ButtonX();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.groupPanel1 = new DevComponents.DotNetBar.Controls.GroupPanel();
            this.rbClassSeatNo = new System.Windows.Forms.RadioButton();
            this.rbStudentNumber = new System.Windows.Forms.RadioButton();
            this.groupPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // labelX1
            // 
            this.labelX1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.Class = "";
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.Location = new System.Drawing.Point(14, 14);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(63, 30);
            this.labelX1.TabIndex = 0;
            this.labelX1.Text = "測驗名稱";
            // 
            // cbxQuizName
            // 
            this.cbxQuizName.DisplayMember = "Text";
            this.cbxQuizName.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cbxQuizName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxQuizName.FormattingEnabled = true;
            this.cbxQuizName.ItemHeight = 19;
            this.cbxQuizName.Location = new System.Drawing.Point(81, 17);
            this.cbxQuizName.Name = "cbxQuizName";
            this.cbxQuizName.Size = new System.Drawing.Size(327, 25);
            this.cbxQuizName.TabIndex = 1;
            // 
            // btnRun
            // 
            this.btnRun.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnRun.BackColor = System.Drawing.Color.Transparent;
            this.btnRun.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnRun.Location = new System.Drawing.Point(261, 88);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(67, 23);
            this.btnRun.TabIndex = 2;
            this.btnRun.Text = "確定";
            this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Location = new System.Drawing.Point(340, 88);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(67, 23);
            this.btnExit.TabIndex = 3;
            this.btnExit.Text = "離開";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // groupPanel1
            // 
            this.groupPanel1.BackColor = System.Drawing.Color.Transparent;
            this.groupPanel1.CanvasColor = System.Drawing.SystemColors.Control;
            this.groupPanel1.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007;
            this.groupPanel1.Controls.Add(this.rbClassSeatNo);
            this.groupPanel1.Controls.Add(this.rbStudentNumber);
            this.groupPanel1.Location = new System.Drawing.Point(10, 52);
            this.groupPanel1.Name = "groupPanel1";
            this.groupPanel1.Size = new System.Drawing.Size(200, 59);
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
            this.groupPanel1.TabIndex = 4;
            this.groupPanel1.Text = "匯入資料依照方式";
            // 
            // rbClassSeatNo
            // 
            this.rbClassSeatNo.AutoSize = true;
            this.rbClassSeatNo.Location = new System.Drawing.Point(95, 7);
            this.rbClassSeatNo.Name = "rbClassSeatNo";
            this.rbClassSeatNo.Size = new System.Drawing.Size(78, 21);
            this.rbClassSeatNo.TabIndex = 1;
            this.rbClassSeatNo.TabStop = true;
            this.rbClassSeatNo.Text = "班級座號";
            this.rbClassSeatNo.UseVisualStyleBackColor = true;
            // 
            // rbStudentNumber
            // 
            this.rbStudentNumber.AutoSize = true;
            this.rbStudentNumber.Location = new System.Drawing.Point(13, 7);
            this.rbStudentNumber.Name = "rbStudentNumber";
            this.rbStudentNumber.Size = new System.Drawing.Size(52, 21);
            this.rbStudentNumber.TabIndex = 0;
            this.rbStudentNumber.TabStop = true;
            this.rbStudentNumber.Text = "學號";
            this.rbStudentNumber.UseVisualStyleBackColor = true;
            // 
            // ImportStudQuizDataSelectQuizNameForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(421, 117);
            this.Controls.Add(this.groupPanel1);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnRun);
            this.Controls.Add(this.cbxQuizName);
            this.Controls.Add(this.labelX1);
            this.DoubleBuffered = true;
            this.Name = "ImportStudQuizDataSelectQuizNameForm";
            this.Text = "請選擇匯入測驗";
            this.Load += new System.EventHandler(this.ImportStudQuizDataSelectQuizNameForm_Load);
            this.groupPanel1.ResumeLayout(false);
            this.groupPanel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cbxQuizName;
        private DevComponents.DotNetBar.ButtonX btnRun;
        private DevComponents.DotNetBar.ButtonX btnExit;
        private DevComponents.DotNetBar.Controls.GroupPanel groupPanel1;
        private System.Windows.Forms.RadioButton rbClassSeatNo;
        private System.Windows.Forms.RadioButton rbStudentNumber;
    }
}