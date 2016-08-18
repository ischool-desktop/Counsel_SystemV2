namespace Counsel_System2.Forms
{
    partial class StudQuizDataForm
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.lblImpName = new DevComponents.DotNetBar.LabelX();
            this.labelX3 = new DevComponents.DotNetBar.LabelX();
            this.dgQuizData = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.colDataField = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colDataValue = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnSave = new DevComponents.DotNetBar.ButtonX();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.dtImplementationDate = new DevComponents.Editors.DateTimeAdv.DateTimeInput();
            this.dtAnalysisDate = new DevComponents.Editors.DateTimeAdv.DateTimeInput();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.cbxQuizName = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            ((System.ComponentModel.ISupportInitialize)(this.dgQuizData)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dtImplementationDate)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dtAnalysisDate)).BeginInit();
            this.SuspendLayout();
            // 
            // lblImpName
            // 
            this.lblImpName.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lblImpName.BackgroundStyle.Class = "";
            this.lblImpName.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblImpName.Location = new System.Drawing.Point(13, 54);
            this.lblImpName.Name = "lblImpName";
            this.lblImpName.Size = new System.Drawing.Size(58, 23);
            this.lblImpName.TabIndex = 1;
            this.lblImpName.Text = "實施日期";
            // 
            // labelX3
            // 
            this.labelX3.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX3.BackgroundStyle.Class = "";
            this.labelX3.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX3.Location = new System.Drawing.Point(272, 54);
            this.labelX3.Name = "labelX3";
            this.labelX3.Size = new System.Drawing.Size(70, 23);
            this.labelX3.TabIndex = 3;
            this.labelX3.Text = "解析日期";
            // 
            // dgQuizData
            // 
            this.dgQuizData.AllowUserToAddRows = false;
            this.dgQuizData.AllowUserToDeleteRows = false;
            this.dgQuizData.BackgroundColor = System.Drawing.Color.White;
            this.dgQuizData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgQuizData.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colDataField,
            this.colDataValue});
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgQuizData.DefaultCellStyle = dataGridViewCellStyle1;
            this.dgQuizData.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.dgQuizData.Location = new System.Drawing.Point(13, 91);
            this.dgQuizData.Name = "dgQuizData";
            this.dgQuizData.RowTemplate.Height = 24;
            this.dgQuizData.Size = new System.Drawing.Size(448, 183);
            this.dgQuizData.TabIndex = 3;
            this.dgQuizData.CurrentCellDirtyStateChanged += new System.EventHandler(this.dgQuizData_CurrentCellDirtyStateChanged);
            // 
            // colDataField
            // 
            this.colDataField.HeaderText = "項目名稱";
            this.colDataField.Name = "colDataField";
            this.colDataField.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // colDataValue
            // 
            this.colDataValue.HeaderText = "測驗結果";
            this.colDataValue.Name = "colDataValue";
            this.colDataValue.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.colDataValue.Width = 300;
            // 
            // btnSave
            // 
            this.btnSave.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnSave.BackColor = System.Drawing.Color.Transparent;
            this.btnSave.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnSave.Location = new System.Drawing.Point(302, 282);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 4;
            this.btnSave.Text = "儲存";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Location = new System.Drawing.Point(386, 282);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 23);
            this.btnExit.TabIndex = 5;
            this.btnExit.Text = "離開";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // dtImplementationDate
            // 
            this.dtImplementationDate.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.dtImplementationDate.BackgroundStyle.Class = "DateTimeInputBackground";
            this.dtImplementationDate.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.dtImplementationDate.ButtonDropDown.Shortcut = DevComponents.DotNetBar.eShortcut.AltDown;
            this.dtImplementationDate.ButtonDropDown.Visible = true;
            this.dtImplementationDate.IsPopupCalendarOpen = false;
            this.dtImplementationDate.Location = new System.Drawing.Point(75, 52);
            // 
            // 
            // 
            this.dtImplementationDate.MonthCalendar.AnnuallyMarkedDates = new System.DateTime[0];
            // 
            // 
            // 
            this.dtImplementationDate.MonthCalendar.BackgroundStyle.BackColor = System.Drawing.SystemColors.Window;
            this.dtImplementationDate.MonthCalendar.BackgroundStyle.Class = "";
            this.dtImplementationDate.MonthCalendar.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.dtImplementationDate.MonthCalendar.ClearButtonVisible = true;
            // 
            // 
            // 
            this.dtImplementationDate.MonthCalendar.CommandsBackgroundStyle.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground2;
            this.dtImplementationDate.MonthCalendar.CommandsBackgroundStyle.BackColorGradientAngle = 90;
            this.dtImplementationDate.MonthCalendar.CommandsBackgroundStyle.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground;
            this.dtImplementationDate.MonthCalendar.CommandsBackgroundStyle.BorderTop = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.dtImplementationDate.MonthCalendar.CommandsBackgroundStyle.BorderTopColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarDockedBorder;
            this.dtImplementationDate.MonthCalendar.CommandsBackgroundStyle.BorderTopWidth = 1;
            this.dtImplementationDate.MonthCalendar.CommandsBackgroundStyle.Class = "";
            this.dtImplementationDate.MonthCalendar.CommandsBackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.dtImplementationDate.MonthCalendar.DayNames = new string[] {
        "日",
        "一",
        "二",
        "三",
        "四",
        "五",
        "六"};
            this.dtImplementationDate.MonthCalendar.DisplayMonth = new System.DateTime(2011, 5, 1, 0, 0, 0, 0);
            this.dtImplementationDate.MonthCalendar.MarkedDates = new System.DateTime[0];
            this.dtImplementationDate.MonthCalendar.MonthlyMarkedDates = new System.DateTime[0];
            // 
            // 
            // 
            this.dtImplementationDate.MonthCalendar.NavigationBackgroundStyle.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground2;
            this.dtImplementationDate.MonthCalendar.NavigationBackgroundStyle.BackColorGradientAngle = 90;
            this.dtImplementationDate.MonthCalendar.NavigationBackgroundStyle.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground;
            this.dtImplementationDate.MonthCalendar.NavigationBackgroundStyle.Class = "";
            this.dtImplementationDate.MonthCalendar.NavigationBackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.dtImplementationDate.MonthCalendar.TodayButtonVisible = true;
            this.dtImplementationDate.MonthCalendar.WeeklyMarkedDays = new System.DayOfWeek[0];
            this.dtImplementationDate.Name = "dtImplementationDate";
            this.dtImplementationDate.Size = new System.Drawing.Size(131, 25);
            this.dtImplementationDate.TabIndex = 1;
            // 
            // dtAnalysisDate
            // 
            this.dtAnalysisDate.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.dtAnalysisDate.BackgroundStyle.Class = "DateTimeInputBackground";
            this.dtAnalysisDate.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.dtAnalysisDate.ButtonDropDown.Shortcut = DevComponents.DotNetBar.eShortcut.AltDown;
            this.dtAnalysisDate.ButtonDropDown.Visible = true;
            this.dtAnalysisDate.IsPopupCalendarOpen = false;
            this.dtAnalysisDate.Location = new System.Drawing.Point(330, 52);
            // 
            // 
            // 
            this.dtAnalysisDate.MonthCalendar.AnnuallyMarkedDates = new System.DateTime[0];
            // 
            // 
            // 
            this.dtAnalysisDate.MonthCalendar.BackgroundStyle.BackColor = System.Drawing.SystemColors.Window;
            this.dtAnalysisDate.MonthCalendar.BackgroundStyle.Class = "";
            this.dtAnalysisDate.MonthCalendar.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.dtAnalysisDate.MonthCalendar.ClearButtonVisible = true;
            // 
            // 
            // 
            this.dtAnalysisDate.MonthCalendar.CommandsBackgroundStyle.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground2;
            this.dtAnalysisDate.MonthCalendar.CommandsBackgroundStyle.BackColorGradientAngle = 90;
            this.dtAnalysisDate.MonthCalendar.CommandsBackgroundStyle.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground;
            this.dtAnalysisDate.MonthCalendar.CommandsBackgroundStyle.BorderTop = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.dtAnalysisDate.MonthCalendar.CommandsBackgroundStyle.BorderTopColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarDockedBorder;
            this.dtAnalysisDate.MonthCalendar.CommandsBackgroundStyle.BorderTopWidth = 1;
            this.dtAnalysisDate.MonthCalendar.CommandsBackgroundStyle.Class = "";
            this.dtAnalysisDate.MonthCalendar.CommandsBackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.dtAnalysisDate.MonthCalendar.DayNames = new string[] {
        "日",
        "一",
        "二",
        "三",
        "四",
        "五",
        "六"};
            this.dtAnalysisDate.MonthCalendar.DisplayMonth = new System.DateTime(2011, 5, 1, 0, 0, 0, 0);
            this.dtAnalysisDate.MonthCalendar.MarkedDates = new System.DateTime[0];
            this.dtAnalysisDate.MonthCalendar.MonthlyMarkedDates = new System.DateTime[0];
            // 
            // 
            // 
            this.dtAnalysisDate.MonthCalendar.NavigationBackgroundStyle.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground2;
            this.dtAnalysisDate.MonthCalendar.NavigationBackgroundStyle.BackColorGradientAngle = 90;
            this.dtAnalysisDate.MonthCalendar.NavigationBackgroundStyle.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground;
            this.dtAnalysisDate.MonthCalendar.NavigationBackgroundStyle.Class = "";
            this.dtAnalysisDate.MonthCalendar.NavigationBackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.dtAnalysisDate.MonthCalendar.TodayButtonVisible = true;
            this.dtAnalysisDate.MonthCalendar.WeeklyMarkedDays = new System.DayOfWeek[0];
            this.dtAnalysisDate.Name = "dtAnalysisDate";
            this.dtAnalysisDate.Size = new System.Drawing.Size(131, 25);
            this.dtAnalysisDate.TabIndex = 2;
            // 
            // labelX1
            // 
            this.labelX1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.Class = "";
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.Location = new System.Drawing.Point(13, 15);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(58, 23);
            this.labelX1.TabIndex = 6;
            this.labelX1.Text = "測驗名稱";
            // 
            // cbxQuizName
            // 
            this.cbxQuizName.DisplayMember = "Text";
            this.cbxQuizName.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cbxQuizName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxQuizName.FormattingEnabled = true;
            this.cbxQuizName.ItemHeight = 19;
            this.cbxQuizName.Location = new System.Drawing.Point(75, 13);
            this.cbxQuizName.Name = "cbxQuizName";
            this.cbxQuizName.Size = new System.Drawing.Size(386, 25);
            this.cbxQuizName.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbxQuizName.TabIndex = 7;
            this.cbxQuizName.SelectedIndexChanged += new System.EventHandler(this.cbxQuizName_SelectedIndexChanged);
            // 
            // StudQuizDataForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(470, 309);
            this.Controls.Add(this.cbxQuizName);
            this.Controls.Add(this.labelX1);
            this.Controls.Add(this.dtAnalysisDate);
            this.Controls.Add(this.dtImplementationDate);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.dgQuizData);
            this.Controls.Add(this.labelX3);
            this.Controls.Add(this.lblImpName);
            this.DoubleBuffered = true;
            this.Name = "StudQuizDataForm";
            this.Text = "學生測驗內容";
            this.Load += new System.EventHandler(this.StudQuizDataForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgQuizData)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dtImplementationDate)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dtAnalysisDate)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.LabelX lblImpName;
        private DevComponents.DotNetBar.LabelX labelX3;
        private DevComponents.DotNetBar.Controls.DataGridViewX dgQuizData;
        private DevComponents.DotNetBar.ButtonX btnSave;
        private DevComponents.DotNetBar.ButtonX btnExit;
        private DevComponents.Editors.DateTimeAdv.DateTimeInput dtImplementationDate;
        private DevComponents.Editors.DateTimeAdv.DateTimeInput dtAnalysisDate;
        private System.Windows.Forms.DataGridViewTextBoxColumn colDataField;
        private System.Windows.Forms.DataGridViewTextBoxColumn colDataValue;
        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cbxQuizName;
    }
}