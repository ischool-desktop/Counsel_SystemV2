namespace Counsel_System2.Forms
{
    partial class ABCardQuestionsForm
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            this.itmPnlNameList = new DevComponents.DotNetBar.ItemPanel();
            this.cbxQuestionType = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.cbxControlType = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.gp1 = new DevComponents.DotNetBar.Controls.GroupPanel();
            this.groupPanel2 = new DevComponents.DotNetBar.Controls.GroupPanel();
            this.chkCanStudentEdit = new DevComponents.DotNetBar.Controls.CheckBoxX();
            this.chkCanPrint = new DevComponents.DotNetBar.Controls.CheckBoxX();
            this.chkCanTeacherEdit = new DevComponents.DotNetBar.Controls.CheckBoxX();
            this.intDisplayOrder = new DevComponents.Editors.IntegerInput();
            this.dgItems = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.colName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colHasRemark = new DevComponents.DotNetBar.Controls.DataGridViewCheckBoxXColumn();
            this.labelX3 = new DevComponents.DotNetBar.LabelX();
            this.labelX2 = new DevComponents.DotNetBar.LabelX();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.cbxGroup = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.btnSave = new DevComponents.DotNetBar.ButtonX();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.btnReset = new DevComponents.DotNetBar.ButtonX();
            this.gp1.SuspendLayout();
            this.groupPanel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.intDisplayOrder)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgItems)).BeginInit();
            this.SuspendLayout();
            // 
            // itmPnlNameList
            // 
            this.itmPnlNameList.AutoScroll = true;
            this.itmPnlNameList.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.itmPnlNameList.BackgroundStyle.Class = "ItemPanel";
            this.itmPnlNameList.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.itmPnlNameList.ContainerControlProcessDialogKey = true;
            this.itmPnlNameList.LayoutOrientation = DevComponents.DotNetBar.eOrientation.Vertical;
            this.itmPnlNameList.LicenseKey = "F962CEC7-CD8F-4911-A9E9-CAB39962FC1F";
            this.itmPnlNameList.Location = new System.Drawing.Point(10, 47);
            this.itmPnlNameList.Name = "itmPnlNameList";
            this.itmPnlNameList.Size = new System.Drawing.Size(206, 306);
            this.itmPnlNameList.TabIndex = 0;
            this.itmPnlNameList.Text = "itemPanel1";
            // 
            // cbxQuestionType
            // 
            this.cbxQuestionType.DisplayMember = "Text";
            this.cbxQuestionType.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cbxQuestionType.FormattingEnabled = true;
            this.cbxQuestionType.ItemHeight = 19;
            this.cbxQuestionType.Location = new System.Drawing.Point(309, 434);
            this.cbxQuestionType.Name = "cbxQuestionType";
            this.cbxQuestionType.Size = new System.Drawing.Size(157, 25);
            this.cbxQuestionType.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbxQuestionType.TabIndex = 1;
            // 
            // cbxControlType
            // 
            this.cbxControlType.DisplayMember = "Text";
            this.cbxControlType.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cbxControlType.FormattingEnabled = true;
            this.cbxControlType.ItemHeight = 19;
            this.cbxControlType.Location = new System.Drawing.Point(309, 467);
            this.cbxControlType.Name = "cbxControlType";
            this.cbxControlType.Size = new System.Drawing.Size(157, 25);
            this.cbxControlType.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbxControlType.TabIndex = 2;
            // 
            // gp1
            // 
            this.gp1.BackColor = System.Drawing.Color.Transparent;
            this.gp1.CanvasColor = System.Drawing.SystemColors.Control;
            this.gp1.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007;
            this.gp1.Controls.Add(this.groupPanel2);
            this.gp1.Controls.Add(this.intDisplayOrder);
            this.gp1.Controls.Add(this.dgItems);
            this.gp1.Controls.Add(this.labelX3);
            this.gp1.Location = new System.Drawing.Point(225, 16);
            this.gp1.Name = "gp1";
            this.gp1.Size = new System.Drawing.Size(428, 337);
            // 
            // 
            // 
            this.gp1.Style.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground2;
            this.gp1.Style.BackColorGradientAngle = 90;
            this.gp1.Style.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground;
            this.gp1.Style.BorderBottom = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.gp1.Style.BorderBottomWidth = 1;
            this.gp1.Style.BorderColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBorder;
            this.gp1.Style.BorderLeft = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.gp1.Style.BorderLeftWidth = 1;
            this.gp1.Style.BorderRight = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.gp1.Style.BorderRightWidth = 1;
            this.gp1.Style.BorderTop = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.gp1.Style.BorderTopWidth = 1;
            this.gp1.Style.Class = "";
            this.gp1.Style.CornerDiameter = 4;
            this.gp1.Style.CornerType = DevComponents.DotNetBar.eCornerType.Rounded;
            this.gp1.Style.Font = new System.Drawing.Font("Microsoft JhengHei", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.gp1.Style.TextColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelText;
            this.gp1.Style.TextLineAlignment = DevComponents.DotNetBar.eStyleTextAlignment.Near;
            // 
            // 
            // 
            this.gp1.StyleMouseDown.Class = "";
            this.gp1.StyleMouseDown.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            // 
            // 
            // 
            this.gp1.StyleMouseOver.Class = "";
            this.gp1.StyleMouseOver.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.gp1.TabIndex = 6;
            this.gp1.Text = "名稱";
            // 
            // groupPanel2
            // 
            this.groupPanel2.CanvasColor = System.Drawing.SystemColors.Control;
            this.groupPanel2.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007;
            this.groupPanel2.Controls.Add(this.chkCanStudentEdit);
            this.groupPanel2.Controls.Add(this.chkCanPrint);
            this.groupPanel2.Controls.Add(this.chkCanTeacherEdit);
            this.groupPanel2.Location = new System.Drawing.Point(11, 7);
            this.groupPanel2.Name = "groupPanel2";
            this.groupPanel2.Size = new System.Drawing.Size(281, 47);
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
            this.groupPanel2.Style.TextAlignment = DevComponents.DotNetBar.eStyleTextAlignment.Center;
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
            this.groupPanel2.TabIndex = 16;
            // 
            // chkCanStudentEdit
            // 
            this.chkCanStudentEdit.AutoSize = true;
            // 
            // 
            // 
            this.chkCanStudentEdit.BackgroundStyle.Class = "";
            this.chkCanStudentEdit.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.chkCanStudentEdit.Location = new System.Drawing.Point(172, 10);
            this.chkCanStudentEdit.Name = "chkCanStudentEdit";
            this.chkCanStudentEdit.Size = new System.Drawing.Size(94, 21);
            this.chkCanStudentEdit.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.chkCanStudentEdit.TabIndex = 14;
            this.chkCanStudentEdit.Text = "學生可編輯";
            // 
            // chkCanPrint
            // 
            this.chkCanPrint.AutoSize = true;
            // 
            // 
            // 
            this.chkCanPrint.BackgroundStyle.Class = "";
            this.chkCanPrint.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.chkCanPrint.Location = new System.Drawing.Point(9, 10);
            this.chkCanPrint.Name = "chkCanPrint";
            this.chkCanPrint.Size = new System.Drawing.Size(67, 21);
            this.chkCanPrint.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.chkCanPrint.TabIndex = 12;
            this.chkCanPrint.Text = "可列印";
            // 
            // chkCanTeacherEdit
            // 
            this.chkCanTeacherEdit.AutoSize = true;
            // 
            // 
            // 
            this.chkCanTeacherEdit.BackgroundStyle.Class = "";
            this.chkCanTeacherEdit.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.chkCanTeacherEdit.Location = new System.Drawing.Point(77, 10);
            this.chkCanTeacherEdit.Name = "chkCanTeacherEdit";
            this.chkCanTeacherEdit.Size = new System.Drawing.Size(94, 21);
            this.chkCanTeacherEdit.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.chkCanTeacherEdit.TabIndex = 13;
            this.chkCanTeacherEdit.Text = "教師可編輯";
            // 
            // intDisplayOrder
            // 
            // 
            // 
            // 
            this.intDisplayOrder.BackgroundStyle.Class = "DateTimeInputBackground";
            this.intDisplayOrder.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.intDisplayOrder.ButtonFreeText.Shortcut = DevComponents.DotNetBar.eShortcut.F2;
            this.intDisplayOrder.Location = new System.Drawing.Point(358, 17);
            this.intDisplayOrder.MinValue = 0;
            this.intDisplayOrder.Name = "intDisplayOrder";
            this.intDisplayOrder.ShowUpDown = true;
            this.intDisplayOrder.Size = new System.Drawing.Size(58, 25);
            this.intDisplayOrder.TabIndex = 15;
            this.intDisplayOrder.Visible = false;
            // 
            // dgItems
            // 
            this.dgItems.AllowUserToResizeRows = false;
            this.dgItems.BackgroundColor = System.Drawing.SystemColors.Window;
            this.dgItems.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgItems.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colName,
            this.colHasRemark});
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgItems.DefaultCellStyle = dataGridViewCellStyle4;
            this.dgItems.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.dgItems.Location = new System.Drawing.Point(8, 63);
            this.dgItems.Name = "dgItems";
            this.dgItems.RowTemplate.Height = 24;
            this.dgItems.Size = new System.Drawing.Size(408, 237);
            this.dgItems.TabIndex = 6;
            // 
            // colName
            // 
            this.colName.HeaderText = "名稱";
            this.colName.Name = "colName";
            this.colName.Width = 250;
            // 
            // colHasRemark
            // 
            this.colHasRemark.Checked = true;
            this.colHasRemark.CheckState = System.Windows.Forms.CheckState.Indeterminate;
            this.colHasRemark.CheckValue = null;
            this.colHasRemark.HeaderText = "可備註";
            this.colHasRemark.Name = "colHasRemark";
            this.colHasRemark.Width = 90;
            // 
            // labelX3
            // 
            this.labelX3.AutoSize = true;
            // 
            // 
            // 
            this.labelX3.BackgroundStyle.Class = "";
            this.labelX3.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX3.Location = new System.Drawing.Point(298, 19);
            this.labelX3.Name = "labelX3";
            this.labelX3.Size = new System.Drawing.Size(60, 21);
            this.labelX3.TabIndex = 9;
            this.labelX3.Text = "顯示順序";
            this.labelX3.Visible = false;
            // 
            // labelX2
            // 
            this.labelX2.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX2.BackgroundStyle.Class = "";
            this.labelX2.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX2.Location = new System.Drawing.Point(227, 467);
            this.labelX2.Name = "labelX2";
            this.labelX2.Size = new System.Drawing.Size(75, 23);
            this.labelX2.TabIndex = 8;
            this.labelX2.Text = "控制項類型";
            this.labelX2.TextAlignment = System.Drawing.StringAlignment.Far;
            // 
            // labelX1
            // 
            this.labelX1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.Class = "";
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.Location = new System.Drawing.Point(227, 434);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(75, 23);
            this.labelX1.TabIndex = 7;
            this.labelX1.Text = "問題類型";
            this.labelX1.TextAlignment = System.Drawing.StringAlignment.Far;
            // 
            // cbxGroup
            // 
            this.cbxGroup.DisplayMember = "Text";
            this.cbxGroup.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cbxGroup.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxGroup.FormattingEnabled = true;
            this.cbxGroup.ItemHeight = 19;
            this.cbxGroup.Location = new System.Drawing.Point(10, 16);
            this.cbxGroup.Name = "cbxGroup";
            this.cbxGroup.Size = new System.Drawing.Size(206, 25);
            this.cbxGroup.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbxGroup.TabIndex = 7;
            this.cbxGroup.WatermarkText = "請選擇群組..";
            this.cbxGroup.SelectedIndexChanged += new System.EventHandler(this.cbxGroup_SelectedIndexChanged);
            // 
            // btnSave
            // 
            this.btnSave.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnSave.BackColor = System.Drawing.Color.Transparent;
            this.btnSave.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnSave.Location = new System.Drawing.Point(488, 364);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 25);
            this.btnSave.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnSave.TabIndex = 8;
            this.btnSave.Text = "儲存";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Location = new System.Drawing.Point(569, 364);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 25);
            this.btnExit.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnExit.TabIndex = 9;
            this.btnExit.Text = "離開";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // btnReset
            // 
            this.btnReset.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnReset.AutoSize = true;
            this.btnReset.BackColor = System.Drawing.Color.Transparent;
            this.btnReset.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnReset.Location = new System.Drawing.Point(12, 364);
            this.btnReset.Name = "btnReset";
            this.btnReset.Size = new System.Drawing.Size(91, 25);
            this.btnReset.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnReset.TabIndex = 11;
            this.btnReset.Text = "全部重設";
            this.btnReset.Click += new System.EventHandler(this.btnReset_Click);
            // 
            // ABCardQuestionsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(658, 397);
            this.Controls.Add(this.btnReset);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.labelX2);
            this.Controls.Add(this.cbxGroup);
            this.Controls.Add(this.labelX1);
            this.Controls.Add(this.cbxQuestionType);
            this.Controls.Add(this.gp1);
            this.Controls.Add(this.itmPnlNameList);
            this.Controls.Add(this.cbxControlType);
            this.DoubleBuffered = true;
            this.Name = "ABCardQuestionsForm";
            this.Text = "綜合表現紀錄表題目設定";
            this.Load += new System.EventHandler(this.ABCardQuestionsForm_Load);
            this.gp1.ResumeLayout(false);
            this.gp1.PerformLayout();
            this.groupPanel2.ResumeLayout(false);
            this.groupPanel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.intDisplayOrder)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgItems)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.ItemPanel itmPnlNameList;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cbxQuestionType;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cbxControlType;
        private DevComponents.DotNetBar.Controls.GroupPanel gp1;
        private DevComponents.DotNetBar.LabelX labelX2;
        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.Controls.DataGridViewX dgItems;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cbxGroup;
        private DevComponents.DotNetBar.ButtonX btnSave;
        private DevComponents.DotNetBar.ButtonX btnExit;
        private DevComponents.DotNetBar.LabelX labelX3;
        private DevComponents.DotNetBar.Controls.CheckBoxX chkCanStudentEdit;
        private DevComponents.DotNetBar.Controls.CheckBoxX chkCanTeacherEdit;
        private DevComponents.DotNetBar.Controls.CheckBoxX chkCanPrint;
        private DevComponents.DotNetBar.Controls.GroupPanel groupPanel2;
        private DevComponents.Editors.IntegerInput intDisplayOrder;
        private System.Windows.Forms.DataGridViewTextBoxColumn colName;
        private DevComponents.DotNetBar.Controls.DataGridViewCheckBoxXColumn colHasRemark;
        private DevComponents.DotNetBar.ButtonX btnReset;
    }
}