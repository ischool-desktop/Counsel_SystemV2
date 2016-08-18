namespace Counsel_System2.Forms
{
    partial class SetABCardAccessStartingDate
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
            this.btnSave = new DevComponents.DotNetBar.ButtonX();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.dgStatEndDateTime = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.colGradeYear = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colStartDateTime = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colEndDateTime = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.lblMsg = new DevComponents.DotNetBar.LabelX();
            ((System.ComponentModel.ISupportInitialize)(this.dgStatEndDateTime)).BeginInit();
            this.SuspendLayout();
            // 
            // btnSave
            // 
            this.btnSave.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnSave.BackColor = System.Drawing.Color.Transparent;
            this.btnSave.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnSave.Location = new System.Drawing.Point(202, 237);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 25);
            this.btnSave.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnSave.TabIndex = 4;
            this.btnSave.Text = "儲存";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Location = new System.Drawing.Point(284, 237);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 25);
            this.btnExit.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnExit.TabIndex = 5;
            this.btnExit.Text = "離開";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // dgStatEndDateTime
            // 
            this.dgStatEndDateTime.AllowUserToAddRows = false;
            this.dgStatEndDateTime.AllowUserToDeleteRows = false;
            this.dgStatEndDateTime.BackgroundColor = System.Drawing.Color.White;
            this.dgStatEndDateTime.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgStatEndDateTime.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colGradeYear,
            this.colStartDateTime,
            this.colEndDateTime});
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgStatEndDateTime.DefaultCellStyle = dataGridViewCellStyle1;
            this.dgStatEndDateTime.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.dgStatEndDateTime.Location = new System.Drawing.Point(11, 36);
            this.dgStatEndDateTime.Name = "dgStatEndDateTime";
            this.dgStatEndDateTime.RowTemplate.Height = 24;
            this.dgStatEndDateTime.Size = new System.Drawing.Size(348, 193);
            this.dgStatEndDateTime.TabIndex = 6;
            this.dgStatEndDateTime.CurrentCellDirtyStateChanged += new System.EventHandler(this.dgStatEndDateTime_CurrentCellDirtyStateChanged);
            // 
            // colGradeYear
            // 
            this.colGradeYear.HeaderText = "年級";
            this.colGradeYear.Name = "colGradeYear";
            this.colGradeYear.Width = 70;
            // 
            // colStartDateTime
            // 
            this.colStartDateTime.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.colStartDateTime.HeaderText = "開始日期時間";
            this.colStartDateTime.Name = "colStartDateTime";
            this.colStartDateTime.Width = 111;
            // 
            // colEndDateTime
            // 
            this.colEndDateTime.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.colEndDateTime.HeaderText = "結束日期時間";
            this.colEndDateTime.Name = "colEndDateTime";
            this.colEndDateTime.Width = 111;
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
            this.lblMsg.Location = new System.Drawing.Point(11, 7);
            this.lblMsg.Name = "lblMsg";
            this.lblMsg.Size = new System.Drawing.Size(347, 23);
            this.lblMsg.TabIndex = 7;
            this.lblMsg.Text = "資料讀取中...";
            // 
            // SetABCardAccessStartingDate
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(371, 268);
            this.Controls.Add(this.lblMsg);
            this.Controls.Add(this.dgStatEndDateTime);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnSave);
            this.DoubleBuffered = true;
            this.Name = "SetABCardAccessStartingDate";
            this.Text = "設定綜合資料表輸入開放時間";
            this.Load += new System.EventHandler(this.SetABCardAccessStartingDate_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgStatEndDateTime)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.ButtonX btnSave;
        private DevComponents.DotNetBar.ButtonX btnExit;
        private DevComponents.DotNetBar.Controls.DataGridViewX dgStatEndDateTime;
        private System.Windows.Forms.DataGridViewTextBoxColumn colGradeYear;
        private System.Windows.Forms.DataGridViewTextBoxColumn colStartDateTime;
        private System.Windows.Forms.DataGridViewTextBoxColumn colEndDateTime;
        private DevComponents.DotNetBar.LabelX lblMsg;
    }
}