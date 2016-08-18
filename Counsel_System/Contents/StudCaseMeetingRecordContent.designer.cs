namespace Counsel_System2.Contents
{
    partial class StudCaseMeetingRecordContent
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.lvCaseMeeting = new DevComponents.DotNetBar.Controls.ListViewEx();
            this.colMeetngDate = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colAuthorID = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colMeetingCause = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.btnInsert = new DevComponents.DotNetBar.ButtonX();
            this.btnEdit = new DevComponents.DotNetBar.ButtonX();
            this.btnDelete = new DevComponents.DotNetBar.ButtonX();
            this.SuspendLayout();
            // 
            // lvCaseMeeting
            // 
            // 
            // 
            // 
            this.lvCaseMeeting.Border.Class = "ListViewBorder";
            this.lvCaseMeeting.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lvCaseMeeting.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colMeetngDate,
            this.colAuthorID,
            this.colMeetingCause});
            this.lvCaseMeeting.FullRowSelect = true;
            this.lvCaseMeeting.Location = new System.Drawing.Point(15, 12);
            this.lvCaseMeeting.MultiSelect = false;
            this.lvCaseMeeting.Name = "lvCaseMeeting";
            this.lvCaseMeeting.Size = new System.Drawing.Size(520, 186);
            this.lvCaseMeeting.TabIndex = 0;
            this.lvCaseMeeting.UseCompatibleStateImageBehavior = false;
            this.lvCaseMeeting.View = System.Windows.Forms.View.Details;
            // 
            // colMeetngDate
            // 
            this.colMeetngDate.Text = "會議日期";
            this.colMeetngDate.Width = 100;
            // 
            // colAuthorID
            // 
            this.colAuthorID.Text = "記錄者";
            this.colAuthorID.Width = 100;
            // 
            // colMeetingCause
            // 
            this.colMeetingCause.Text = "會議事由";
            this.colMeetingCause.Width = 200;
            // 
            // btnInsert
            // 
            this.btnInsert.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnInsert.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnInsert.Location = new System.Drawing.Point(15, 205);
            this.btnInsert.Name = "btnInsert";
            this.btnInsert.Size = new System.Drawing.Size(75, 25);
            this.btnInsert.TabIndex = 1;
            this.btnInsert.Text = "新增";
            this.btnInsert.Click += new System.EventHandler(this.btnInsert_Click);
            // 
            // btnEdit
            // 
            this.btnEdit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnEdit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnEdit.Location = new System.Drawing.Point(100, 205);
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.Size = new System.Drawing.Size(75, 25);
            this.btnEdit.TabIndex = 2;
            this.btnEdit.Text = "修改";
            this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnDelete.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnDelete.Location = new System.Drawing.Point(185, 205);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(75, 25);
            this.btnDelete.TabIndex = 3;
            this.btnDelete.Text = "刪除";
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // StudCaseMeetingRecordContent
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btnDelete);
            this.Controls.Add(this.btnEdit);
            this.Controls.Add(this.btnInsert);
            this.Controls.Add(this.lvCaseMeeting);
            this.Name = "StudCaseMeetingRecordContent";
            this.Size = new System.Drawing.Size(550, 240);
            this.Load += new System.EventHandler(this.StudCaseMeetingRecordContent_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private  DevComponents.DotNetBar.Controls.ListViewEx lvCaseMeeting;
        private DevComponents.DotNetBar.ButtonX btnInsert;
        private DevComponents.DotNetBar.ButtonX btnEdit;
        private DevComponents.DotNetBar.ButtonX btnDelete;
        private System.Windows.Forms.ColumnHeader colMeetngDate;
        private System.Windows.Forms.ColumnHeader colAuthorID;
        private System.Windows.Forms.ColumnHeader colMeetingCause;
    }
}
