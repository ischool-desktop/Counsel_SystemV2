namespace Counsel_System2.Contents
{
    partial class CounselStudentListAContent
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
            this.lvStudentList = new DevComponents.DotNetBar.Controls.ListViewEx();
            this.colClassName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colSeatNo = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colStudentNumber = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colStatus = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.btnRemoveStudent = new DevComponents.DotNetBar.ButtonX();
            this.btnAddTempStudent = new DevComponents.DotNetBar.ButtonX();
            this.buttonItem2 = new DevComponents.DotNetBar.ButtonItem();
            this.lblStudentCount = new DevComponents.DotNetBar.LabelX();
            this.SuspendLayout();
            // 
            // lvStudentList
            // 
            // 
            // 
            // 
            this.lvStudentList.Border.Class = "ListViewBorder";
            this.lvStudentList.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lvStudentList.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colClassName,
            this.colSeatNo,
            this.colName,
            this.colStudentNumber,
            this.colStatus});
            this.lvStudentList.FullRowSelect = true;
            this.lvStudentList.Location = new System.Drawing.Point(15, 12);
            this.lvStudentList.Name = "lvStudentList";
            this.lvStudentList.Size = new System.Drawing.Size(520, 186);
            this.lvStudentList.TabIndex = 0;
            this.lvStudentList.UseCompatibleStateImageBehavior = false;
            this.lvStudentList.View = System.Windows.Forms.View.Details;
            // 
            // colClassName
            // 
            this.colClassName.Text = "班級";
            this.colClassName.Width = 80;
            // 
            // colSeatNo
            // 
            this.colSeatNo.Text = "座號";
            // 
            // colName
            // 
            this.colName.Text = "姓名";
            this.colName.Width = 100;
            // 
            // colStudentNumber
            // 
            this.colStudentNumber.Text = "學號";
            this.colStudentNumber.Width = 100;
            // 
            // colStatus
            // 
            this.colStatus.Text = "學生狀態";
            this.colStatus.Width = 80;
            // 
            // btnRemoveStudent
            // 
            this.btnRemoveStudent.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnRemoveStudent.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnRemoveStudent.Location = new System.Drawing.Point(16, 207);
            this.btnRemoveStudent.Name = "btnRemoveStudent";
            this.btnRemoveStudent.Size = new System.Drawing.Size(99, 23);
            this.btnRemoveStudent.TabIndex = 1;
            this.btnRemoveStudent.Text = "移除所選學生";
            this.btnRemoveStudent.Click += new System.EventHandler(this.btnRemoveStudent_Click);
            // 
            // btnAddTempStudent
            // 
            this.btnAddTempStudent.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnAddTempStudent.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnAddTempStudent.Location = new System.Drawing.Point(133, 207);
            this.btnAddTempStudent.Name = "btnAddTempStudent";
            this.btnAddTempStudent.Size = new System.Drawing.Size(112, 23);
            this.btnAddTempStudent.SubItems.AddRange(new DevComponents.DotNetBar.BaseItem[] {
            this.buttonItem2});
            this.btnAddTempStudent.TabIndex = 2;
            this.btnAddTempStudent.Text = "加入待處理學生";
            this.btnAddTempStudent.PopupOpen += new System.EventHandler(this.btnAddTempStudent_PopupOpen);
            this.btnAddTempStudent.Click += new System.EventHandler(this.btnAddTempStudent_Click);
            // 
            // buttonItem2
            // 
            this.buttonItem2.Name = "buttonItem2";
            this.buttonItem2.Text = "New Item";
            // 
            // lblStudentCount
            // 
            // 
            // 
            // 
            this.lblStudentCount.BackgroundStyle.Class = "";
            this.lblStudentCount.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblStudentCount.Location = new System.Drawing.Point(269, 207);
            this.lblStudentCount.Name = "lblStudentCount";
            this.lblStudentCount.Size = new System.Drawing.Size(266, 23);
            this.lblStudentCount.TabIndex = 7;
            // 
            // CounselStudentListAContent
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.Controls.Add(this.lblStudentCount);
            this.Controls.Add(this.btnAddTempStudent);
            this.Controls.Add(this.btnRemoveStudent);
            this.Controls.Add(this.lvStudentList);
            this.Name = "CounselStudentListAContent";
            this.Size = new System.Drawing.Size(550, 240);
            this.Load += new System.EventHandler(this.CounselStudentListAContent_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private  DevComponents.DotNetBar.Controls.ListViewEx lvStudentList;
        private System.Windows.Forms.ColumnHeader colClassName;
        private System.Windows.Forms.ColumnHeader colSeatNo;
        private System.Windows.Forms.ColumnHeader colName;
        private System.Windows.Forms.ColumnHeader colStudentNumber;
        private System.Windows.Forms.ColumnHeader colStatus;
        private DevComponents.DotNetBar.ButtonX btnRemoveStudent;
        private DevComponents.DotNetBar.ButtonX btnAddTempStudent;
        private DevComponents.DotNetBar.ButtonItem buttonItem2;
        private DevComponents.DotNetBar.LabelX lblStudentCount;
    }
}
