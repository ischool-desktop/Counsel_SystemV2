namespace Counsel_System2.Contents
{
    partial class CounselStudentListContent
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
            this.colTeacherName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colTeacherType = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.btnSetCounselTeacherA = new DevComponents.DotNetBar.ButtonX();
            this.btnSetCounselTeacherB = new DevComponents.DotNetBar.ButtonX();
            this.btnRemoveCounselTeacher = new DevComponents.DotNetBar.ButtonX();
            this.lblTeacherCount = new DevComponents.DotNetBar.LabelX();
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
            this.colTeacherName,
            this.colTeacherType});
            this.lvStudentList.FullRowSelect = true;
            this.lvStudentList.Location = new System.Drawing.Point(15, 12);
            this.lvStudentList.Name = "lvStudentList";
            this.lvStudentList.Size = new System.Drawing.Size(520, 186);
            this.lvStudentList.TabIndex = 0;
            this.lvStudentList.UseCompatibleStateImageBehavior = false;
            this.lvStudentList.View = System.Windows.Forms.View.Details;
            // 
            // colTeacherName
            // 
            this.colTeacherName.Text = "教師姓名";
            this.colTeacherName.Width = 150;
            // 
            // colTeacherType
            // 
            this.colTeacherType.Text = "教師類別";
            this.colTeacherType.Width = 100;
            // 
            // btnSetCounselTeacherA
            // 
            this.btnSetCounselTeacherA.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnSetCounselTeacherA.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnSetCounselTeacherA.Location = new System.Drawing.Point(15, 205);
            this.btnSetCounselTeacherA.Name = "btnSetCounselTeacherA";
            this.btnSetCounselTeacherA.Size = new System.Drawing.Size(97, 23);
            this.btnSetCounselTeacherA.TabIndex = 1;
            this.btnSetCounselTeacherA.Text = "指定認輔老師";
            this.btnSetCounselTeacherA.Click += new System.EventHandler(this.btnSetCounselTeacherA_Click);
            // 
            // btnSetCounselTeacherB
            // 
            this.btnSetCounselTeacherB.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnSetCounselTeacherB.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnSetCounselTeacherB.Location = new System.Drawing.Point(118, 205);
            this.btnSetCounselTeacherB.Name = "btnSetCounselTeacherB";
            this.btnSetCounselTeacherB.Size = new System.Drawing.Size(97, 23);
            this.btnSetCounselTeacherB.TabIndex = 2;
            this.btnSetCounselTeacherB.Text = "指定輔導老師";
            this.btnSetCounselTeacherB.Click += new System.EventHandler(this.btnSetCounselTeacherB_Click);
            // 
            // btnRemoveCounselTeacher
            // 
            this.btnRemoveCounselTeacher.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnRemoveCounselTeacher.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnRemoveCounselTeacher.Location = new System.Drawing.Point(223, 205);
            this.btnRemoveCounselTeacher.Name = "btnRemoveCounselTeacher";
            this.btnRemoveCounselTeacher.Size = new System.Drawing.Size(97, 23);
            this.btnRemoveCounselTeacher.TabIndex = 3;
            this.btnRemoveCounselTeacher.Text = "移除所選老師";
            this.btnRemoveCounselTeacher.Click += new System.EventHandler(this.btnRemoveCounselTeacher_Click);
            // 
            // lblTeacherCount
            // 
            // 
            // 
            // 
            this.lblTeacherCount.BackgroundStyle.Class = "";
            this.lblTeacherCount.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblTeacherCount.Location = new System.Drawing.Point(326, 205);
            this.lblTeacherCount.Name = "lblTeacherCount";
            this.lblTeacherCount.Size = new System.Drawing.Size(209, 23);
            this.lblTeacherCount.TabIndex = 4;
            // 
            // CounselStudentListContent
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.lblTeacherCount);
            this.Controls.Add(this.btnRemoveCounselTeacher);
            this.Controls.Add(this.btnSetCounselTeacherB);
            this.Controls.Add(this.btnSetCounselTeacherA);
            this.Controls.Add(this.lvStudentList);
            this.Name = "CounselStudentListContent";
            this.Size = new System.Drawing.Size(550, 240);
            this.ResumeLayout(false);

        }

        #endregion

        private  DevComponents.DotNetBar.Controls.ListViewEx lvStudentList;
        private System.Windows.Forms.ColumnHeader colTeacherName;
        private System.Windows.Forms.ColumnHeader colTeacherType;
        private DevComponents.DotNetBar.ButtonX btnSetCounselTeacherA;
        private DevComponents.DotNetBar.ButtonX btnSetCounselTeacherB;
        private DevComponents.DotNetBar.ButtonX btnRemoveCounselTeacher;
        private DevComponents.DotNetBar.LabelX lblTeacherCount;
    }
}
