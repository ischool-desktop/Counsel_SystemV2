namespace Counsel_System2.Contents
{
    partial class StudInterviewDataContent
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
            this.btnAdd = new DevComponents.DotNetBar.ButtonX();
            this.btnEdit = new DevComponents.DotNetBar.ButtonX();
            this.btnDel = new DevComponents.DotNetBar.ButtonX();
            this.lvInterview = new DevComponents.DotNetBar.Controls.ListViewEx();
            this.colInterviewDate = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colInterviewType = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colTeacherID = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.ColCause = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.SuspendLayout();
            // 
            // btnAdd
            // 
            this.btnAdd.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnAdd.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnAdd.Location = new System.Drawing.Point(15, 205);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 25);
            this.btnAdd.TabIndex = 1;
            this.btnAdd.Text = "新增";
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
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
            // btnDel
            // 
            this.btnDel.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnDel.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnDel.Location = new System.Drawing.Point(185, 205);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(75, 25);
            this.btnDel.TabIndex = 3;
            this.btnDel.Text = "刪除";
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // lvInterview
            // 
            // 
            // 
            // 
            this.lvInterview.Border.Class = "ListViewBorder";
            this.lvInterview.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lvInterview.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colInterviewDate,
            this.colInterviewType,
            this.colTeacherID,
            this.ColCause});
            this.lvInterview.FullRowSelect = true;
            this.lvInterview.Location = new System.Drawing.Point(15, 12);
            this.lvInterview.MultiSelect = false;
            this.lvInterview.Name = "lvInterview";
            this.lvInterview.Size = new System.Drawing.Size(520, 186);
            this.lvInterview.TabIndex = 4;
            this.lvInterview.UseCompatibleStateImageBehavior = false;
            this.lvInterview.View = System.Windows.Forms.View.Details;
            // 
            // colInterviewDate
            // 
            this.colInterviewDate.Text = "晤談日期";
            this.colInterviewDate.Width = 100;
            // 
            // colInterviewType
            // 
            this.colInterviewType.Text = "晤談對象";
            this.colInterviewType.Width = 100;
            // 
            // colTeacherID
            // 
            this.colTeacherID.Text = "晤談老師";
            this.colTeacherID.Width = 100;
            // 
            // ColCause
            // 
            this.ColCause.Text = "晤談事由";
            this.ColCause.Width = 100;
            // 
            // StudInterviewDataContent
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.Controls.Add(this.lvInterview);
            this.Controls.Add(this.btnDel);
            this.Controls.Add(this.btnEdit);
            this.Controls.Add(this.btnAdd);
            this.Name = "StudInterviewDataContent";
            this.Size = new System.Drawing.Size(550, 240);
            this.Load += new System.EventHandler(this.StudInterviewDataContent_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.ButtonX btnAdd;
        private DevComponents.DotNetBar.ButtonX btnEdit;
        private DevComponents.DotNetBar.ButtonX btnDel;
        private  DevComponents.DotNetBar.Controls.ListViewEx lvInterview;
        private System.Windows.Forms.ColumnHeader colInterviewDate;
        private System.Windows.Forms.ColumnHeader colInterviewType;
        private System.Windows.Forms.ColumnHeader colTeacherID;
        private System.Windows.Forms.ColumnHeader ColCause;
    }
}
