namespace Counsel_System2.Forms
{
    partial class SetCounselTeacherForm
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
            this.lvTeacherName = new  DevComponents.DotNetBar.Controls.ListViewEx();
            this.colTeacherName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.btnOk = new DevComponents.DotNetBar.ButtonX();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.SuspendLayout();
            // 
            // lvTeacherName
            // 
            // 
            // 
            // 
            this.lvTeacherName.Border.Class = "ListViewBorder";
            this.lvTeacherName.CheckBoxes = true;
            this.lvTeacherName.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colTeacherName});
            this.lvTeacherName.Location = new System.Drawing.Point(13, 13);
            this.lvTeacherName.Name = "lvTeacherName";
            this.lvTeacherName.Size = new System.Drawing.Size(207, 150);
            this.lvTeacherName.TabIndex = 0;
            this.lvTeacherName.UseCompatibleStateImageBehavior = false;
            this.lvTeacherName.View = System.Windows.Forms.View.Details;
            // 
            // colTeacherName
            // 
            this.colTeacherName.Text = "教師姓名";
            this.colTeacherName.Width = 200;
            // 
            // btnOk
            // 
            this.btnOk.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnOk.BackColor = System.Drawing.Color.Transparent;
            this.btnOk.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnOk.Location = new System.Drawing.Point(75, 178);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(65, 23);
            this.btnOk.TabIndex = 1;
            this.btnOk.Text = "確定";
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Location = new System.Drawing.Point(154, 178);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(65, 23);
            this.btnExit.TabIndex = 2;
            this.btnExit.Text = "離開";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // SetCounselTeacherForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(229, 209);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.lvTeacherName);
            this.DoubleBuffered = true;
            this.Name = "SetCounselTeacherForm";
            this.Text = "設定";
            this.Load += new System.EventHandler(this.SetCounselTeacherForm_Load);
            this.ResumeLayout(false);

        }

        #endregion
        
        private  DevComponents.DotNetBar.Controls.ListViewEx lvTeacherName;
        private System.Windows.Forms.ColumnHeader colTeacherName;
        private DevComponents.DotNetBar.ButtonX btnOk;
        private DevComponents.DotNetBar.ButtonX btnExit;
    }
}