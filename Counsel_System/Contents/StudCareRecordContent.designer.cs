namespace Counsel_System2.Contents
{
    partial class StudCareRecordContent
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
            this.lvCareRecord = new DevComponents.DotNetBar.Controls.ListViewEx();
            this.colEditor = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colFileDate = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colCaseCategory = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.btnInsert = new DevComponents.DotNetBar.ButtonX();
            this.btnEdit = new DevComponents.DotNetBar.ButtonX();
            this.btnDelete = new DevComponents.DotNetBar.ButtonX();
            this.SuspendLayout();
            // 
            // lvCareRecord
            // 
            // 
            // 
            // 
            this.lvCareRecord.Border.Class = "ListViewBorder";
            this.lvCareRecord.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lvCareRecord.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colEditor,
            this.colFileDate,
            this.colCaseCategory});
            this.lvCareRecord.FullRowSelect = true;
            this.lvCareRecord.Location = new System.Drawing.Point(15, 12);
            this.lvCareRecord.MultiSelect = false;
            this.lvCareRecord.Name = "lvCareRecord";
            this.lvCareRecord.Size = new System.Drawing.Size(520, 186);
            this.lvCareRecord.TabIndex = 0;
            this.lvCareRecord.UseCompatibleStateImageBehavior = false;
            this.lvCareRecord.View = System.Windows.Forms.View.Details;
            // 
            // colEditor
            // 
            this.colEditor.Text = "記錄者";
            this.colEditor.Width = 100;
            // 
            // colFileDate
            // 
            this.colFileDate.Text = "立案日期";
            this.colFileDate.Width = 100;
            // 
            // colCaseCategory
            // 
            this.colCaseCategory.Text = "個案類別";
            this.colCaseCategory.Width = 200;
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
            // StudCareRecordContent
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btnDelete);
            this.Controls.Add(this.btnEdit);
            this.Controls.Add(this.btnInsert);
            this.Controls.Add(this.lvCareRecord);
            this.Name = "StudCareRecordContent";
            this.Size = new System.Drawing.Size(550, 240);
            this.Load += new System.EventHandler(this.StudCareRecordContent_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private  DevComponents.DotNetBar.Controls.ListViewEx lvCareRecord;
        private System.Windows.Forms.ColumnHeader colEditor;
        private System.Windows.Forms.ColumnHeader colFileDate;
        private System.Windows.Forms.ColumnHeader colCaseCategory;
        private DevComponents.DotNetBar.ButtonX btnInsert;
        private DevComponents.DotNetBar.ButtonX btnEdit;
        private DevComponents.DotNetBar.ButtonX btnDelete;
    }
}
