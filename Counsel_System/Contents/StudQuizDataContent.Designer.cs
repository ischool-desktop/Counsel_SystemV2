namespace Counsel_System2.Contents
{
    partial class StudQuizDataContent
    {
        /// <summary> 
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 元件設計工具產生的程式碼

        /// <summary> 
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器
        /// 修改這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btnEdit = new DevComponents.DotNetBar.ButtonX();
            this.lvStudQuizData = new DevComponents.DotNetBar.Controls.ListViewEx();
            this.colQuizName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colImplementationDate = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colAnalysisDate = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.btnAdd = new DevComponents.DotNetBar.ButtonX();
            this.btnDel = new DevComponents.DotNetBar.ButtonX();
            this.SuspendLayout();
            // 
            // btnEdit
            // 
            this.btnEdit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnEdit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnEdit.Location = new System.Drawing.Point(100, 205);
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.Size = new System.Drawing.Size(75, 25);
            this.btnEdit.TabIndex = 1;
            this.btnEdit.Text = "修改";
            this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
            // 
            // lvStudQuizData
            // 
            // 
            // 
            // 
            this.lvStudQuizData.Border.Class = "ListViewBorder";
            this.lvStudQuizData.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lvStudQuizData.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colQuizName,
            this.colImplementationDate,
            this.colAnalysisDate});
            this.lvStudQuizData.FullRowSelect = true;
            this.lvStudQuizData.Location = new System.Drawing.Point(15, 12);
            this.lvStudQuizData.MultiSelect = false;
            this.lvStudQuizData.Name = "lvStudQuizData";
            this.lvStudQuizData.Size = new System.Drawing.Size(492, 186);
            this.lvStudQuizData.TabIndex = 2;
            this.lvStudQuizData.UseCompatibleStateImageBehavior = false;
            this.lvStudQuizData.View = System.Windows.Forms.View.Details;
            this.lvStudQuizData.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.lvStudQuizData_MouseDoubleClick);
            // 
            // colQuizName
            // 
            this.colQuizName.Text = "測驗名稱";
            this.colQuizName.Width = 200;
            // 
            // colImplementationDate
            // 
            this.colImplementationDate.Text = "實施日期";
            this.colImplementationDate.Width = 120;
            // 
            // colAnalysisDate
            // 
            this.colAnalysisDate.Text = "解析日期";
            this.colAnalysisDate.Width = 120;
            // 
            // btnAdd
            // 
            this.btnAdd.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnAdd.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnAdd.Location = new System.Drawing.Point(15, 205);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 25);
            this.btnAdd.TabIndex = 3;
            this.btnAdd.Text = "新增";
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnDel
            // 
            this.btnDel.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnDel.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnDel.Location = new System.Drawing.Point(185, 205);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(75, 25);
            this.btnDel.TabIndex = 4;
            this.btnDel.Text = "刪除";
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // StudQuizDataContent
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btnDel);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.lvStudQuizData);
            this.Controls.Add(this.btnEdit);
            this.Font = new System.Drawing.Font("Microsoft JhengHei", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "StudQuizDataContent";
            this.Size = new System.Drawing.Size(550, 240);
            this.Load += new System.EventHandler(this.StudQuizDataContent_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.ButtonX btnEdit;
        private  DevComponents.DotNetBar.Controls.ListViewEx lvStudQuizData;
        private System.Windows.Forms.ColumnHeader colQuizName;
        private System.Windows.Forms.ColumnHeader colImplementationDate;
        private System.Windows.Forms.ColumnHeader colAnalysisDate;
        private DevComponents.DotNetBar.ButtonX btnAdd;
        private DevComponents.DotNetBar.ButtonX btnDel;
    }
}
