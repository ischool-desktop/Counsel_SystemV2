namespace Counsel_System2.Forms
{
    partial class StudABCard_PriorityForm_Combo
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
            this.cbx01 = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.lbl01 = new DevComponents.DotNetBar.LabelX();
            this.lbl02 = new DevComponents.DotNetBar.LabelX();
            this.lbl03 = new DevComponents.DotNetBar.LabelX();
            this.cbx02 = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.cbx03 = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.btnSave = new DevComponents.DotNetBar.ButtonX();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.SuspendLayout();
            // 
            // cbx01
            // 
            this.cbx01.DisplayMember = "Text";
            this.cbx01.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cbx01.FormattingEnabled = true;
            this.cbx01.ItemHeight = 19;
            this.cbx01.Location = new System.Drawing.Point(74, 11);
            this.cbx01.Name = "cbx01";
            this.cbx01.Size = new System.Drawing.Size(173, 25);
            this.cbx01.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbx01.TabIndex = 1;
            // 
            // lbl01
            // 
            this.lbl01.AutoSize = true;
            this.lbl01.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lbl01.BackgroundStyle.Class = "";
            this.lbl01.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbl01.Location = new System.Drawing.Point(23, 11);
            this.lbl01.Name = "lbl01";
            this.lbl01.Size = new System.Drawing.Size(45, 21);
            this.lbl01.TabIndex = 1;
            this.lbl01.Text = "意願 1";
            // 
            // lbl02
            // 
            this.lbl02.AutoSize = true;
            this.lbl02.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lbl02.BackgroundStyle.Class = "";
            this.lbl02.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbl02.Location = new System.Drawing.Point(23, 48);
            this.lbl02.Name = "lbl02";
            this.lbl02.Size = new System.Drawing.Size(45, 21);
            this.lbl02.TabIndex = 2;
            this.lbl02.Text = "意願 2";
            // 
            // lbl03
            // 
            this.lbl03.AutoSize = true;
            this.lbl03.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lbl03.BackgroundStyle.Class = "";
            this.lbl03.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbl03.Location = new System.Drawing.Point(23, 85);
            this.lbl03.Name = "lbl03";
            this.lbl03.Size = new System.Drawing.Size(45, 21);
            this.lbl03.TabIndex = 3;
            this.lbl03.Text = "意願 3";
            // 
            // cbx02
            // 
            this.cbx02.DisplayMember = "Text";
            this.cbx02.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cbx02.FormattingEnabled = true;
            this.cbx02.ItemHeight = 19;
            this.cbx02.Location = new System.Drawing.Point(74, 48);
            this.cbx02.Name = "cbx02";
            this.cbx02.Size = new System.Drawing.Size(173, 25);
            this.cbx02.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbx02.TabIndex = 2;
            // 
            // cbx03
            // 
            this.cbx03.DisplayMember = "Text";
            this.cbx03.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cbx03.FormattingEnabled = true;
            this.cbx03.ItemHeight = 19;
            this.cbx03.Location = new System.Drawing.Point(74, 85);
            this.cbx03.Name = "cbx03";
            this.cbx03.Size = new System.Drawing.Size(173, 25);
            this.cbx03.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbx03.TabIndex = 3;
            // 
            // btnSave
            // 
            this.btnSave.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnSave.BackColor = System.Drawing.Color.Transparent;
            this.btnSave.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnSave.Location = new System.Drawing.Point(91, 131);
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
            this.btnExit.Location = new System.Drawing.Point(172, 131);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 25);
            this.btnExit.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnExit.TabIndex = 5;
            this.btnExit.Text = "離開";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // StudABCard_PriorityForm_Combo
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(271, 166);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.cbx03);
            this.Controls.Add(this.cbx02);
            this.Controls.Add(this.lbl03);
            this.Controls.Add(this.lbl02);
            this.Controls.Add(this.lbl01);
            this.Controls.Add(this.cbx01);
            this.DoubleBuffered = true;
            this.Name = "StudABCard_PriorityForm_Combo";
            this.Text = "StudABCard_PriorityForm_Combo";
            this.Load += new System.EventHandler(this.StudABCard_PriorityForm_Combo_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.Controls.ComboBoxEx cbx01;
        private DevComponents.DotNetBar.LabelX lbl01;
        private DevComponents.DotNetBar.LabelX lbl02;
        private DevComponents.DotNetBar.LabelX lbl03;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cbx02;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cbx03;
        private DevComponents.DotNetBar.ButtonX btnSave;
        private DevComponents.DotNetBar.ButtonX btnExit;
    }
}