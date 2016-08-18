namespace Counsel_System2.Forms
{
    partial class SetABCardTemplateForm
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
            this.dgDataField = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.btnAdd = new DevComponents.DotNetBar.ButtonX();
            this.btnSave = new DevComponents.DotNetBar.ButtonX();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.txtSubjectName = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.panelEx1 = new DevComponents.DotNetBar.PanelEx();
            this.itmPnlSubjectName = new DevComponents.DotNetBar.ItemPanel();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.btnDel = new DevComponents.DotNetBar.ButtonX();
            this.colQGLabel = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colQqLabel = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colQqType = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.colQqDetail = new System.Windows.Forms.DataGridViewButtonColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dgDataField)).BeginInit();
            this.panelEx1.SuspendLayout();
            this.SuspendLayout();
            // 
            // dgDataField
            // 
            this.dgDataField.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgDataField.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colQGLabel,
            this.colQqLabel,
            this.colQqType,
            this.colQqDetail});
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgDataField.DefaultCellStyle = dataGridViewCellStyle1;
            this.dgDataField.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.dgDataField.Location = new System.Drawing.Point(15, 45);
            this.dgDataField.Name = "dgDataField";
            this.dgDataField.RowTemplate.Height = 24;
            this.dgDataField.Size = new System.Drawing.Size(547, 193);
            this.dgDataField.TabIndex = 7;
            this.dgDataField.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgDataField_CellContentClick);
            // 
            // btnAdd
            // 
            this.btnAdd.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnAdd.BackColor = System.Drawing.Color.Transparent;
            this.btnAdd.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnAdd.Location = new System.Drawing.Point(12, 290);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(55, 23);
            this.btnAdd.TabIndex = 8;
            this.btnAdd.Text = "新增";
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnSave
            // 
            this.btnSave.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnSave.BackColor = System.Drawing.Color.Transparent;
            this.btnSave.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnSave.Location = new System.Drawing.Point(507, 12);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(55, 23);
            this.btnSave.TabIndex = 6;
            this.btnSave.Text = "儲存";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // labelX1
            // 
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.Class = "";
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.Location = new System.Drawing.Point(19, 16);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(32, 23);
            this.labelX1.TabIndex = 2;
            this.labelX1.Text = "名稱";
            // 
            // txtSubjectName
            // 
            // 
            // 
            // 
            this.txtSubjectName.Border.Class = "TextBoxBorder";
            this.txtSubjectName.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.txtSubjectName.Location = new System.Drawing.Point(56, 12);
            this.txtSubjectName.Name = "txtSubjectName";
            this.txtSubjectName.Size = new System.Drawing.Size(421, 25);
            this.txtSubjectName.TabIndex = 0;
            // 
            // panelEx1
            // 
            this.panelEx1.CanvasColor = System.Drawing.SystemColors.Control;
            this.panelEx1.Controls.Add(this.dgDataField);
            this.panelEx1.Controls.Add(this.btnSave);
            this.panelEx1.Controls.Add(this.labelX1);
            this.panelEx1.Controls.Add(this.txtSubjectName);
            this.panelEx1.Location = new System.Drawing.Point(135, 31);
            this.panelEx1.Name = "panelEx1";
            this.panelEx1.Size = new System.Drawing.Size(576, 253);
            this.panelEx1.Style.Alignment = System.Drawing.StringAlignment.Center;
            this.panelEx1.Style.BackColor1.Color = System.Drawing.Color.White;
            this.panelEx1.Style.BackColor2.Color = System.Drawing.Color.White;
            this.panelEx1.Style.Border = DevComponents.DotNetBar.eBorderType.SingleLine;
            this.panelEx1.Style.BorderColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBorder;
            this.panelEx1.Style.ForeColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelText;
            this.panelEx1.Style.GradientAngle = 90;
            this.panelEx1.TabIndex = 12;
            this.panelEx1.Text = "panelEx1";
            // 
            // itmPnlSubjectName
            // 
            this.itmPnlSubjectName.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.itmPnlSubjectName.BackgroundStyle.BackColor = System.Drawing.Color.White;
            this.itmPnlSubjectName.BackgroundStyle.BorderBottom = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.itmPnlSubjectName.BackgroundStyle.BorderBottomWidth = 1;
            this.itmPnlSubjectName.BackgroundStyle.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(127)))), ((int)(((byte)(157)))), ((int)(((byte)(185)))));
            this.itmPnlSubjectName.BackgroundStyle.BorderLeft = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.itmPnlSubjectName.BackgroundStyle.BorderLeftWidth = 1;
            this.itmPnlSubjectName.BackgroundStyle.BorderRight = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.itmPnlSubjectName.BackgroundStyle.BorderRightWidth = 1;
            this.itmPnlSubjectName.BackgroundStyle.BorderTop = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.itmPnlSubjectName.BackgroundStyle.BorderTopWidth = 1;
            this.itmPnlSubjectName.BackgroundStyle.Class = "";
            this.itmPnlSubjectName.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.itmPnlSubjectName.BackgroundStyle.PaddingBottom = 1;
            this.itmPnlSubjectName.BackgroundStyle.PaddingLeft = 1;
            this.itmPnlSubjectName.BackgroundStyle.PaddingRight = 1;
            this.itmPnlSubjectName.BackgroundStyle.PaddingTop = 1;
            this.itmPnlSubjectName.ContainerControlProcessDialogKey = true;
            this.itmPnlSubjectName.LayoutOrientation = DevComponents.DotNetBar.eOrientation.Vertical;
            this.itmPnlSubjectName.LicenseKey = "F962CEC7-CD8F-4911-A9E9-CAB39962FC1F";
            this.itmPnlSubjectName.Location = new System.Drawing.Point(12, 31);
            this.itmPnlSubjectName.Name = "itmPnlSubjectName";
            this.itmPnlSubjectName.Size = new System.Drawing.Size(116, 253);
            this.itmPnlSubjectName.TabIndex = 11;
            this.itmPnlSubjectName.Text = "itemPanel1";
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Location = new System.Drawing.Point(642, 290);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(55, 23);
            this.btnExit.TabIndex = 10;
            this.btnExit.Text = "離開";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // btnDel
            // 
            this.btnDel.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnDel.BackColor = System.Drawing.Color.Transparent;
            this.btnDel.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnDel.Location = new System.Drawing.Point(74, 290);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(55, 23);
            this.btnDel.TabIndex = 9;
            this.btnDel.Text = "刪除";
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // colQGLabel
            // 
            this.colQGLabel.HeaderText = "題目群組名稱";
            this.colQGLabel.Name = "colQGLabel";
            // 
            // colQqLabel
            // 
            this.colQqLabel.HeaderText = "題目名稱";
            this.colQqLabel.Name = "colQqLabel";
            this.colQqLabel.Width = 200;
            // 
            // colQqType
            // 
            this.colQqType.HeaderText = "題目類型";
            this.colQqType.Name = "colQqType";
            this.colQqType.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            // 
            // colQqDetail
            // 
            this.colQqDetail.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells;
            this.colQqDetail.FillWeight = 50F;
            this.colQqDetail.HeaderText = "題目內容";
            this.colQqDetail.Name = "colQqDetail";
            this.colQqDetail.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.colQqDetail.Text = "...";
            this.colQqDetail.Width = 48;
            // 
            // SetABCardTemplateForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(722, 319);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.panelEx1);
            this.Controls.Add(this.itmPnlSubjectName);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnDel);
            this.DoubleBuffered = true;
            this.Name = "SetABCardTemplateForm";
            this.Text = "設定綜合表現樣板";
            this.Load += new System.EventHandler(this.SetABCardTemplateForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgDataField)).EndInit();
            this.panelEx1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.Controls.DataGridViewX dgDataField;
        private DevComponents.DotNetBar.ButtonX btnAdd;
        private DevComponents.DotNetBar.ButtonX btnSave;
        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.Controls.TextBoxX txtSubjectName;
        private DevComponents.DotNetBar.PanelEx panelEx1;
        private DevComponents.DotNetBar.ItemPanel itmPnlSubjectName;
        private DevComponents.DotNetBar.ButtonX btnExit;
        private DevComponents.DotNetBar.ButtonX btnDel;
        private System.Windows.Forms.DataGridViewTextBoxColumn colQGLabel;
        private System.Windows.Forms.DataGridViewTextBoxColumn colQqLabel;
        private System.Windows.Forms.DataGridViewComboBoxColumn colQqType;
        private System.Windows.Forms.DataGridViewButtonColumn colQqDetail;
    }
}