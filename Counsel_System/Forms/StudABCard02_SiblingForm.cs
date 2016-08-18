using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Counsel_System2.DAO;

namespace Counsel_System2.Forms
{
    public partial class StudABCard02_SiblingForm : FISCA.Presentation.Controls.BaseForm
    {
        UDTSiblingDef _data;        
        public StudABCard02_SiblingForm(UDTSiblingDef data)
        {
            InitializeComponent();
            _data = data;
            LoadData();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.OK;

            SetData();

            List<UDTSiblingDef> dataList = new List<UDTSiblingDef>();
            dataList.Add(_data);
            if (string.IsNullOrEmpty(_data.UID))
                UDTTransfer.ABUDTSiblingInsert(dataList);
            else
                UDTTransfer.ABUDTSiblingUpdate(dataList);
                        
            this.Close();

        }

        private void LoadData()
        {
            this.Text+="-"+ _data.Title;
            txtName.Text = _data.Name;
            txtSchoolName.Text = _data.SchoolName;
            if (_data.BirthYear.HasValue)
                txtBirthYear.Text = _data.BirthYear.Value.ToString();
            else
                txtBirthYear.Text = "";
            txtRemark.Text = _data.Remark;
        }

        private void SetData()
        {         
            _data.Name = txtName.Text;
            _data.SchoolName = txtSchoolName.Text;
            int y;
            if (int.TryParse(txtBirthYear.Text, out y))
                _data.BirthYear = y;
            else
                _data.BirthYear = null;
            txtRemark.Text = _data.Remark;

        
        }

        private void StudABCard02_SiblingForm_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;
        }


    }
}
