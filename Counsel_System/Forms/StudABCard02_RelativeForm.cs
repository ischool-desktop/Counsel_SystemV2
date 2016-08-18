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
    public partial class StudABCard02_RelativeForm : FISCA.Presentation.Controls.BaseForm
    {
        UDTRelativeDef _data;

        public StudABCard02_RelativeForm(UDTRelativeDef data)
        {
            InitializeComponent();
            _data = data;
        }
        
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            SetData();

            this.DialogResult = System.Windows.Forms.DialogResult.OK;

            List<UDTRelativeDef> dataList = new List<UDTRelativeDef>();
            dataList.Add(_data);

            if (string.IsNullOrEmpty(_data.UID))
                UDTTransfer.ABUDTRelativeInsert(dataList);
            else
                UDTTransfer.ABUDTRelativeUpdate(dataList);
                        
            this.Close();
        }

        private void StudABCard02_RelativeForm_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;
            LoadData();
        }

        private void SetData()
        {            
            _data.Name = txtName.Text;

            if (cbxAlive.Text == "存")
                _data.IsAlive = true;
            else if (cbxAlive.Text == "沒")
                _data.IsAlive = false;
            else
                _data.IsAlive = null;
                
            int y;
            if (int.TryParse(txtBirthYear.Text, out y))
                _data.BirthYear = y;
            else
                _data.BirthYear = null;
            _data.Job = cbxJob.Text;
            _data.Institute=txtInstiute.Text;
            _data.JobTitle = txtJobTitle.Text;
            _data.EduDegree = cbxEduDegree.Text;
        }

        private void LoadData()
        {
            this.Text += "-" + _data.Title;
            txtName.Text = _data.Name;
            cbxAlive.Text = "";
            if (_data.IsAlive.HasValue)
            {
                if (_data.IsAlive.Value)
                    cbxAlive.Text = "存";
                else
                    cbxAlive.Text = "沒";
            }
            if (_data.BirthYear.HasValue)
                txtBirthYear.Text = _data.BirthYear.Value.ToString();
            else
                txtBirthYear.Text = "";

            cbxJob.Text = _data.Job;
            txtInstiute.Text = _data.Institute;
            txtJobTitle.Text = _data.JobTitle;
            cbxEduDegree.Text = _data.EduDegree;
            cbxAlive.Items.Add("存");
            cbxAlive.Items.Add("沒");
        }

        public void SetJobSelectItems(List<string> items)
        {
            cbxJob.Items.Clear();
            cbxJob.Items.AddRange(items.ToArray());
        }

        public void SetEduDegreeSelectItems(List<string> items)
        {
            cbxEduDegree.Items.Clear();
            cbxEduDegree.Items.AddRange(items.ToArray());
        }
    }
}
